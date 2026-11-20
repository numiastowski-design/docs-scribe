/**
 * Four names of the published reading door (`docs-scribe`) that no other test of the package takes:
 * `assertZipDeclaredSizesSafe`, `diagramTextWords`, `undrawnDiagramWords` and `detectLanguage`, called as a client
 * calls them — on an archive written byte by byte here, and on document trees written by hand.
 * NOT here: the five readers of the door (each has its own tests), nor what the publisher's application makes of a
 * refusal.
 */
import { describe, expect, it } from "vitest";

import { assertZipDeclaredSizesSafe, detectLanguage, diagramTextWords, undrawnDiagramWords } from "./docs";
import { ReaderError, type IRBlock, type IRDoc } from "./docs-light";

const u16 = (value: number): number[] => [value & 0xff, (value >>> 8) & 0xff];
const u32 = (value: number): number[] => [
  value & 0xff,
  (value >>> 8) & 0xff,
  (value >>> 16) & 0xff,
  (value >>> 24) & 0xff,
];

interface ZipEntry {
  name: string;
  /** The bytes written after the local header. */
  data: readonly number[];
  /** 0 stored, 8 deflated. */
  method: 0 | 8;
  /** The uncompressed size both headers DECLARE; nothing is ever inflated to check it. */
  declaredSize: number;
  /** Where the central record says the local header is, when it lies. */
  centralOffset?: number;
}

/** A zip archive written field by field: one local header and its data per entry, the central directory, the end record. */
const zipOf = (entries: readonly ZipEntry[]): Uint8Array => {
  const body: number[] = [];
  const central: number[] = [];
  for (const entry of entries) {
    const name = Array.from(new TextEncoder().encode(entry.name));
    const offset = body.length;
    const sizes = [...u32(0), ...u32(entry.data.length), ...u32(entry.declaredSize)];
    body.push(
      ...u32(0x04034b50),
      ...u16(20),
      ...u16(0),
      ...u16(entry.method),
      ...u16(0),
      ...u16(0),
      ...sizes,
      ...u16(name.length),
      ...u16(0),
      ...name,
      ...entry.data,
    );
    central.push(
      ...u32(0x02014b50),
      ...u16(20),
      ...u16(20),
      ...u16(0),
      ...u16(entry.method),
      ...u16(0),
      ...u16(0),
      ...sizes,
      ...u16(name.length),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(entry.centralOffset ?? offset),
      ...name,
    );
  }
  const end = [
    ...u32(0x06054b50),
    ...u16(0),
    ...u16(0),
    ...u16(entries.length),
    ...u16(entries.length),
    ...u32(central.length),
    ...u32(body.length),
    ...u16(0),
  ];
  return new Uint8Array([...body, ...central, ...end]);
};

/** A deflated entry of two bytes (an empty deflate stream) that DECLARES `declaredSize` bytes once inflated. */
const declaring = (name: string, declaredSize: number): ZipEntry => ({
  name,
  data: [0x03, 0x00],
  method: 8,
  declaredSize,
});

const refusalOf = (bytes: Uint8Array): { code: string; message: string } | "no refusal" => {
  try {
    assertZipDeclaredSizesSafe(bytes);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    const { code, message } = error as ReaderError;
    return { code, message };
  }
  return "no refusal";
};

describe("assertZipDeclaredSizesSafe — an archive is refused on the sizes it DECLARES, nothing inflated", () => {
  it("one part declaring 70 000 000 bytes, past the 64 000 000 of a part: FILE_UNSAFE", () => {
    expect(refusalOf(zipOf([declaring("word/document.xml", 70_000_000)]))).toEqual({
      code: "FILE_UNSAFE",
      message: "the file is refused for safety: a part declares a size past the per-part ceiling",
    });
  });

  it("the same part declaring exactly 64 000 000 bytes passes: the ceiling is the last size accepted", () => {
    expect(refusalOf(zipOf([declaring("word/document.xml", 64_000_000)]))).toBe("no refusal");
    expect(refusalOf(zipOf([declaring("word/document.xml", 64_000_001)]))).toMatchObject({ code: "FILE_UNSAFE" });
  });

  it("five parts of 60 000 000 bytes each, 300 000 000 in all, past the 256 000 000 of an archive: FILE_UNSAFE", () => {
    const parts = ["a", "b", "c", "d", "e"].map((name) => declaring(`xl/media/${name}.bin`, 60_000_000));
    expect(refusalOf(zipOf(parts))).toEqual({
      code: "FILE_UNSAFE",
      message: "the file is refused for safety: the parts declare a size past the package ceiling",
    });
    // Control: four of them, 240 000 000 bytes, are under the ceiling.
    expect(refusalOf(zipOf(parts.slice(0, 4)))).toBe("no refusal");
  });

  it("an entry whose header and data lie inside another entry's data (the overlapping-entries bomb): FILE_UNSAFE", () => {
    const hello = Array.from(new TextEncoder().encode("hello"));
    const inner: ZipEntry = { name: "b.xml", data: hello, method: 0, declaredSize: hello.length };
    // b.xml's own local header and data, quoted inside a.xml's stored data; b's central record points there.
    const quoted = Array.from(zipOf([inner]).subarray(0, 30 + "b.xml".length + hello.length));
    const bytes = zipOf([
      { name: "a.xml", data: quoted, method: 0, declaredSize: quoted.length },
      { ...inner, centralOffset: 30 + "a.xml".length },
    ]);
    expect(refusalOf(bytes)).toEqual({
      code: "FILE_UNSAFE",
      message: "the file is refused for safety: overlapping entries",
    });
    // Control: the same two entries one after the other, the central record telling the truth, pass.
    expect(refusalOf(zipOf([{ name: "a.xml", data: quoted, method: 0, declaredSize: quoted.length }, inner]))).toBe(
      "no refusal",
    );
  });

  it("a small honest archive passes, and so do bytes that are no zip at all (damaged is not hostile)", () => {
    expect(refusalOf(zipOf([declaring("word/document.xml", 10)]))).toBe("no refusal");
    expect(refusalOf(new Uint8Array([1, 2, 3]))).toBe("no refusal");
    expect(refusalOf(new Uint8Array())).toBe("no refusal");
  });
});

/** A document with a diagram published as a drawing: a tab that holds a paragraph and one section, then the sections
 *  gathered at the top level. The second top-level section carries `levels`: it is PUBLISHED, as a list. */
const DOCUMENT: IRBlock[] = [
  {
    kind: "tabs",
    tabs: [
      {
        title: "Slide 1",
        blocks: [
          { kind: "paragraph", inlines: [{ text: "Body text of the slide" }] },
          { kind: "diagramText", sections: [{ inlines: [{ text: "Inside a tab" }], undrawn: ["tab"] }] },
        ],
      },
    ],
  },
  {
    kind: "diagramText",
    sections: [
      {
        label: "Diagram of slide 2",
        inlines: [{ text: "Order " }, { text: "received" }, { text: "Invoice sent" }],
        breaks: [2],
        undrawn: ["invoice"],
      },
      { inlines: [{ text: "Published outline" }], levels: [0], undrawn: ["outline"] },
    ],
  },
];

describe("diagramTextWords — the words of the diagrams the page does not carry", () => {
  it("reads every section that stays off the page, inside a tab too, cut as the word check cuts", () => {
    // Not read: the paragraph (it is on the page), the label (our wording), the section that carries `levels`.
    expect(diagramTextWords(DOCUMENT)).toEqual(["inside", "a", "tab", "order", "received", "invoice", "sent"]);
  });

  it("a paragraph break ends a word: « received » and « Invoice » are not welded", () => {
    expect(diagramTextWords(DOCUMENT)).not.toContain("receivedinvoice");
    // Inside ONE paragraph the inlines are joined as they stand: « Order » + « received » keeps its own space.
    expect(
      diagramTextWords([{ kind: "diagramText", sections: [{ inlines: [{ text: "Or" }, { text: "der" }] }] }]),
    ).toEqual(["order"]);
  });

  it("a document with no diagram has none", () => {
    expect(diagramTextWords([{ kind: "paragraph", inlines: [{ text: "Only text" }] }])).toEqual([]);
    expect(diagramTextWords([])).toEqual([]);
  });
});

describe("detectLanguage — the language of a document, read from its own text", () => {
  const documentOf = (...paragraphs: string[]): IRDoc => ({
    blocks: paragraphs.map((text) => ({ kind: "paragraph", inlines: [{ text }] })),
  });
  const ENGLISH = [
    "The committee met on the first day of the month to review the work that had been done since the last report.",
    "It was agreed that the members of the team would have to send their notes to the office before the end of the week.",
    "There is no change in the way the requests are handled, and all of them will be answered in the order they were received.",
    "This document is one of the few that have been kept from that period, and it is not known who wrote it.",
  ];
  const FRENCH = [
    "Le comité s'est réuni le premier jour du mois pour examiner le travail qui avait été fait depuis le dernier rapport.",
    "Il a été convenu que les membres de l'équipe devraient envoyer leurs notes au bureau avant la fin de la semaine.",
    "Il n'y a pas de changement dans la manière dont les demandes sont traitées, et elles seront toutes prises dans l'ordre.",
    "Ce document est l'un des rares qui aient été conservés de cette période, et on ne sait pas qui l'a écrit.",
  ];

  it("English prose is « en », French prose is « fr »", () => {
    expect(detectLanguage(documentOf(...ENGLISH))).toBe("en");
    expect(detectLanguage(documentOf(...FRENCH))).toBe("fr");
  });

  it("a text too short to decide gives no language, never a default one", () => {
    expect(detectLanguage(documentOf("Invoice 2026-0042"))).toBeUndefined();
    expect(detectLanguage(documentOf())).toBeUndefined();
  });
});

describe("undrawnDiagramWords — the SmartArt words a picture does not show", () => {
  it("lists the undrawn words of every section that stays off the page, as the section gives them", () => {
    // The published section's « outline » is on the page as a list: a drawing that leaves a shape out loses nothing.
    expect(undrawnDiagramWords(DOCUMENT)).toEqual(["tab", "invoice"]);
  });

  it("a section that lists none, and a document with no diagram, give none", () => {
    expect(undrawnDiagramWords([{ kind: "diagramText", sections: [{ inlines: [{ text: "All drawn" }] }] }])).toEqual(
      [],
    );
    expect(undrawnDiagramWords([])).toEqual([]);
  });
});
