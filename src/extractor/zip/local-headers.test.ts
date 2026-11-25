/**
 * Slice 2 of the jszip detachment: each entry's local file header, read
 * without jszip and checked against its central directory record.
 *
 * Hand-built archives pin each rule (PKWARE APPNOTE 6.3.9 §4.3.7 local file
 * header, §4.3.9 data descriptor, §4.4.4 bit 3, §4.5.3 zip64 extended
 * information) and each hostile shape. jszip is the second, independent
 * reading: on every real Office fixture both readers must agree on names,
 * directory flags and the compressed bytes of every entry, or both refuse.
 */
import { readFileSync } from "node:fs";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { publicFixtures } from "../public-fixtures.test.support";
import { readLocalHeaders } from "./local-headers";

// ---------------------------------------------------------------------------
// A complete archive builder: local headers, data, optional data descriptors,
// central directory, end record. Every local field can be made to diverge.
// ---------------------------------------------------------------------------

const u16 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff];
const u32 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
const u64 = (v: number): number[] => {
  const out: number[] = [];
  let big = BigInt(v);
  for (let i = 0; i < 8; i++) {
    out.push(Number(big & 0xffn));
    big >>= 8n;
  }
  return out;
};
const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

interface Entry {
  name: string;
  /** Stored as-is: the reader under test never inflates. */
  data?: Uint8Array;
  method?: number;
  flags?: number;
  externalAttributes?: number;
  /** Local fields, when they must differ from the central record. */
  localName?: string;
  localMethod?: number;
  localFlags?: number;
  localCompressed?: number;
  localUncompressed?: number;
  /** Write CRC and sizes as zero locally, and a signed data descriptor after the data. */
  descriptor?: boolean;
  /** Saturate sizes and offset and move them to zip64 extra fields (§4.5.3). */
  zip64?: boolean;
  /** Saturate the central sizes but write no zip64 extra field. */
  saturatedWithoutExtra?: boolean;
  /** Overrides of the central record. */
  cdCompressed?: number;
  cdOffset?: number;
}

const SAT32 = 0xffffffff;

function build(entries: Entry[], opts: { gapBeforeCd?: number } = {}): Uint8Array {
  const body: number[] = [];
  const cd: number[] = [];
  for (const e of entries) {
    const data = e.data ?? new Uint8Array();
    const name = utf8(e.name);
    const localName = utf8(e.localName ?? e.name);
    const offset = body.length;
    const crc = 0x12345678;
    const descriptor = e.descriptor === true;
    const localFlags = e.localFlags ?? (e.flags ?? 0) | (descriptor ? 0x08 : 0);
    const flags = e.flags ?? (descriptor ? 0x08 : 0);
    let localCompressed = e.localCompressed ?? data.length;
    let localUncompressed = e.localUncompressed ?? data.length;
    if (descriptor) localCompressed = localUncompressed = 0;
    let localExtra: number[] = [];
    if (e.zip64) {
      localExtra = [...u16(1), ...u16(16), ...u64(localUncompressed), ...u64(localCompressed)];
      localCompressed = localUncompressed = SAT32;
    }
    body.push(
      ...u32(0x04034b50),
      ...u16(e.zip64 ? 45 : 20),
      ...u16(localFlags),
      ...u16(e.localMethod ?? e.method ?? 0),
      ...u16(0),
      ...u16(0),
      ...u32(descriptor ? 0 : crc),
      ...u32(localCompressed),
      ...u32(localUncompressed),
      ...u16(localName.length),
      ...u16(localExtra.length),
      ...localName,
      ...localExtra,
      ...data,
    );
    if (descriptor) body.push(...u32(0x08074b50), ...u32(crc), ...u32(data.length), ...u32(data.length));

    const cdCompressed = e.cdCompressed ?? data.length;
    const cdOffset = e.cdOffset ?? offset;
    const saturate = e.zip64 === true || e.saturatedWithoutExtra === true;
    const cdExtra = e.zip64 ? [...u16(1), ...u16(24), ...u64(data.length), ...u64(cdCompressed), ...u64(cdOffset)] : [];
    cd.push(
      ...u32(0x02014b50),
      ...u16(e.zip64 ? 45 : 20),
      ...u16(e.zip64 ? 45 : 20),
      ...u16(flags),
      ...u16(e.method ?? 0),
      ...u16(0),
      ...u16(0),
      ...u32(crc),
      ...u32(saturate ? SAT32 : cdCompressed),
      ...u32(saturate ? SAT32 : data.length),
      ...u16(name.length),
      ...u16(cdExtra.length),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(e.externalAttributes ?? 0),
      ...u32(saturate ? SAT32 : cdOffset),
      ...name,
      ...cdExtra,
    );
  }
  body.push(...new Array<number>(opts.gapBeforeCd ?? 0).fill(0));
  const cdAt = body.length;
  return Uint8Array.from([
    ...body,
    ...cd,
    ...u32(0x06054b50),
    ...u16(0),
    ...u16(0),
    ...u16(entries.length),
    ...u16(entries.length),
    ...u32(cd.length),
    ...u32(cdAt),
    ...u16(0),
  ]);
}

function refusal(bytes: Uint8Array): ReaderError {
  try {
    readLocalHeaders(bytes);
  } catch (e) {
    if (e instanceof ReaderError) return e;
    throw e;
  }
  throw new Error("expected a refusal, got a reading");
}

const slice = (bytes: Uint8Array, e: { dataOffset: number; compressedSize: number }): number[] =>
  Array.from(bytes.subarray(e.dataOffset, e.dataOffset + e.compressedSize));

// ---------------------------------------------------------------------------

describe("readLocalHeaders — what a well-formed archive declares", () => {
  it("each entry: name, directory flag, method, sizes, and where its data starts", () => {
    const bytes = build([
      { name: "word/" },
      { name: "word/document.xml", data: utf8("<w:document/>") },
      { name: "[Content_Types].xml", data: utf8("<Types/>"), method: 8 },
    ]);
    const read = readLocalHeaders(bytes);
    expect(read.count).toBe(3);
    expect(
      read.entries.map((e) => [e.name, e.dir, e.method, e.compressedSize, e.uncompressedSize, e.hasDataDescriptor]),
    ).toEqual([
      ["word/", true, 0, 0, 0, false],
      ["word/document.xml", false, 0, 13, 13, false],
      ["[Content_Types].xml", false, 8, 8, 8, false],
    ]);
    expect(slice(bytes, read.entries[1]!)).toEqual(Array.from(utf8("<w:document/>")));
    expect(slice(bytes, read.entries[2]!)).toEqual(Array.from(utf8("<Types/>")));
  });

  it("a directory flagged by its attribute only: the local name has no slash either, and they agree", () => {
    const read = readLocalHeaders(build([{ name: "media", externalAttributes: 0x10 }]));
    expect(read.entries.map((e) => [e.name, e.dir])).toEqual([["media/", true]]);
  });

  it("bit 3 set (data descriptor, §4.4.4): local CRC and sizes are zero, the central sizes are kept", () => {
    const bytes = build([
      { name: "a.xml", data: utf8("streamed"), descriptor: true },
      { name: "b.xml", data: utf8("next") },
    ]);
    const read = readLocalHeaders(bytes);
    expect(read.entries.map((e) => [e.name, e.compressedSize, e.uncompressedSize, e.hasDataDescriptor])).toEqual([
      ["a.xml", 8, 8, true],
      ["b.xml", 4, 4, false],
    ]);
    expect(slice(bytes, read.entries[0]!)).toEqual(Array.from(utf8("streamed")));
  });

  it("an archive written in streaming mode by jszip itself (data descriptors) is accepted", async () => {
    const zip = new JSZip();
    zip.file("word/document.xml", "<w:document>streamed</w:document>");
    zip.file("[Content_Types].xml", "<Types/>");
    const bytes = await zip.generateAsync({ type: "uint8array", streamFiles: true, compression: "STORE" });
    const read = readLocalHeaders(bytes);
    expect(read.entries.every((e) => e.dir || e.hasDataDescriptor)).toBe(true);
    const doc = read.entries.find((e) => e.name === "word/document.xml")!;
    expect(new TextDecoder().decode(bytes.subarray(doc.dataOffset, doc.dataOffset + doc.compressedSize))).toBe(
      "<w:document>streamed</w:document>",
    );
  });

  it("zip64 (§4.5.3): saturated central sizes and offset, and saturated local sizes, are read from the extra fields", () => {
    const bytes = build([
      { name: "a.xml", data: utf8("first") },
      { name: "b.xml", data: utf8("second"), zip64: true },
    ]);
    const read = readLocalHeaders(bytes);
    expect(read.entries.map((e) => [e.name, e.compressedSize, e.uncompressedSize])).toEqual([
      ["a.xml", 5, 5],
      ["b.xml", 6, 6],
    ]);
    expect(slice(bytes, read.entries[1]!)).toEqual(Array.from(utf8("second")));
  });
});

describe("readLocalHeaders — the two headers of an entry must agree", () => {
  it("a local name that differs from the central name is refused (jszip would take the local one)", () => {
    const e = refusal(build([{ name: "word/document.xml", localName: "word/evil.xml", data: utf8("x") }]));
    expect(e.message).toBe("zip: local header name does not match the central directory");
  });

  it("the refusal never echoes either name", () => {
    const e = refusal(build([{ name: "secret-title.xml", localName: "other-secret.xml" }]));
    expect(e.message).not.toMatch(/secret/);
  });

  it("a local name that differs only by a trailing slash is refused (byte comparison, no rewriting)", () => {
    const e = refusal(build([{ name: "media", externalAttributes: 0x10, localName: "media/" }]));
    expect(e.message).toBe("zip: local header name does not match the central directory");
  });

  it("a different compression method is refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("x"), method: 0, localMethod: 8 }]));
    expect(e.message).toBe("zip: local header does not match the central directory");
  });

  it("a different compressed size (bit 3 clear) is refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), localCompressed: 2 }]));
    expect(e.message).toBe("zip: local header does not match the central directory");
  });

  it("a different uncompressed size (bit 3 clear) is refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), localUncompressed: 1_000_000 }]));
    expect(e.message).toBe("zip: local header does not match the central directory");
  });

  it("bit 3 set locally but not centrally is refused (the two records disagree on how the entry is written)", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), descriptor: true, flags: 0 }]));
    expect(e.message).toBe("zip: local header does not match the central directory");
  });
});

describe("readLocalHeaders — hostile shapes", () => {
  it("a local header offset past the central directory is refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), cdOffset: 10_000 }]));
    expect(e.message).toBe("zip: local header out of bounds");
  });

  it("a local header that would cross into the central directory is refused", () => {
    const e = refusal(build([{ name: "a.xml", cdOffset: 20 }], { gapBeforeCd: 0 }));
    expect(e.message).toBe("zip: local header out of bounds");
  });

  it("no local header signature at the declared offset is refused", () => {
    const e = refusal(
      build([
        { name: "a.xml", data: utf8("abcdefgh") },
        { name: "b.xml", cdOffset: 4 },
      ]),
    );
    expect(e.message).toBe("zip: bad local header signature");
  });

  it("a compressed size larger than the bytes before the central directory is refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), cdCompressed: 5_000, localCompressed: 5_000 }]));
    expect(e.message).toBe("zip: entry data out of bounds");
  });

  it("an unsupported compression method is refused, as jszip does (only stored and deflate)", async () => {
    const bytes = build([{ name: "a.xml", data: utf8("abc"), method: 12 }]);
    expect(refusal(bytes).message).toBe("zip: unsupported compression method");
    await expect(JSZip.loadAsync(bytes)).rejects.toThrow(/compression/);
  });

  it("saturated central sizes without a zip64 extra field are refused", () => {
    const e = refusal(build([{ name: "a.xml", data: utf8("abc"), saturatedWithoutExtra: true }]));
    expect(e.message).toBe("zip: zip64 extra field missing");
  });

  it("what the central directory refuses is still refused (not a zip)", () => {
    expect(refusal(utf8("not a zip at all")).message).toBe("zip: no end of central directory record");
  });

  it("an entry whose header and data lie inside another entry's data is refused (overlapping-entries bomb)", () => {
    // b.xml's local header, as `build` writes it, quoted inside a.xml's stored data; b's central record points there.
    const inner = build([{ name: "b.xml", data: utf8("hello") }]);
    const quoted = inner.subarray(0, 30 + "b.xml".length + "hello".length);
    const bytes = build([
      { name: "a.xml", data: quoted },
      { name: "b.xml", data: utf8("hello"), cdOffset: 30 + "a.xml".length },
    ]);
    expect(refusal(bytes).message).toBe("zip: entries overlap");
  });
});

// ---------------------------------------------------------------------------
// The independent reading: jszip on real documents nobody here wrote.
// ---------------------------------------------------------------------------

/** The public Office documents (`publicFixtures`), as paths from the root of the repository. */
function officeFixtures(): string[] {
  return publicFixtures(".docx", ".pptx", ".xlsx");
}

type Reading = { kind: "read"; entries: string[] } | { kind: "refused" };

/** jszip keeps the raw compressed bytes of each entry in `_data`, a private
 *  field; read here only to compare, and checked to exist so a jszip upgrade
 *  that renames it fails loudly instead of comparing nothing. `isView`, not
 *  `instanceof Uint8Array`: the test runner gives jszip another realm. */
interface JszipInternals {
  _data?: { compressedSize: number; uncompressedSize: number; compressedContent: Uint8Array };
}

const sha = (b: Uint8Array): string => {
  let h = 0x811c9dc5;
  for (const x of b) h = Math.imul(h ^ x, 0x01000193) >>> 0;
  return `${b.length}:${h.toString(16)}`;
};

async function jszipReading(bytes: Uint8Array): Promise<Reading> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { createFolders: false });
  } catch {
    return { kind: "refused" };
  }
  const entries = await Promise.all(
    Object.values(zip.files).map(async (f) => {
      // jszip drops a directory's data: a directory is compared by name only.
      if (f.dir) return `D ${f.name}`;
      const data = (f as unknown as JszipInternals)._data;
      // An EMPTY entry (0 byte, `customXml/item2.xml` of fixtures/docx-corpus-cases/123, 19/10/2026): jszip keeps no compressed bytes, only a
      // promise of its (empty) content. Compared by name and size 0 only, checked empty, so a renamed field still fails loudly elsewhere.
      if (data !== undefined && typeof (data as unknown as { then?: unknown }).then === "function") {
        if ((await f.async("uint8array")).length !== 0) throw new Error("jszip internals changed");
        return `F ${f.name} 0 empty`;
      }
      if (!data || !ArrayBuffer.isView(data.compressedContent)) throw new Error("jszip internals changed");
      return `F ${f.name} ${data.uncompressedSize} ${sha(data.compressedContent)}`;
    }),
  );
  return { kind: "read", entries: entries.sort() };
}

function ourReading(bytes: Uint8Array): Reading {
  try {
    const entries = readLocalHeaders(bytes).entries.map((e) =>
      e.dir
        ? `D ${e.name}`
        : e.uncompressedSize === 0
          ? `F ${e.name} 0 empty`
          : `F ${e.name} ${e.uncompressedSize} ${sha(bytes.subarray(e.dataOffset, e.dataOffset + e.compressedSize))}`,
    );
    return { kind: "read", entries: entries.sort() };
  } catch (e) {
    if (e instanceof ReaderError) return { kind: "refused" };
    throw e;
  }
}

describe("readLocalHeaders — agrees with jszip", () => {
  const fixtures = officeFixtures();

  it("the fixture corpus is not empty (a silent zero would prove nothing)", () => {
    expect(fixtures.length).toBeGreaterThan(20);
  });

  it("on every real Office fixture: same names, flags, sizes and compressed bytes, or both refuse", async () => {
    const disagreements: string[] = [];
    let read = 0;
    for (const path of fixtures) {
      const bytes = new Uint8Array(readFileSync(path));
      const theirs = await jszipReading(bytes);
      const ours = ourReading(bytes);
      if (ours.kind === "read") read++;
      if (JSON.stringify(theirs) !== JSON.stringify(ours)) disagreements.push(path);
    }
    expect(disagreements).toEqual([]);
    expect(read).toBeGreaterThan(20);
  });
});
