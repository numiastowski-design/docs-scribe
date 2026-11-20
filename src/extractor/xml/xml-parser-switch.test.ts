/**
 * The reader reaches OUR XML module and nothing else. The mock below wraps it in a class that counts calls, built
 * with the caller's own options (which proves our reader accepts them as written). A caller that went back to another
 * parser would never reach the wrapper: a test checks that our module was called, and another reads the sources: no
 * production file of this package imports fast-xml-parser, and its manifest does not install it for a consumer.
 * The XML cases of `docx-reader.test.ts` (DTD, depth, namespaces under any prefix, `mc:Choice`, a part that does not
 * parse) run once through our reader, with the outcome pinned as literals.
 * The other caller, the application's reading of a published page, holds its side of the same switch with its own test.
 * « The switch » below is the day the product's callers went from fast-xml-parser to our reader (26/10/2026); the file
 * name keeps the word, and is not renamed, by project rule.
 */
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";

const flag = vi.hoisted(() => ({ calls: 0 }));

vi.mock("./xml-parser", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./xml-parser")>();
  /** Our reader, built from the caller's own options; `parse` counts the call and answers as it does. */
  class Counted<Ours extends { parse(input: string, encoding?: string): unknown }> {
    private readonly mine: Ours;
    constructor(make: new (options: never) => Ours, options: Record<string, unknown>) {
      this.mine = new make(options as never);
    }
    /** `encoding` is what `parseXml` found when it decoded the part's bytes (ours reads under it). */
    parse(input: string, encoding?: string): unknown {
      flag.calls++;
      return this.mine.parse(input, encoding);
    }
  }
  class XmlParser extends Counted<InstanceType<typeof actual.XmlParser>> {
    constructor(options: Record<string, unknown>) {
      super(actual.XmlParser, options);
    }
  }
  class XmlFragmentParser extends Counted<InstanceType<typeof actual.XmlFragmentParser>> {
    constructor(options: Record<string, unknown>) {
      super(actual.XmlFragmentParser, options);
    }
  }
  return { ...actual, XmlParser, XmlFragmentParser };
});

const { readDocx } = await import("../docx-reader/index");
const { default: JSZip } = await import("jszip");

const walk = (dir: string, ext: RegExp): string[] =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return walk(path, ext);
    return ext.test(name) ? [path] : [];
  });

/** A reader's outcome: its result, or the error it threw (class and message: ours are literals), as data. */
const outcome = async (run: () => unknown): Promise<{ ok: unknown } | { error: string }> => {
  try {
    return { ok: await run() };
  } catch (error) {
    return { error: error instanceof Error ? `${error.constructor.name}: ${error.message}` : String(error) };
  }
};

/** The text of every heading and paragraph of a read document, in order (what `readDocx` returned inside `{ ok }`). */
const textsOf = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((block): string[] => {
    if (block.kind === "heading" || block.kind === "paragraph")
      return [block.inlines.map((inline) => inline.text).join("")];
    if (block.kind === "section" || block.kind === "footnote") return textsOf(block.blocks);
    return [];
  });
const readTexts = (result: unknown): string[] =>
  textsOf((result as { ok: { doc: { blocks: readonly IRBlock[] } } }).ok.doc.blocks);
const DAMAGED = { error: "PartDamagedError: the file is damaged or is not a valid Word document" };

describe("the switch, done: the callers go through OUR reader", () => {
  // 🔴 13/11/2026 — THIS GUARD HAD STOPPED COVERING THE READER, WITHOUT FAILING. It swept only `src/`; the split of 12/11
  // moved the readers into `packages/doc-read/src`, and the sweep no longer saw them. It did not go red for all that: it
  // simply no longer looked at anything it protects.
  // A guard that shrinks in silence is worse than no guard — hence the floor below.
  // 19/11/2026: it sweeps ITS OWN PACKAGE alone — in its public repository this package is the whole tree, and the
  // others are installed dependencies or absent. The home repository holds the rest: the three packages swept together,
  // with the root manifest, in `tools/package-source-sweeps.test.ts`; the app's two roots in
  // `src/no-fast-xml-parser.test.ts`.
  // Its own package is found from where this file stands, never by its folder's name: the same path in both
  // repositories.
  const OWN_SOURCES = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
  const OWN_MANIFEST = join(OWN_SOURCES, "..", "package.json");
  const importsIt = /(?:from|import\()\s*["']fast-xml-parser["']/;

  it("no source file of this package outside the tests imports fast-xml-parser, and its manifest does not install it for a consumer", () => {
    const scanned = walk(OWN_SOURCES, /\.tsx?$/i).filter((path) => !/\.test(\.support)?\.tsx?$/.test(path));
    // MEASURED on 19/11/2026: 101 production files in this package. 268 when this test also swept the base (32) and the
    // PDF reader (135), which `tools/package-source-sweeps.test.ts` sweeps since; 461 on 13/11 with the app's two roots.
    // The floor says the sweep does not collapse; it never rises by itself.
    expect(scanned.length, "the sweep shrank: the package's sources have moved").toBeGreaterThan(90);
    const importers = scanned.filter((path) => importsIt.test(readFileSync(path, "utf8")));
    expect(importers.map((path) => path.slice(OWN_SOURCES.length + 1))).toEqual([]);
    const manifest = JSON.parse(readFileSync(OWN_MANIFEST, "utf8")) as {
      dependencies?: Record<string, string>;
      peerDependencies?: Record<string, string>;
    };
    expect(manifest.dependencies ?? {}).not.toHaveProperty("fast-xml-parser");
    expect(manifest.peerDependencies ?? {}).not.toHaveProperty("fast-xml-parser");
  });

  it("the reader reaches our module: a Word part goes through the wrapper", async () => {
    flag.calls = 0;
    await readDocx(await docx(body("<w:p><w:r><w:t>x</w:t></w:r></w:p>")));
    expect(flag.calls, "ooxml/xml.ts").toBeGreaterThan(0);
  });
});

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const docx = async (documentXml: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "uint8array" });
};
const body = (inner: string, ns = W_NS): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${ns}><w:body>${inner}</w:body></w:document>`;
const nested = (controls: number): string =>
  body(
    "<w:sdt><w:sdtContent>".repeat(controls) +
      "<w:p><w:r><w:t>deep</w:t></w:r></w:p>" +
      "</w:sdtContent></w:sdt>".repeat(controls),
  );
const MC = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;

describe("the XML cases of docx-reader.test.ts through readDocx and our reader", () => {
  // What OUR reader returns for each case, as literals.
  const OURS: Record<string, string[] | typeof DAMAGED> = {
    "a DOCTYPE with its own entity (finding 8: refused)": DAMAGED,
    "nesting ~500 deep (finding 7: read)": ["deep"],
    "nesting past the ceiling (finding 7: refused)": DAMAGED,
    "nesting 20 000 deep (finding 7: refused, never a RangeError)": DAMAGED,
    "elements known by their namespace, whatever prefix the producer bound": ["a"],
    "the Strict namespace, read as Transitional": ["s"],
    "mc:AlternateContent: an unknown Requires sends the reader to the Fallback (finding 12)": ["fallback"],
    "references: the five predefined ones, decimal and hexadecimal, a CR LF normalised": ["<>&\"' \u00e9\u20ac\nb"],
    "a comment, a PI and a CDATA section inside a run": ["ab<c>"],
    "a byte-order mark before the declaration": ["bom"],
  };
  it.each([
    [
      "a DOCTYPE with its own entity (finding 8: refused)",
      `<?xml version="1.0"?><!DOCTYPE w:document [<!ENTITY who "Mallory">]><w:document ${W_NS}><w:body><w:p><w:r><w:t>&who;</w:t></w:r></w:p></w:body></w:document>`,
    ],
    ["nesting ~500 deep (finding 7: read)", nested(245)],
    ["nesting past the ceiling (finding 7: refused)", nested(260)],
    ["nesting 20 000 deep (finding 7: refused, never a RangeError)", nested(20_000)],
    [
      "elements known by their namespace, whatever prefix the producer bound",
      body(
        `<x:p><x:r><x:t>a</x:t></x:r></x:p>`,
        `xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
      ),
    ],
    [
      "the Strict namespace, read as Transitional",
      body(`<w:p><w:r><w:t>s</w:t></w:r></w:p>`, `xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"`),
    ],
    [
      "mc:AlternateContent: an unknown Requires sends the reader to the Fallback (finding 12)",
      body(
        `<w:p><w:r><mc:AlternateContent ${MC} xmlns:zz="urn:unknown"><mc:Choice Requires="zz"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>`,
      ),
    ],
    [
      "references: the five predefined ones, decimal and hexadecimal, a CR LF normalised",
      body(`<w:p><w:r><w:t xml:space="preserve">&lt;&gt;&amp;&quot;&apos; &#233;&#x20AC;\r\nb</w:t></w:r></w:p>`),
    ],
    [
      "a comment, a PI and a CDATA section inside a run",
      body(`<w:p><w:r><w:t>a<!--c-->b</w:t><?mso-contentType?><w:t><![CDATA[<c>]]></w:t></w:r></w:p>`),
    ],
    ["a byte-order mark before the declaration", `\uFEFF${body(`<w:p><w:r><w:t>bom</w:t></w:r></w:p>`)}`],
  ])(
    "%s",
    async (label, xml) => {
      const bytes = await docx(xml);
      const mine = await outcome(() => readDocx(bytes));
      const expected = OURS[label]!;
      expect(Array.isArray(expected) ? readTexts(mine) : mine).toEqual(expected);
    },
    60_000,
  );

  it("a part that does not parse (an element never closed) is « damaged »", async () => {
    const broken = await docx(body(`<w:p><w:r><w:t>broken</w:t></w:r>`));
    const mine = await outcome(() => readDocx(broken));
    expect(mine).toEqual({ error: "PartDamagedError: the file is damaged or is not a valid Word document" });
  });

  it("an unclosed tag inside the XML declaration (docx-reader.test.ts): « damaged »", async () => {
    const broken = await docx(`<?xml version="1.0"?<w:corruptedUnclosed><w:document ${W_NS}><w:body/></w:document>`);
    const mine = await outcome(() => readDocx(broken));
    expect(mine).toEqual(DAMAGED);
  });
});
