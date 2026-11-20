/** Behaviour tests for `readDocxConservationSource`
 *  (MIKSUDOC « Spécification technique » T11): a real docx yields a
 *  usable declared-format source, and a read that cannot be trusted becomes
 *  a named failure — never a thrown error, never a reason not to publish
 *  (21/09/2026). `readSource`/`wordsOf` themselves are exercised
 *  exhaustively, against every real fixture, by the app's `conservation/no-word-lost.test.ts`. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { isDeclaredFormatSource, isReadFailure } from "./conservation-source";
import { readDocxConservationSource, readNumberedParagraphs, readSource, wordsOf } from "./source-reading-docx";

const root = join(dirname(fileURLToPath(import.meta.url)), "../..");
const fixture = (relative: string): Uint8Array => new Uint8Array(readFileSync(join(root, relative)));

describe("wordsOf — a word is a maximal run of letters/digits, NFKC lower case", () => {
  it("tokenises accents, hyphens and digits the same way the page is read", () => {
    expect(wordsOf("L'ÉTÉ 2026 : c'est décidé-là.")).toEqual(["l", "été", "2026", "c", "est", "décidé", "là"]);
  });
});

describe("readDocxConservationSource — real docx and Google-Doc-export fixtures", () => {
  it("a real docx: one unit, no declared removal, no excluded unit, words present", async () => {
    const source = await readDocxConservationSource(fixture("fixtures/docx-reel/01-osfi-guide-portail.docx"));
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source, got a read failure");
    expect(source.units.length).toBe(1);
    expect(source.units[0]!.tokens.length).toBeGreaterThan(100);
    expect(source.deliberateRemovals).toEqual([]);
    expect(source.excludedUnits.size).toBe(0);
  });

  it("a Google Doc export (same docx shape, fixtures/gdoc-docx/): reads the same way", async () => {
    const source = await readDocxConservationSource(fixture("fixtures/gdoc-docx/01-Guide-Onboarding-Atria.docx"));
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source, got a read failure");
    expect(source.units[0]!.tokens.length).toBeGreaterThan(20);
  });

  it("🔴 GUARD BITES: bytes that are not a zip at all — a named failure, never a thrown error", async () => {
    const source = await readDocxConservationSource(new TextEncoder().encode("not a docx"));
    expect(isReadFailure(source)).toBe(true);
  });

  it("🔴 GUARD BITES: a zip with no word/document.xml — a named failure", async () => {
    const zip = new JSZip();
    zip.file("word/styles.xml", "<w:styles/>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const source = await readDocxConservationSource(bytes);
    expect(isReadFailure(source)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Review findings A, H, 14 (30/09/2026): made-up packages.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const docxOf = async (parts: { body: string; footnotes?: string }): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>${parts.body}</w:body></w:document>`,
  );
  if (parts.footnotes !== undefined) zip.file("word/footnotes.xml", parts.footnotes);
  return zip.generateAsync({ type: "uint8array" });
};

describe("readDocxConservationSource — hardened reads of the hard-coded parts (finding A)", () => {
  it("🔴 a DTD in word/footnotes.xml (a part the extractor never opens) is refused: an abstention, not a reading", async () => {
    const bytes = await docxOf({
      body: "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>",
      footnotes: `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><w:footnotes xmlns:w="${W}"/>`,
    });
    expect(isReadFailure(await readDocxConservationSource(bytes))).toBe(true);
  });

  it("🔴 a part nested past the extractor's depth ceiling is refused, not walked", async () => {
    const depth = 3000;
    const bytes = await docxOf({
      body: "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>",
      footnotes: `<?xml version="1.0"?><w:footnotes xmlns:w="${W}">${"<w:x>".repeat(depth)}${"</w:x>".repeat(depth)}</w:footnotes>`,
    });
    expect(isReadFailure(await readDocxConservationSource(bytes))).toBe(true);
  });

  it("a part inflating past PART_MAX_BYTES abstains", async () => {
    const zip = new JSZip();
    const bytes = await docxOf({ body: "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" });
    const loaded = await JSZip.loadAsync(bytes);
    zip.file("[Content_Types].xml", await loaded.file("[Content_Types].xml")!.async("string"));
    zip.file("_rels/.rels", await loaded.file("_rels/.rels")!.async("string"));
    zip.file("word/document.xml", await loaded.file("word/document.xml")!.async("string"));
    zip.file("word/footnotes.xml", `<w:footnotes xmlns:w="${W}"><!--${" ".repeat(65_000_000)}--></w:footnotes>`, {
      compression: "DEFLATE",
    });
    const bomb = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    expect(bomb.byteLength).toBeLessThan(1_000_000);
    expect(isReadFailure(await readDocxConservationSource(bomb))).toBe(true);
  }, 60_000);
});

describe("readSource — linear in the note references (finding H) and the text (finding 14)", () => {
  it("🔴 40000 distinct note references read in linear time (quadratic took ~70 s)", async () => {
    const n = 40_000;
    let refs = "";
    let notes = "";
    for (let i = 1; i <= n; i++) {
      refs += `<w:r><w:footnoteReference w:id="${i}"/></w:r>`;
      notes += `<w:footnote w:id="${i}"><w:p><w:r><w:t>n${i}</w:t></w:r></w:p></w:footnote>`;
    }
    const bytes = await docxOf({
      body: `<w:p>${refs}</w:p>`,
      footnotes: `<w:footnotes xmlns:w="${W}">${notes}</w:footnotes>`,
    });
    const t0 = performance.now();
    const { text } = await readSource(bytes);
    const elapsed = performance.now() - t0;
    expect(text).toContain("n40000");
    expect(elapsed).toBeLessThan(15_000);
  }, 60_000);

  it("word causes stay aligned with words across chunk boundaries", async () => {
    const bytes = await docxOf({
      body: "<w:p><w:r><w:t>al</w:t></w:r><w:r><w:t>pha beta</w:t></w:r></w:p><w:p><w:r><w:t>gamma</w:t></w:r></w:p>",
    });
    const { text, wordCauses } = await readSource(bytes);
    expect(wordsOf(text)).toEqual(["alpha", "beta", "gamma"]);
    expect(wordCauses).toEqual(["text", "text", "text"]);
  });
});

describe("readSource — no spread of an unbounded list into call arguments", () => {
  it("🔴 one paragraph of 250000 runs is read (a spread past ~100000 arguments threw a RangeError)", async () => {
    const runs = "<w:r><w:t>a</w:t></w:r>".repeat(250_000);
    const bytes = await docxOf({ body: `<w:p>${runs}</w:p>` });
    const { text } = await readSource(bytes);
    expect(text.length).toBeGreaterThanOrEqual(250_000);
  }, 60_000);

  it("readNumberedParagraphs names the numbered paragraphs, readSource no longer carries them", async () => {
    const bytes = await docxOf({
      body: '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item one</w:t></w:r></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>',
    });
    expect(await readNumberedParagraphs(bytes)).toEqual(["item one"]);
    // The structure counts joined the reading on 20/10/2026 (`source-structure-docx.ts`); the numbered paragraphs are still not carried.
    expect(Object.keys(await readSource(bytes)).sort()).toEqual(["structure", "text", "wordCauses"]);
  });
});
