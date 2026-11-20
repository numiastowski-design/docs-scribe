/** Behaviour tests for `readPptxConservationSource`
 *  (MIKSUDOC « Spécification technique » T11): a real pptx yields a
 *  usable declared-format source, a read that cannot be trusted becomes a
 *  named failure — never a thrown error, never a reason not to publish
 *  (21/09/2026) — and the three declared removals (display cap,
 *  element-cap withhold, repeated title) never surface as missing words.
 *  `readPptxSource` itself is exercised exhaustively, against every real
 *  fixture, by the app's `conservation/no-word-lost.test.ts`. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_PRESENTATION_TABS, MAX_SLIDE_ELEMENTS, PAGE_SIZED_LIMITS } from "../extractor/page-caps.test.support";
import { isDeclaredFormatSource, isReadFailure } from "./conservation-source";
import { readPptxConservationSource } from "./source-reading-pptx";

const root = join(dirname(fileURLToPath(import.meta.url)), "../..");
const fixture = (relative: string): Uint8Array => new Uint8Array(readFileSync(join(root, relative)));

// ---------------------------------------------------------------------------
// Made-up decks — same shape as `extractor/pptx.test.ts`'s own `deckOf`.

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

interface MadeSlide {
  tree: string;
  hidden?: boolean;
}

const deckOf = async (slides: readonly MadeSlide[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("") +
      `</p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      slides
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  slides.forEach((s, i) => {
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}${s.hidden ? ' show="0"' : ""}><p:cSld><p:spTree>${s.tree}</p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const text = (id: number, words: string, ph = ""): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
const title = (id: number, words: string): string => text(id, words, '<p:ph type="title"/>');

const allWords = (source: Awaited<ReturnType<typeof readPptxConservationSource>>): string[] => {
  if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
  return source.units.flatMap((u) => u.tokens.map((t) => t.text));
};

describe("readPptxConservationSource — the Symbol font, like the page (26/10/2026)", () => {
  it("🔴 a run whose a:sym typeface is Symbol is read as the characters it draws", async () => {
    // FAILS WITH: « \uF071 » — the page publishes θ, the second reading would count it invented.
    const tree =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:sym typeface="Symbol"/></a:rPr><a:t>cos \uF071</a:t></a:r></a:p></p:txBody></p:sp>';
    expect(allWords(await readPptxConservationSource(await deckOf([{ tree }]), "en", PAGE_SIZED_LIMITS))).toContain(
      "θ",
    );
  });
});

describe("readPptxConservationSource — a:sym decides first (counter-review of 26/10/2026)", () => {
  it("🔴 a:latin Symbol with a:sym Wingdings: the point is kept, like the page", async () => {
    // FAILS WITH: « θ » read where the page keeps the Wingdings point.
    const tree =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="Symbol"/><a:sym typeface="Wingdings"/></a:rPr><a:t>ok \uF071</a:t></a:r></a:p></p:txBody></p:sp>';
    expect(allWords(await readPptxConservationSource(await deckOf([{ tree }]), "en", PAGE_SIZED_LIMITS))).not.toContain(
      "θ",
    );
  });
});

describe("readPptxConservationSource — real pptx fixtures", () => {
  it("a real deck: one unit per slide, no excluded unit, words present", async () => {
    const source = await readPptxConservationSource(
      fixture("fixtures/pptx/03-Presentation-Resultats-T1.pptx"),
      "en",
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source, got a read failure");
    expect(source.units.length).toBe(6);
    expect(source.units.map((u) => u.label)).toEqual([
      "Slide 1",
      "Slide 2",
      "Slide 3",
      "Slide 4",
      "Slide 5",
      "Slide 6",
    ]);
    expect(source.excludedUnits.size).toBe(0);
    expect(source.deliberateRemovals).toEqual([]);
    expect(source.units.flatMap((u) => u.tokens).length).toBeGreaterThan(20);
  });

  it("🔴 GUARD BITES: bytes that are not a zip at all — a named failure, never a thrown error", async () => {
    const source = await readPptxConservationSource(new TextEncoder().encode("not a pptx"), "en", PAGE_SIZED_LIMITS);
    expect(isReadFailure(source)).toBe(true);
  });

  it("🔴 GUARD BITES: a zip with no ppt/presentation.xml — a named failure", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", "<p:sld/>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const source = await readPptxConservationSource(bytes, "en", PAGE_SIZED_LIMITS);
    expect(isReadFailure(source)).toBe(true);
  });
});

describe("readPptxConservationSource — the three declared removals never surface as missing", () => {
  it("a slide past the display cap: its unit is excluded, never a source of tokens", async () => {
    const slides: MadeSlide[] = Array.from({ length: MAX_PRESENTATION_TABS + 2 }, (_, i) => ({
      tree: text(2, `word${i + 1}unique`),
    }));
    const source = await readPptxConservationSource(await deckOf(slides), "en", PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(MAX_PRESENTATION_TABS + 2);
    // The two over-cap slides: excluded, and their own words never collected.
    expect([...source.excludedUnits].sort((a, b) => a - b)).toEqual([MAX_PRESENTATION_TABS, MAX_PRESENTATION_TABS + 1]);
    expect(source.units[MAX_PRESENTATION_TABS]!.tokens).toEqual([]);
    expect(source.units[MAX_PRESENTATION_TABS + 1]!.tokens).toEqual([]);
    // A kept slide's word is still there.
    expect(allWords(source)).toContain("word1unique");
    expect(allWords(source)).not.toContain(`word${MAX_PRESENTATION_TABS + 1}unique`);
  });

  it("a slide over MAX_SLIDE_ELEMENTS: withheld whole, its unit excluded and empty", async () => {
    const tree = Array.from({ length: MAX_SLIDE_ELEMENTS + 2 }, (_, i) => text(i + 2, `e${i}`)).join("");
    const source = await readPptxConservationSource(
      await deckOf([{ tree }, { tree: text(2, "next slide unique") }]),
      "en",
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(2);
    expect([...source.excludedUnits]).toEqual([0]);
    expect(source.units[0]!.tokens).toEqual([]);
    expect(allWords(source)).not.toContain("e0");
    expect(allWords(source)).toContain("unique");
  });

  it("a slide title identical to the previous slide's: not collected a second time", async () => {
    const slides: MadeSlide[] = [
      { tree: title(2, "Overview") + text(3, "first body unique") },
      { tree: title(2, "Overview") + text(3, "second body unique") },
    ];
    const source = await readPptxConservationSource(await deckOf(slides), "en", PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.excludedUnits.size).toBe(0);
    const words = allWords(source);
    expect(words.filter((w) => w === "overview")).toEqual(["overview"]);
    expect(words).toContain("first");
    expect(words).toContain("second");
  });

  it("a slide-number field: the number PowerPoint displays is expected, never the stored placeholder (19/10/2026)", async () => {
    const slideNumber =
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="n3"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:p><a:fld id="{B2}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>';
    const words = allWords(
      await readPptxConservationSource(
        await deckOf([{ tree: text(2, "bodywords") + slideNumber }]),
        "en",
        PAGE_SIZED_LIMITS,
      ),
    );
    expect(words).toContain("1");
    expect(words).not.toContain("‹#›");
  });

  it("a hidden slide: not a unit at all", async () => {
    const slides: MadeSlide[] = [
      { tree: text(2, "shown one unique") },
      { tree: text(2, "hidden words unique"), hidden: true },
      { tree: text(2, "shown two unique") },
    ];
    const source = await readPptxConservationSource(await deckOf(slides), "en", PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(2);
    expect(allWords(source)).not.toContain("hidden");
  });

  // Decision n° 7 (12/10/2026): the « hidden » flag of a shape is read tolerantly, like the extractor (own copy of the rule).
  it.each(["1", "true", "TRUE", " 1 ", " True "])('a shape hidden="%s": its words are not expected', async (value) => {
    const hiddenShape = text(3, "tolerantsecret").replace('name="s3"', `name="s3" hidden="${value}"`);
    const words = allWords(
      await readPptxConservationSource(
        await deckOf([{ tree: text(2, "shownwords") + hiddenShape }]),
        "en",
        PAGE_SIZED_LIMITS,
      ),
    );
    expect(words).toContain("shownwords");
    expect(words).not.toContain("tolerantsecret");
  });

  it.each(["0", "false", "no"])('a shape hidden="%s": its words are still expected', async (value) => {
    const shown = text(3, "stillshown").replace('name="s3"', `name="s3" hidden="${value}"`);
    expect(
      allWords(await readPptxConservationSource(await deckOf([{ tree: shown }]), "en", PAGE_SIZED_LIMITS)),
    ).toContain("stillshown");
  });
});
