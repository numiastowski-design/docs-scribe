/**
 * Each pptx tab carries the slide's number (08/10/2026), so the renderer can write
 * the label in the page language. N is the rank among VISIBLE slides (same convention as the "Slide N" title).
 * Run: npx vitest run --project=main src/extractor/pptx.slide-number.test.ts
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractPptx } from "./pptx";

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

interface MadeSlide {
  words: string;
  hidden?: boolean;
}
const deckOf = async (slides: readonly MadeSlide[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("") +
      "</p:sldIdLst></p:presentation>",
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
      `<?xml version="1.0"?><p:sld ${P_NS}${s.hidden ? ' show="0"' : ""}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="s2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>` +
        `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${s.words}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const tabsOf = (blocks: readonly IRBlock[]): { title: string; slideNumber?: number | undefined; text: string }[] =>
  blocks.flatMap((b) =>
    b.kind === "tabs"
      ? b.tabs.map((t) => ({
          title: t.title,
          slideNumber: (t as { slideNumber?: number }).slideNumber,
          text: JSON.stringify(t.blocks),
        }))
      : [],
  );

describe("extractPptx: slideNumber on each tab", () => {
  it("two slides: numbers 1 and 2", async () => {
    const { doc } = await extractPptx(await deckOf([{ words: "alpha" }, { words: "beta" }]));
    expect(tabsOf(doc.blocks).map((t) => t.slideNumber)).toEqual([1, 2]);
  });

  it("a hidden slide in the middle: the numbers follow the rank among the visible ones", async () => {
    const { doc } = await extractPptx(
      await deckOf([{ words: "alpha" }, { words: "secretword", hidden: true }, { words: "gamma" }]),
    );
    const tabs = tabsOf(doc.blocks);
    expect(tabs.map((t) => t.slideNumber)).toEqual([1, 2]);
    expect(tabs[1]!.text).toContain("gamma");
    expect(tabs.map((t) => t.title)).toEqual(["Slide 1", "Slide 2"]);
  });

  it("the number is the same whatever the page language", async () => {
    const { doc } = await extractPptx(await deckOf([{ words: "alpha" }, { words: "beta" }]), undefined, "fr");
    expect(tabsOf(doc.blocks).map((t) => t.slideNumber)).toEqual([1, 2]);
  });
});
