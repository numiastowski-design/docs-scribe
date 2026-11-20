/**
 * A slide's footer, date and slide-number placeholders are published on its tab (19/10/2026: « on les
 * garde, sur chaque onglet, sans se soucier » — no text lost, no special handling). PowerPoint only shows these
 * placeholders when the slide itself carries them, so the slide is the only place read.
 * Run: npx vitest run --project=main src/extractor/pptx.footer-placeholders.test.ts
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractPptx } from "./pptx";
import { readPptx } from "./pptx-reader";

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const xfrm = (x: number, y: number, cx: number, cy: number): string =>
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>`;

const placeholder = (id: number, type: string, paragraph: string, frame = "<p:spPr/>"): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${id + 10}"/></p:nvPr></p:nvSpPr>${frame}` +
  `<p:txBody><a:bodyPr/><a:p>${paragraph}</a:p></p:txBody></p:sp>`;

/** One slide with footer, date and number; `hiddenFirst` puts a hidden slide before it (it still counts in the number). */
const deckWithFooters = async (opts: { firstSlideNum?: number; hiddenFirst?: boolean } = {}): Promise<Uint8Array> => {
  const zip = new JSZip();
  const first = opts.firstSlideNum === undefined ? "" : ` firstSlideNum="${opts.firstSlideNum}"`;
  const ids = opts.hiddenFirst ? ["rId2", "rId1"] : ["rId1"];
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}${first}><p:sldIdLst>${ids.map((id, i) => `<p:sldId id="${256 + i}" r:id="${id}"/>`).join("")}</p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      `<Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="slides/slide2.xml"/>`,
    ),
  );
  zip.file(
    "ppt/slides/slide2.xml",
    `<?xml version="1.0"?><p:sld ${P_NS} show="0"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="h"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>hiddenslide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly results</a:t></a:r></a:p></p:txBody></p:sp>` +
      placeholder(3, "ftr", "<a:r><a:t>Atria Solutions confidential</a:t></a:r>") +
      placeholder(4, "dt", '<a:fld id="{A1}" type="datetime1"><a:t>19/10/2026</a:t></a:fld>') +
      placeholder(5, "sldNum", '<a:fld id="{B2}" type="slidenum"><a:t>‹#›</a:t></a:fld>') +
      `</p:spTree></p:cSld></p:sld>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

/** A declared `p:sldSz` and positions: `layoutSlide` really runs. The body fills the slide's top; footer, date and
 *  number sit side by side along the bottom, so were they laid out they would form a row of columns. */
const positionedDeck = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/>`),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${xfrm(500000, 500000, 11000000, 5000000)}<p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly results</a:t></a:r></a:p></p:txBody></p:sp>` +
      placeholder(
        3,
        "ftr",
        "<a:r><a:t>Atria Solutions confidential</a:t></a:r>",
        xfrm(500000, 6300000, 4000000, 400000),
      ) +
      placeholder(
        4,
        "dt",
        '<a:fld id="{A1}" type="datetime1"><a:t>19/10/2026</a:t></a:fld>',
        xfrm(5000000, 6300000, 3000000, 400000),
      ) +
      placeholder(
        5,
        "sldNum",
        '<a:fld id="{B2}" type="slidenum"><a:t>‹#›</a:t></a:fld>',
        xfrm(9000000, 6300000, 2500000, 400000),
      ) +
      `</p:spTree></p:cSld></p:sld>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

const tabTexts = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((b) => (b.kind === "tabs" ? b.tabs.map((t) => JSON.stringify(t.blocks)) : []));

describe("extractPptx: footer, date and slide number stay on the tab", () => {
  it("publishes the footer placeholder's text", async () => {
    const { doc } = await extractPptx(await deckWithFooters());
    expect(tabTexts(doc.blocks)[0]).toContain("Atria Solutions confidential");
  });

  it("publishes the date field's displayed text", async () => {
    const { doc } = await extractPptx(await deckWithFooters());
    expect(tabTexts(doc.blocks)[0]).toContain("19/10/2026");
  });

  // The slide-number field stores a placeholder (« ‹#› » in real decks) and PowerPoint DISPLAYS the slide's number: its
  // position in `p:sldIdLst`, hidden slides counted, plus `firstSlideNum` − 1 (ECMA-376 Part 1, §19.3.1.26 « presentation »,
  // attribute firstSlideNum, default 1). The displayed value is published, never the stored one.
  it("publishes the number PowerPoint displays, never the stored placeholder", async () => {
    const { doc } = await extractPptx(await deckWithFooters());
    expect(tabTexts(doc.blocks)[0]).toContain('"text":"1"');
    expect(tabTexts(doc.blocks)[0]).not.toContain("‹#›");
  });

  it("follows firstSlideNum and counts hidden slides", async () => {
    const { doc } = await extractPptx(await deckWithFooters({ firstSlideNum: 5, hiddenFirst: true }));
    expect(tabTexts(doc.blocks)).toHaveLength(1);
    expect(tabTexts(doc.blocks)[0]).toContain('"text":"6"');
  });

  it("keeps the body a one-column flow and closes the tab with footer, date, number", async () => {
    const deck = await readPptx(await positionedDeck());
    expect(deck.slideSize).toBeDefined(); // layoutSlide runs only with a declared size
    const { doc } = await extractPptx(await positionedDeck());
    const first = doc.blocks[0];
    if (first?.kind !== "tabs") throw new Error("expected tabs");
    const blocks = first.tabs[0]!.blocks;
    expect(blocks.some((b) => b.kind === "table")).toBe(false);
    const texts = blocks.map((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : `<${b.kind}>`));
    expect(texts).toEqual(["Quarterly results", "Atria Solutions confidential", "19/10/2026", "1"]);
  });
});
