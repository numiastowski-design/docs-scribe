import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readDocx } from "./docx-reader";
import { extractPptx } from "./pptx";

// Decision n° 6 (12/10/2026): a SmartArt drawing cache with unreadable XML, for a SmartArt with NO data model, is
// left out (no words for that cache); the document is extracted, everything else intact. Before: the whole extraction threw
// the « damaged document » error.
// Run: npx vitest run --project=main src/extractor/smartart.damaged-cache.test.ts

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const TYPES =
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>';
const GOOD_CACHE = `<?xml version="1.0"?><dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${A}"><dsp:spTree><dsp:sp><dsp:txBody><a:p><a:r><a:t>cachedwords</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`;
const BROKEN_CACHE = '<?xml version="1.0"?><dsp:drawing <<< not xml';

const docxOf = async (cache: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", TYPES);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  // The frame's r:dm names no relationship: a SmartArt with no data model part; the cache is nobody's.
  zip.file(
    "word/_rels/document.xml.rels",
    RELS(`<Relationship Id="rId6" Type="${DRAWING_REL}" Target="diagrams/drawing1.xml"/>`),
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${A}" xmlns:r="${REL}"><w:body>` +
      "<w:p><w:r><w:t>beforewords</w:t></w:r></w:p>" +
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
      "<w:p><w:r><w:t>afterwords</w:t></w:r></w:p></w:body></w:document>",
  );
  zip.file("word/diagrams/drawing1.xml", cache);
  return zip.generateAsync({ type: "uint8array" });
};

const P_NS = `xmlns:a="${A}" xmlns:r="${REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const textShape = (id: number, words: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
const smartFrame =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="sa"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>' +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/></a:graphicData></a:graphic></p:graphicFrame>`;
const deckOf = async (cache: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", TYPES);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>${[1, 2, 3].map((n) => `<p:sldId id="${255 + n}" r:id="rId${n}"/>`).join("")}</p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      [1, 2, 3].map((n) => `<Relationship Id="rId${n}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`).join(""),
    ),
  );
  const slide = (tree: string): string =>
    `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${tree}</p:spTree></p:cSld></p:sld>`;
  zip.file("ppt/slides/slide1.xml", slide(textShape(2, "firstslidewords")));
  zip.file("ppt/slides/slide2.xml", slide(smartFrame));
  zip.file(
    "ppt/slides/_rels/slide2.xml.rels",
    RELS(`<Relationship Id="rId20" Type="${DRAWING_REL}" Target="../diagrams/drawing1.xml"/>`),
  );
  zip.file("ppt/slides/slide3.xml", slide(textShape(2, "thirdslidewords")));
  zip.file("ppt/diagrams/drawing1.xml", cache);
  return zip.generateAsync({ type: "uint8array" });
};

describe("a SmartArt with no data model and a damaged drawing cache", () => {
  it("control: a readable cache gives its words (docx)", async () => {
    const { doc } = await readDocx(await docxOf(GOOD_CACHE));
    expect(JSON.stringify(doc.blocks)).toContain("cachedwords");
  });

  it("control: a readable cache gives its words (pptx)", async () => {
    const { doc } = await extractPptx(await deckOf(GOOD_CACHE));
    expect(JSON.stringify(doc.blocks)).toContain("cachedwords");
  });

  it("🔴 docx: the document is extracted, the other paragraphs intact, the SmartArt without words", async () => {
    const { doc } = await readDocx(await docxOf(BROKEN_CACHE));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("beforewords");
    expect(text).toContain("afterwords");
    expect(text).not.toContain("cachedwords");
  });

  it("🔴 pptx: the deck is extracted, the other slides intact, the SmartArt without words", async () => {
    const { doc } = await extractPptx(await deckOf(BROKEN_CACHE));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("firstslidewords");
    expect(text).toContain("thirdslidewords");
    expect(text).not.toContain("cachedwords");
  });
});
