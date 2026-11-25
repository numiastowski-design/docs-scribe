/** What `readDocx` hands the stories of the notes: a drawing in a footnote is kept out like one in the body
 *  (`docx-reader/index.ts#readDocx` builds the context of the notes apart from the body's). Written before `readDocx` was cut
 *  into one function per step (lot 10C1); every expectation was seen green on the uncut function. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readDocx } from "./index";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:r="${REL}"`;

/** A text box drawing; `placement` is the wp:inline or the wp:anchor that carries it. */
const textBox = (words: string, id: number, docPr: string, placement: "inline" | "far"): string => {
  const head =
    placement === "inline"
      ? "<wp:inline>"
      : '<wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>99000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionV>';
  const tail = placement === "inline" ? "</wp:inline>" : "</wp:anchor>";
  return (
    `<w:r><w:drawing>${head}<wp:extent cx="1000000" cy="500000"/><wp:docPr id="${id}" name="b${id}"${docPr}/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
    '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></wps:spPr>' +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${words}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>${tail}</w:drawing></w:r>`
  );
};

const withFootnote = async (noteRuns: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r><w:t>host</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>${a4}</w:body></w:document>`,
  );
  zip.file(
    "word/footnotes.xml",
    `<?xml version="1.0"?><w:footnotes ${NS}><w:footnote w:id="1"><w:p>${noteRuns}</w:p></w:footnote></w:footnotes>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

describe("readDocx — hidden and off-page drawings in a footnote", () => {
  it("a drawing hidden by its docPr is kept out and counted, its words unpublished", async () => {
    const { doc, notices } = await readDocx(
      await withFootnote(textBox("secretwords", 1, ' hidden="1"', "inline") + textBox("shownwords", 2, "", "inline")),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).not.toContain("secretwords");
    expect(text).toContain("shownwords");
    expect(notices).toContainEqual({ kind: "hidden-dropped", what: "shape", count: 1 });
  });

  it("a drawing anchored wholly off the page (the page being the body's) is not published", async () => {
    const { doc } = await readDocx(
      await withFootnote(textBox("farwords", 1, "", "far") + textBox("nearwords", 2, "", "inline")),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).not.toContain("farwords");
    expect(text).toContain("nearwords");
  });
});

describe("readDocx — what the result carries", () => {
  it("a document without drawings has no drawingShapesNotDrawn at all", async () => {
    const result = await readDocx(await withFootnote(""));
    expect(result).not.toHaveProperty("drawingShapesNotDrawn");
    expect(result).not.toHaveProperty("notices");
  });
});
