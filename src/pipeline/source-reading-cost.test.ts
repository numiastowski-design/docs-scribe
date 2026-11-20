/** 09/10/2026, counter-review of 9af78c86: two control readings whose cost grew faster than the file. Measured on the review's
 *  scratch copies: a PowerPoint table of n columns whose cells carry gridSpan="n" rowSpan="2" took 444 ms (n=1000), 1.7 s
 *  (2000), 7.7 s and 854 MB (4000), the columns of a span expanded one by one where the page keeps one range
 *  (lib/table-empty-parts.ts#anchorsOf); a chain of n Word footnotes, each calling the one before, listed in increasing order,
 *  took one pass per note: 7.8 s for n=4000 (480 KB) against 160 ms in calling order.
 *  Run: npx vitest run --project=pdf src/pipeline/source-reading-cost.test.ts (serial batch: timing bounds). */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readDocxConservationSource } from "./source-reading-docx";
import { readPptxConservationSource } from "./source-reading-pptx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const CT =
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>';

const P_NS =
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}" ` +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const deckOf = async (tree: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CT);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/>`),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${tree}</p:spTree></p:cSld></p:sld>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};
const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const frame = (data: string): string =>
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="f2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${TABLE_URI}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}"`;
const docxOf = async (body: string, footnotes: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CT);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  zip.file(
    "word/_rels/document.xml.rels",
    RELS(`<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>`),
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  zip.file("word/footnotes.xml", `<?xml version="1.0"?><w:footnotes ${W_NS}>${footnotes}</w:footnotes>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

const timed = async (run: () => Promise<unknown>): Promise<number> => {
  const t = performance.now();
  await run();
  return performance.now() - t;
};

describe("control readings: cost bounded on hostile tables and notes", () => {
  it("PowerPoint: 4 000 columns, each cell gridSpan=4000 rowSpan=2, under 1 s (7.7 s before)", async () => {
    const n = 4000;
    const grid = `<a:tblGrid>${'<a:gridCol w="1"/>'.repeat(n)}</a:tblGrid>`;
    const cell = (t: string, attrs: string): string =>
      `<a:tc${attrs}><a:txBody><a:bodyPr/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
    const rowOne = `<a:tr h="1">${Array.from({ length: n }, (_, c) => cell(`s${c}`, ` gridSpan="${n}" rowSpan="2"`)).join("")}</a:tr>`;
    const rowTwo = `<a:tr h="1">${cell("last", "")}</a:tr>`;
    const bytes = await deckOf(frame(`<a:tbl><a:tblPr/>${grid}${rowOne}${rowTwo}</a:tbl>`));
    expect(await timed(() => readPptxConservationSource(bytes, "en"))).toBeLessThan(1000);
  });

  it("Word: a chain of 4 000 footnotes each calling the previous one, listed in increasing order, under 1 s (7.8 s before)", async () => {
    const n = 4000;
    const notes = Array.from(
      { length: n },
      (_, k) =>
        `<w:footnote w:id="${k + 1}"><w:p><w:r><w:t>note${k + 1}</w:t></w:r>${k > 0 ? `<w:r><w:footnoteReference w:id="${k}"/></w:r>` : ""}</w:p></w:footnote>`,
    ).join("");
    const bytes = await docxOf(
      `<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="${n}"/></w:r></w:p>`,
      notes,
    );
    expect(await timed(() => readDocxConservationSource(bytes))).toBeLessThan(1000);
  });
});
