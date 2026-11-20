/** The control reading of a chart's titles (`pipeline/source-reading-chart.ts`, independent of `extractor/pptx-reader/chart.ts`) expects
 *  the words the page shows: a title in several paragraphs or with `a:br` is read word by word (« Total revenue », never the welded
 *  « Totalrevenue »), and a title linked to a cell (`c:tx/c:strRef`, ECMA-376 Part 1 §21.2.2.178 `strRef`, §21.2.2.176 `strCache`) is read
 *  from its cache, for the chart's own title and for the axes'. A deleted axis stays unexpected. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readSource, wordsOf } from "./source-reading-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"`;
const drawing =
  `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const run = (t: string): string => `<a:r><a:t>${t}</a:t></a:r>`;
const rich = (...paragraphs: string[]): string =>
  `<c:title><c:tx><c:rich>${paragraphs.map((p) => `<a:p>${p}</a:p>`).join("")}</c:rich></c:tx></c:title>`;
const linked = (text: string): string =>
  `<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>`;

const chartPart = (mainTitle: string, axes: string): string =>
  `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>${mainTitle}` +
  `<c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>S!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>English</c:v></c:pt><c:pt idx="1"><c:v>French</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>S!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser></c:barChart>${axes}</c:plotArea></c:chart></c:chartSpace>`;

const docxWithChart = async (chart: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="${REL}/chart" Target="charts/chart1.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r>${drawing}</w:r></w:p></w:body></w:document>`,
  );
  zip.file("word/charts/chart1.xml", chart);
  return zip.generateAsync({ type: "uint8array" });
};

const bag = async (chart: string): Promise<string[]> =>
  [...wordsOf((await readSource(await docxWithChart(chart))).text)].sort();
const base = ["share", "english", "french", "5", "7"];

describe("the control reading of a chart's titles", () => {
  it("🔴 two paragraphs and an `a:br` are separate words, not welded", async () => {
    const axes = `<c:catAx><c:axId val="1"/>${rich(`${run("Calendar")}<a:br/>${run("days")}`)}</c:catAx>`;
    expect(await bag(chartPart(rich(run("Total"), run("revenue")), axes))).toEqual(
      [...base, "total", "revenue", "calendar", "days"].sort(),
    );
  });

  it("🔴 counter-test: the runs of one paragraph stay joined (« Reve » + « nue » is one word)", async () => {
    expect(await bag(chartPart(rich(`${run("Reve")}${run("nue")}`), ""))).toEqual([...base, "revenue"].sort());
  });

  it("🔴 a title linked to a cell is read from its cache, main and axes", async () => {
    const axes = `<c:catAx><c:axId val="1"/>${linked("Language")}</c:catAx><c:valAx><c:axId val="2"/>${linked("Euros")}</c:valAx>`;
    expect(await bag(chartPart(linked("Revenue by language"), axes))).toEqual(
      [...base, "revenue", "by", "language", "language", "euros"].sort(),
    );
  });

  it("🔴 counter-test: a deleted axis linked to a cell is still not expected", async () => {
    const axes = `<c:valAx><c:axId val="2"/><c:delete val="1"/>${linked("Hidden")}</c:valAx>`;
    expect(await bag(chartPart("", axes))).toEqual([...base].sort());
  });
});
