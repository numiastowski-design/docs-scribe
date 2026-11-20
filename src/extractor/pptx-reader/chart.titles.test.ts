/** How a chart's titles are READ (`pptx-reader/chart.ts#readChartOutcome`): the chart's own title and the axis titles.
 *  - Paragraphs (`a:p`) and line breaks (`a:br`) of a title are separated by one space, the runs of one paragraph are joined as written
 *    (« Total » + « revenue » in two paragraphs is « Total revenue », never « Totalrevenue »).
 *  - A title linked to a cell (`c:tx/c:strRef`, what Excel writes) holds its text in `c:strCache/c:pt/c:v` (ECMA-376 Part 1 §21.2.2.178
 *    `strRef`, §21.2.2.176 `strCache`; neither is listed in docs/SOURCES-OFFICIELLES.md, which only names the standard: taken from the
 *    norm's schema, not re-checked here). The cache is the text the chart shows.
 *  The independent control reading is held in `pipeline/source-reading-docx.chart-titles.test.ts`. */
import { describe, expect, it } from "vitest";

import type { OpcPackage } from "../ooxml/package";
import { parseXml } from "../ooxml/xml";
import { readChartOutcome } from "./chart";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const run = (t: string): string => `<a:r><a:t>${t}</a:t></a:r>`;
const rich = (...paragraphs: string[]): string =>
  `<c:title><c:tx><c:rich>${paragraphs.map((p) => `<a:p>${p}</a:p>`).join("")}</c:rich></c:tx></c:title>`;
const linked = (text: string): string =>
  `<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>`;

const chartXml = (mainTitle: string, axes: string): string =>
  `<?xml version="1.0"?><c:chartSpace ${C}><c:chart>${mainTitle}<c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>English</c:v></c:pt><c:pt idx="1"><c:v>French</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser></c:barChart>${axes}</c:plotArea></c:chart></c:chartSpace>`;

const readOf = async (xml: string) => {
  const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
  const read = await readChartOutcome(pkg, "word/charts/chart1.xml", { picture: false });
  if (read === undefined || read === "withheld") throw new Error(`no chart data: ${String(read)}`);
  return read;
};

describe("a title in several paragraphs or with line breaks is not welded", () => {
  it("🔴 two paragraphs: « Total revenue », not « Totalrevenue »", async () => {
    expect((await readOf(chartXml(rich(run("Total"), run("revenue")), ""))).title).toBe("Total revenue");
  });

  it("🔴 a line break (`a:br`) inside a paragraph is a space", async () => {
    expect((await readOf(chartXml(rich(`${run("Total")}<a:br/>${run("revenue")}`), ""))).title).toBe("Total revenue");
  });

  it("🔴 the same two rules hold for an axis title (heading the corner cell, and a value axis line)", async () => {
    const axes =
      `<c:catAx><c:axId val="1"/>${rich(run("Calendar"), run("days"))}</c:catAx>` +
      `<c:valAx><c:axId val="2"/>${rich(`${run("Request")}<a:br/>${run("count")}`)}</c:valAx>`;
    const read = await readOf(chartXml("", axes));
    expect(read.data.headers?.[0]).toBe("Calendar days");
    expect(read.axisTitles).toEqual([{ axis: "y", text: "Request count" }]);
  });

  it("🔴 counter-test: the runs of ONE paragraph stay joined as written (a word split in two runs is one word)", async () => {
    expect((await readOf(chartXml(rich(`${run("Reve")}${run("nue")}`), ""))).title).toBe("Revenue");
    expect((await readOf(chartXml(rich(`${run("Total ")}${run("revenue")}`), ""))).title).toBe("Total revenue");
  });
});

describe("a title linked to a cell is read from its cache", () => {
  it("🔴 the chart's own title", async () => {
    expect((await readOf(chartXml(linked("Revenue by language"), ""))).title).toBe("Revenue by language");
  });

  it("🔴 the axis titles (corner cell and a value axis line)", async () => {
    const axes = `<c:catAx><c:axId val="1"/>${linked("Language")}</c:catAx><c:valAx><c:axId val="2"/>${linked("Euros")}</c:valAx>`;
    const read = await readOf(chartXml("", axes));
    expect(read.data.headers?.[0]).toBe("Language");
    expect(read.axisTitles).toEqual([{ axis: "y", text: "Euros" }]);
  });

  it("🔴 counter-test: a deleted axis linked to a cell still shows nothing, and a linked title with an empty cache gives no title", async () => {
    const axes = `<c:valAx><c:axId val="2"/><c:delete val="1"/>${linked("Hidden")}</c:valAx>`;
    expect((await readOf(chartXml("", axes))).axisTitles).toBeUndefined();
    const empty = `<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/></c:strCache></c:strRef></c:tx></c:title>`;
    expect((await readOf(chartXml(empty, ""))).title).toBeUndefined();
  });
});
