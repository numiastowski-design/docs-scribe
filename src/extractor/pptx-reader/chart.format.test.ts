/** The numbers of a chart's data table are the numbers the chart DISPLAYS, not the ones its cache stores (rule,
 *  19/10/2026; constat on `fixtures/docx-corpus-cases/200-…`: `0.80861244019138756` published where Word shows « 81 % »).
 *  The format of a value is, in this order: the `c:numFmt` of the series' data labels (`c:dLbls`) when it is not source-linked,
 *  else the `c:formatCode` of the series' `c:numCache` (ECMA-376 Part 1 §21.2.2.121 `numFmt`, §21.2.2.71 `formatCode`, §21.2.2.150
 *  `pt`); an AXIS `c:numFmt` formats the ticks, never the cells. The engine is Excel's (`xlsx-number-format.ts#formatNumberCell`,
 *  « General » included: `37.799999999999997` is « 37.8 »). The same reader serves pptx, xlsx and docx charts
 *  (`pptx-reader/chart.ts#readChartOutcome`), so this is the one place the rule is held; the app's `src/conservation/confluence-page-reading.docx-chart.test.ts`
 *  holds that the independent Word reading of the page says the same. */
import { describe, expect, it } from "vitest";

import type { OpcPackage } from "../ooxml/package";
import { parseXml } from "../ooxml/xml";
import { readChartOutcome } from "./chart";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** One bar series « Share » over the categories « English », « French », with the given cache format and extra series children. */
const chartXml = (opts: { formatCode?: string; values: string[]; dLbls?: string; axisFormat?: string }): string =>
  `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `${opts.dLbls ?? ""}` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>English</c:v></c:pt><c:pt idx="1"><c:v>French</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>${opts.formatCode === undefined ? "" : `<c:formatCode>${opts.formatCode}</c:formatCode>`}` +
  `<c:ptCount val="${opts.values.length}"/>${opts.values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val>` +
  `</c:ser></c:barChart>${opts.axisFormat === undefined ? "" : `<c:valAx><c:axId val="2"/><c:numFmt formatCode="${opts.axisFormat}" sourceLinked="0"/></c:valAx>`}</c:plotArea></c:chart></c:chartSpace>`;

/** The cells of the data table the reader makes of `xml`: the label column, then the series column. */
const rowsOf = async (xml: string): Promise<string[][]> => {
  const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
  const read = await readChartOutcome(pkg, "word/charts/chart1.xml", { picture: false });
  if (read === undefined || read === "withheld") throw new Error(`no chart data: ${String(read)}`);
  return read.data.rows;
};

const labels = (fmt: string, sourceLinked: "0" | "1"): string =>
  `<c:dLbls><c:numFmt formatCode="${fmt}" sourceLinked="${sourceLinked}"/><c:showVal val="1"/></c:dLbls>`;

describe("a chart's table shows the DISPLAYED value of the cache", () => {
  it("🔴 a percent format: the stored 0.80861244019138756 under « 0% » is « 81% », 2.5477707006369428E-2 is « 3% » (chart13 of the SSHRC report)", async () => {
    // Red today: [["English","0.80861244019138756"],["French","2.5477707006369428E-2"]].
    expect(
      await rowsOf(chartXml({ formatCode: "0%", values: ["0.80861244019138756", "2.5477707006369428E-2"] })),
    ).toEqual([
      ["English", "81%"],
      ["French", "3%"],
    ]);
  });

  it("🔴 a number format with decimals: « 0.0% » gives one decimal, « 0.00 » two", async () => {
    expect(
      await rowsOf(chartXml({ formatCode: "0.0%", values: ["0.8282905516111414", "0.17170944838885854"] })),
    ).toEqual([
      ["English", "82.8%"],
      ["French", "17.2%"],
    ]);
    expect(await rowsOf(chartXml({ formatCode: "0.00", values: ["1234.5678", "2"] }))).toEqual([
      ["English", "1234.57"],
      ["French", "2.00"],
    ]);
  });

  it("🔴 « General » shows what Excel's General shows: 37.799999999999997 is « 37.8 » (chart27 of the SSHRC report)", async () => {
    expect(await rowsOf(chartXml({ formatCode: "General", values: ["37.799999999999997", "1840"] }))).toEqual([
      ["English", "37.8"],
      ["French", "1840"],
    ]);
  });

  it("a value that is already what it displays stays as written (cache with no format code, a whole number)", async () => {
    expect(await rowsOf(chartXml({ values: ["1840", "7"] }))).toEqual([
      ["English", "1840"],
      ["French", "7"],
    ]);
  });

  it("🔴 the data labels' own format wins when it is NOT source-linked: « 0% » over the cache's « 0.0% » (chart31 of the SSHRC report: Word labels 83 %)", async () => {
    expect(
      await rowsOf(
        chartXml({
          formatCode: "0.0%",
          values: ["0.8282905516111414", "0.17170944838885854"],
          dLbls: labels("0%", "0"),
        }),
      ),
    ).toEqual([
      ["English", "83%"],
      ["French", "17%"],
    ]);
  });

  it("🔴 a SOURCE-LINKED label format defers to the cache's: « 0% » linked, cache « 0.0% », the cell is « 82.8% »", async () => {
    expect(
      await rowsOf(
        chartXml({
          formatCode: "0.0%",
          values: ["0.8282905516111414", "0.17170944838885854"],
          dLbls: labels("0%", "1"),
        }),
      ),
    ).toEqual([
      ["English", "82.8%"],
      ["French", "17.2%"],
    ]);
  });

  it("an AXIS format never formats the cells: « 0% » on the value axis, « General » cache, 1.05 stays « 1.05 » (chart5 of the SSHRC report)", async () => {
    // Green today and must stay green: the guard against reading the wrong numFmt.
    expect(await rowsOf(chartXml({ formatCode: "General", values: ["1.05", "0.54"], axisFormat: "0%" }))).toEqual([
      ["English", "1.05"],
      ["French", "0.54"],
    ]);
  });
});
