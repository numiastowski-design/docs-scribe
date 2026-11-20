/** Which native charts get a DRAWING (the Forge macro) and with which numbers: the type of the ONE
 *  plot group (ECMA-376 Part 1 §21.2.2: `c:barChart` + `c:barDir` §21.2.2.17 + `c:grouping` §21.2.2.76, `c:lineChart`, `c:pieChart`,
 *  `c:doughnutChart`), and per series the numbers as the table shows them (`81%` gives 81). A chart with no faithful component gets no drawing. */
import { describe, expect, it } from "vitest";

import type { OpcPackage } from "../ooxml/package";
import { parseXml } from "../ooxml/xml";
import { parseDisplayedNumber, readChartOutcome } from "./chart";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const series = (name: string, values: (string | null)[], opts: { format?: string; cats?: string } = {}): string =>
  `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  (opts.cats ??
    `<c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((_, i) => `<c:pt idx="${i}"><c:v>cat${i + 1}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>`) +
  `<c:val><c:numRef><c:f>S!$B$2</c:f><c:numCache>${opts.format === undefined ? "" : `<c:formatCode>${opts.format}</c:formatCode>`}<c:ptCount val="${values.length}"/>` +
  `${values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache></c:numRef></c:val></c:ser>`;

const chartXml = (...groups: string[]): string =>
  `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>${groups.join("")}</c:plotArea></c:chart></c:chartSpace>`;
const group = (tag: string, inner: string, body: string): string => `<c:${tag}>${inner}${body}</c:${tag}>`;

const read = async (xml: string) => {
  const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
  const out = await readChartOutcome(pkg, "word/charts/chart1.xml", { picture: false });
  if (out === undefined || out === "withheld") throw new Error(`no chart: ${String(out)}`);
  return out;
};

describe("the drawing of a chart read", () => {
  it("a vertical clustered bar chart is `bar`, with the series numbers per row", async () => {
    const r = await read(
      chartXml(
        group(
          "barChart",
          '<c:barDir val="col"/><c:grouping val="clustered"/>',
          series("S", ["3", "5.5"]) + series("T", ["1", null]),
        ),
      ),
    );
    expect(r.drawing).toEqual({
      type: "bar",
      values: [
        [3, 5.5],
        [1, null],
      ],
    });
  });

  it.each([
    ['<c:barDir val="bar"/><c:grouping val="clustered"/>', "hBar"],
    ['<c:barDir val="col"/><c:grouping val="stacked"/>', "stackBar"],
    ['<c:barDir val="bar"/><c:grouping val="stacked"/>', "hStackBar"],
    ["", "bar"],
  ])("bar chart header %s is `%s`", async (inner, type) => {
    const r = await read(chartXml(group("barChart", inner, series("S", ["1", "2"]))));
    expect(r.drawing?.type).toBe(type);
  });

  it("a 100%-stacked bar chart has no drawing (absolute bars would be another chart)", async () => {
    const r = await read(
      chartXml(group("barChart", '<c:barDir val="col"/><c:grouping val="percentStacked"/>', series("S", ["1", "2"]))),
    );
    expect(r.drawing).toBeUndefined();
  });

  it("a line chart is `line` when standard, none when stacked", async () => {
    expect(
      (await read(chartXml(group("lineChart", '<c:grouping val="standard"/>', series("S", ["1", "2"]))))).drawing?.type,
    ).toBe("line");
    expect(
      (await read(chartXml(group("lineChart", '<c:grouping val="stacked"/>', series("S", ["1", "2"]))))).drawing,
    ).toBeUndefined();
  });

  it("a pie with one series is `pie`, a doughnut `donut`; two series give none", async () => {
    expect((await read(chartXml(group("pieChart", "", series("S", ["1", "2"]))))).drawing?.type).toBe("pie");
    expect((await read(chartXml(group("doughnutChart", "", series("S", ["1", "2"]))))).drawing?.type).toBe("donut");
    expect(
      (await read(chartXml(group("pieChart", "", series("S", ["1", "2"]) + series("T", ["1", "2"]))))).drawing,
    ).toBeUndefined();
  });

  it("the numbers are the DISPLAYED ones: 0% over 0.8086 is 81", async () => {
    const r = await read(
      chartXml(group("barChart", "", series("S", ["0.80861244019138756", "0.025"], { format: "0%" }))),
    );
    expect(r.data.rows).toEqual([
      ["cat1", "81%"],
      ["cat2", "3%"],
    ]);
    expect(r.drawing?.values).toEqual([[81, 3]]);
  });

  it("a cell the chart shows as no plain number is a gap (null), NEVER the stored value", async () => {
    const unit = await read(chartXml(group("barChart", "", series("S", ["1234.5"], { format: '#,##0.0" kg"' }))));
    expect(unit.drawing?.values).toEqual([[null]]);
    const hidden = await read(chartXml(group("barChart", "", series("S", ["5", "6"], { format: ";;;" }))));
    expect(hidden.data.rows.map((r) => r[1])).toEqual(["", ""]);
    expect(hidden.drawing?.values).toEqual([[null, null]]);
    // A conditional format shows a WORD for 3 (« faible ») and the number for 7: the drawing has a gap where the table has the word.
    const conditional = await read(
      chartXml(group("barChart", "", series("S", ["3", "7"], { format: "[&lt;5]&quot;faible&quot;;0" }))),
    );
    expect(conditional.data.rows.map((r) => r[1])).toEqual(["faible", "7"]);
    expect(conditional.drawing?.values).toEqual([[null, 7]]);
  });

  it("thousands separators: only well-placed commas count (1,234 is 1234; 12,34 and 1,23 are gaps); accounting parentheses are negative", async () => {
    const r = await read(chartXml(group("barChart", "", series("S", ["1234", "-5"], { format: "#,##0;(#,##0)" }))));
    expect(r.data.rows.map((x) => x[1])).toEqual(["1,234", "(5)"]);
    expect(r.drawing?.values).toEqual([[1234, -5]]);
  });

  it("parseDisplayedNumber reads a displayed cell and nothing else", () => {
    expect(["1,234,567", "1234", "-3.5%", "81%", "1,234.5", "(5)", "(1,200)"].map(parseDisplayedNumber)).toEqual([
      1234567, 1234, -3.5, 81, 1234.5, -5, -1200,
    ]);
    expect(
      ["12,34", "1,23", ",5", "1,2345", "", "faible", "5 kg", "1 234", "1e3", "--5", "()"].map(parseDisplayedNumber),
    ).toEqual(Array(11).fill(null));
  });

  it("parseDisplayedNumber reads past one currency sign, before or after the number, as the chart shows it", () => {
    expect(["$5", "$1,234.50", "5 €", "€5", "-$5", "($1,200)", "£3.5", "¥100"].map(parseDisplayedNumber)).toEqual([
      5, 1234.5, 5, 5, -5, -1200, 3.5, 100,
    ]);
    expect(["$", "$$5", "5 € €", "US$5"].map(parseDisplayedNumber)).toEqual(Array(4).fill(null));
  });

  it("an of-pie chart has no drawing (a pie of pie is not a pie)", async () => {
    expect((await read(chartXml(group("ofPieChart", "", series("S", ["1", "2"]))))).drawing).toBeUndefined();
  });

  it.each(["areaChart", "radarChart", "stockChart", "surfaceChart"])("%s has no drawing", async (tag) => {
    expect((await read(chartXml(group(tag, "", series("S", ["1", "2"]))))).drawing).toBeUndefined();
  });

  it("a combined chart (two plot groups) has no drawing", async () => {
    const r = await read(
      chartXml(
        group("barChart", "", series("S", ["1", "2"])),
        group("lineChart", '<c:grouping val="standard"/>', series("T", ["1", "2"])),
      ),
    );
    expect(r.drawing).toBeUndefined();
  });

  it("categories on several levels: no drawing (the table has several label columns)", async () => {
    const cats =
      '<c:cat><c:multiLvlStrRef><c:f>S!$A$2</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:lvl>' +
      '<c:lvl><c:pt idx="0"><c:v>G</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>';
    const r = await read(chartXml(group("barChart", "", series("S", ["1", "2"], { cats }))));
    expect(r.drawing).toBeUndefined();
  });

  // ECMA-376 Part 1 §21.2.2.130 (`c:orientation`, child of `c:scaling` §21.2.2.160 of `c:catAx` §21.2.2.25 or `c:dateAx` §21.2.2.39) and
  // §21.2.3.30 (ST_Orientation: `minMax`, `maxMin`). Google Sheets writes `maxMin` on the category axis of a horizontal bar chart, so
  // that the first row is the top bar (fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx).
  const axis = (tag: string, orientation: string | undefined): string =>
    `<c:${tag}><c:axId val="1"/><c:scaling>${orientation === undefined ? "" : `<c:orientation val="${orientation}"/>`}</c:scaling></c:${tag}>`;
  const bars = group("barChart", '<c:barDir val="bar"/><c:grouping val="clustered"/>', series("S", ["1", "2"]));

  it.each(["catAx", "dateAx"])(
    "a category axis (`c:%s`) running `maxMin` gives a drawing whose categories are reversed; the values stay in table order",
    async (tag) => {
      const r = await read(chartXml(bars, axis(tag, "maxMin"), axis("valAx", undefined)));
      expect(r.drawing).toEqual({ type: "hBar", values: [[1, 2]], categoriesReversed: true });
    },
  );

  it.each([
    ["`minMax`", axis("catAx", "minMax") + axis("valAx", "maxMin")],
    ["no orientation", axis("catAx", undefined)],
    ["no axis", ""],
  ])(
    "a category axis with %s is not reversed (a reversed VALUE axis says nothing of the categories)",
    async (_label, axes) => {
      const r = await read(chartXml(bars, axes));
      expect(r.drawing).toEqual({ type: "hBar", values: [[1, 2]] });
    },
  );

  it("a pie has no category axis: never reversed, even next to a stray reversed axis", async () => {
    const r = await read(chartXml(group("pieChart", "", series("S", ["1", "2"])), axis("catAx", "maxMin")));
    expect(r.drawing).toEqual({ type: "pie", values: [[1, 2]] });
  });
});
