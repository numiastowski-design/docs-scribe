/** A scatter (`c:scatterChart`) or bubble (`c:bubbleChart`) series writes its horizontal values in `c:xVal`, its vertical ones in `c:yVal`
 *  and, for a bubble, the sizes in `c:bubbleSize` (ECMA-376 Part 1 §21.2: `scatterChart`, `bubbleChart`, `xVal`, `yVal`, `bubbleSize`; the
 *  clause numbers are not re-checked in docs/SOURCES-OFFICIELLES.md, which names the standard only). The table publishes them all, as displayed:
 *  - the X values head the table in the left column when no series writes categories (corner cell: the title of the X axis);
 *  - several series with the SAME X (as displayed) share that column; series with different X each get their own (X, Y) pair: a
 *    shared column would have to put points of different X on one row, and a number would be misplaced or lost;
 *  - a bubble's sizes are one more column, header empty (a header of ours would be a word the document never wrote);
 *  - every `c:valAx` is classed by `c:axPos` (`b`/`t` x, `l`/`r` y) when the chart has a scatter group, with a category axis or not.
 *  The independent control reading is held in the app's `src/conservation/confluence-page-reading.docx-chart-scatter.test.ts`. */
import { describe, expect, it } from "vitest";

import type { OpcPackage } from "../ooxml/package";
import { parseXml } from "../ooxml/xml";
import { readChartOutcome } from "./chart";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const title = (text: string): string =>
  `<c:title><c:tx><c:rich><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
const name = (text: string): string =>
  `<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
const pts = (values: readonly string[]): string =>
  `<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}`;
const num = (values: readonly string[], code = "General"): string =>
  `<c:numRef><c:f>S!$A$1</c:f><c:numCache><c:formatCode>${code}</c:formatCode>${pts(values)}</c:numCache></c:numRef>`;
const str = (values: readonly string[]): string =>
  `<c:strRef><c:f>S!$A$1</c:f><c:strCache>${pts(values)}</c:strCache></c:strRef>`;

const xy = (label: string, x: string, y: readonly string[], extra = ""): string =>
  `<c:ser><c:idx val="0"/><c:order val="0"/>${name(label)}<c:xVal>${x}</c:xVal><c:yVal>${num(y)}</c:yVal>${extra}</c:ser>`;

const chart = (groups: string, axes = ""): string =>
  `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>${groups}${axes}</c:plotArea></c:chart></c:chartSpace>`;
const scatter = (...series: string[]): string =>
  `<c:scatterChart><c:scatterStyle val="lineMarker"/>${series.join("")}</c:scatterChart>`;
const axis = (tag: string, pos: string, text: string): string =>
  `<c:${tag}><c:axId val="${pos}"/><c:axPos val="${pos}"/>${title(text)}</c:${tag}>`;

const readOf = async (xml: string) => {
  const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
  return readChartOutcome(pkg, "word/charts/chart1.xml", { picture: false });
};
const dataOf = async (xml: string) => {
  const read = await readOf(xml);
  if (read === undefined || read === "withheld") throw new Error(`no chart data: ${String(read)}`);
  return read;
};

describe("a scatter series publishes its X values and its Y values", () => {
  it("🔴 one series: the left column holds the X, headed by the X axis title; the Y column is headed by the series name", async () => {
    const read = await dataOf(
      chart(
        scatter(xy("Sales", num(["1", "2", "3"]), ["10", "20", "30"])),
        axis("valAx", "b", "Year") + axis("valAx", "l", "Amount"),
      ),
    );
    expect(read.data.headers).toEqual(["Year", "Sales"]);
    expect(read.data.rows).toEqual([
      ["1", "10"],
      ["2", "20"],
      ["3", "30"],
    ]);
    expect(read.axisTitles).toEqual([{ axis: "y", text: "Amount" }]);
  });

  it("🔴 no X axis title: the corner cell stays empty", async () => {
    const read = await dataOf(chart(scatter(xy("Sales", num(["1", "2"]), ["10", "20"]))));
    expect(read.data.headers).toEqual(["", "Sales"]);
    expect(read.data.rows).toEqual([
      ["1", "10"],
      ["2", "20"],
    ]);
  });

  it("🔴 the X values are shown as the cache formats them (0% over 0.5 is 50%), and a text X (`c:strCache`) is kept as written", async () => {
    const formatted = await dataOf(chart(scatter(xy("S", num(["0.5", "0.25"], "0%"), ["1", "2"]))));
    expect(formatted.data.rows).toEqual([
      ["50%", "1"],
      ["25%", "2"],
    ]);
    const text = await dataOf(chart(scatter(xy("S", str(["north", "south"]), ["1", "2"]))));
    expect(text.data.rows).toEqual([
      ["north", "1"],
      ["south", "2"],
    ]);
  });

  it("🔴 a point whose X is blank keeps its row (the Y is never lost)", async () => {
    const x = `<c:numRef><c:f>S!$A$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef>`;
    const read = await dataOf(chart(scatter(xy("S", x, ["10", "20"]))));
    expect(read.data.rows).toEqual([
      ["", "10"],
      ["7", "20"],
    ]);
  });
});

describe("several scatter series", () => {
  it("🔴 the same X (as displayed): one shared X column, one Y column per series", async () => {
    const x = num(["1", "2"]);
    const read = await dataOf(chart(scatter(xy("A", x, ["10", "20"]), xy("B", x, ["30", "40"]))));
    expect(read.data.headers).toEqual(["", "A", "B"]);
    expect(read.data.rows).toEqual([
      ["1", "10", "30"],
      ["2", "20", "40"],
    ]);
  });

  it("🔴 different X: a pair (X, Y) per series, no number lost nor put on another point's row", async () => {
    const read = await dataOf(
      chart(
        scatter(xy("A", num(["1", "2"]), ["10", "20"]), xy("B", num(["5", "6", "7"]), ["30", "40", "50"])),
        axis("valAx", "b", "Year"),
      ),
    );
    expect(read.data.headers).toEqual(["Year", "A", "", "B"]);
    expect(read.data.rows).toEqual([
      ["1", "10", "5", "30"],
      ["2", "20", "6", "40"],
      ["", "", "7", "50"],
    ]);
  });
});

describe("a bubble series", () => {
  it("🔴 the sizes (`c:bubbleSize`) are one more column after the Y, header empty, shown as their cache formats them", async () => {
    const bubble =
      `<c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/>${name("Deals")}<c:xVal>${num(["1", "2"])}</c:xVal><c:yVal>${num(["10", "20"])}</c:yVal>` +
      `<c:bubbleSize>${num(["0.1", "0.2"], "0%")}</c:bubbleSize></c:ser></c:bubbleChart>`;
    const read = await dataOf(chart(bubble, axis("valAx", "b", "Size of firm")));
    expect(read.data.headers).toEqual(["Size of firm", "Deals", ""]);
    expect(read.data.rows).toEqual([
      ["1", "10", "10%"],
      ["2", "20", "20%"],
    ]);
  });
});

describe("axes are classed by their position, with a category axis in the chart or not", () => {
  const combined = (): string => {
    const bar = `<c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>${name("Share")}<c:cat>${str(["English", "French"])}</c:cat><c:val>${num(["5", "7"])}</c:val></c:ser></c:barChart>`;
    const axes =
      axis("valAx", "b", "Horizontal") +
      axis("valAx", "l", "Vertical") +
      axis("valAx", "r", "Secondary") +
      axis("catAx", "b", "Language");
    return chart(bar + scatter(xy("Dots", num(["1", "2"]), ["3", "4"])), axes);
  };

  it("🔴 a combined chart: the value axes at bottom/top are x, left/right y; the category axis stays x and heads the corner", async () => {
    const read = await dataOf(combined());
    expect(read.data.headers?.[0]).toBe("Language");
    expect(read.axisTitles).toEqual([
      { axis: "x", text: "Horizontal" },
      { axis: "y", text: "Vertical" },
      { axis: "y", text: "Secondary" },
    ]);
  });

  it("🔴 a combined chart keeps every number: categories and bar values, then the scatter's X and Y in their own columns", async () => {
    const read = await dataOf(combined());
    expect(read.data.headers).toEqual(["Language", "Share", "", "Dots"]);
    expect(read.data.rows).toEqual([
      ["English", "5", "1", "3"],
      ["French", "7", "2", "4"],
    ]);
  });

  it("counter-test: a horizontal bar chart (no scatter) keeps its value axis as y even at the bottom", async () => {
    const bar = `<c:barChart><c:barDir val="bar"/><c:ser><c:idx val="0"/><c:order val="0"/>${name("Share")}<c:cat>${str(["English"])}</c:cat><c:val>${num(["5"])}</c:val></c:ser></c:barChart>`;
    const read = await dataOf(chart(bar, axis("catAx", "l", "Language") + axis("valAx", "b", "Count")));
    expect(read.axisTitles).toEqual([{ axis: "y", text: "Count" }]);
  });
});

describe("the caps still withhold a scatter", () => {
  it("🔴 a point at or past the point cap in the X values withholds the chart whole", async () => {
    const x = `<c:numRef><c:f>S!$A$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10001"/><c:pt idx="10000"><c:v>1</c:v></c:pt></c:numCache></c:numRef>`;
    expect(await readOf(chart(scatter(xy("S", x, ["1"]))))).toBe("withheld");
  });

  it("🔴 the X columns count in the cell cap: one (X, Y) pair of 9 999 rows is kept, twelve pairs (24 columns = 239 976 cells) are withheld", async () => {
    const rows = Array.from({ length: 9_999 }, (_, i) => String(i));
    expect(await readOf(chart(scatter(xy("S", num(rows), rows))))).not.toBe("withheld");
    const many = chart(
      scatter(...Array.from({ length: 12 }, (_, k) => xy(`S${k}`, num(rows.map((r) => `${r}.${k}`)), rows))),
    );
    expect(await readOf(many)).toBe("withheld");
  });
});

describe("a series name written as a literal", () => {
  it("🔴 `c:tx/c:v` (no cell reference) heads the series' column", async () => {
    const bar = `<c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Bruto</c:v></c:tx><c:cat>${str(["Ano"])}</c:cat><c:val>${num(["5"])}</c:val></c:ser></c:barChart>`;
    expect((await dataOf(chart(bar))).data.headers).toEqual(["", "Bruto"]);
  });
});
