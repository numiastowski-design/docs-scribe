/** The chart reader, given a `resolver` (only the workbook reader gives one), reads a chart that carries no value in any reference from the
 *  sheet's points (docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md): same table, same bounds, same drawing as cached values. Charts are built here with
 *  a fake resolver, so each rule is held alone; the real resolver and the real files are held in `extractor/xlsx.chart-from-sheet.test.ts`. */
import { describe, expect, it } from "vitest";

import type { ReaderLimits } from "doc-ir";

import type { OpcPackage } from "../ooxml/package";
import { parseXml } from "../ooxml/xml";
import { sheetChartResolver, type ChartCellPoint, type ChartPoints, type ChartResolver } from "../xlsx-chart-reference";
import { readChartOutcome } from "./chart";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const cached = (address: string): string => `<c:f>${address}</c:f>`;
const name = (address: string): string => `<c:tx><c:strRef>${cached(address)}</c:strRef></c:tx>`;
const cat = (address: string): string => `<c:cat><c:strRef>${cached(address)}</c:strRef></c:cat>`;
const val = (address: string): string => `<c:val><c:numRef>${cached(address)}</c:numRef></c:val>`;
const bar = (label: string, categories: string, values: string, extra = ""): string =>
  `<c:ser><c:idx val="0"/><c:order val="0"/>${name(label)}${extra}${cat(categories)}${val(values)}</c:ser>`;
const chartOf = (group: string): string =>
  `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea>${group}</c:plotArea></c:chart></c:chartSpace>`;
const barChart = (...series: string[]): string =>
  chartOf(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series.join("")}</c:barChart>`);

const read = async (xml: string, resolver: ChartResolver | undefined) => {
  const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
  return readChartOutcome(pkg, "xl/charts/chart1.xml", { picture: false, resolver });
};

const text = (index: number, shown: string): ChartCellPoint => ({ index, text: shown });
const number = (index: number, value: number, formatCode?: string): ChartCellPoint => ({
  index,
  text: String(value),
  stored: String(value),
  formatCode,
});

/** Fixed points, with the extent a sheet would give them: the dense length, up to the last point. */
const fixed = (points: readonly ChartCellPoint[]): ChartPoints =>
  Object.assign([...points], { extent: points.length === 0 ? 0 : points[points.length - 1]!.index + 1 });

/** A resolver over fixed answers; an address not listed is refused. */
const answers =
  (table: Record<string, ChartCellPoint[]>): ChartResolver =>
  (address) => {
    const points = table[address];
    return points === undefined ? undefined : fixed(points);
  };

describe("a chart with no value in any reference is read from the resolver", () => {
  const sheet = answers({
    "S!$B$1": [text(0, "Sales")],
    "S!$A$2:$A$4": [text(0, "North"), text(1, "South"), text(2, "East")],
    "S!$B$2:$B$4": [number(0, 10), number(1, 20), number(2, 30)],
  });
  const chart = barChart(bar("S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4"));

  it("🔴 the series name, the categories and the numbers fill the table and the drawing", async () => {
    const result = await read(chart, sheet);
    expect(result).toMatchObject({
      data: {
        headers: ["", "Sales"],
        rows: [
          ["North", "10"],
          ["South", "20"],
          ["East", "30"],
        ],
      },
      drawing: { type: "bar", values: [[10, 20, 30]] },
    });
  });

  it("🔴 without a resolver the chart is what the file says: no point", async () => {
    expect(await read(chart, undefined)).toMatchObject({ data: { rows: [] } });
  });

  it("🔴 a text cell in the values gives no point (a gap), and the row keeps its category", async () => {
    const withText = answers({
      "S!$B$1": [text(0, "Sales")],
      "S!$A$2:$A$4": [text(0, "North"), text(1, "South"), text(2, "East")],
      "S!$B$2:$B$4": [number(0, 10), text(1, "n/a"), number(2, 30)],
    });
    expect(await read(chart, withText)).toMatchObject({
      data: {
        rows: [
          ["North", "10"],
          ["South", ""],
          ["East", "30"],
        ],
      },
      drawing: { values: [[10, null, 30]] },
    });
  });

  it("🔴 the cell's own format shows the number, as the chart displays it", async () => {
    const percent = answers({
      "S!$B$1": [text(0, "Share")],
      "S!$A$2:$A$2": [text(0, "North")],
      "S!$B$2:$B$2": [number(0, 0.25, "0%")],
    });
    expect(await read(barChart(bar("S!$B$1", "S!$A$2:$A$2", "S!$B$2:$B$2")), percent)).toMatchObject({
      data: { rows: [["North", "25%"]] },
      drawing: { values: [[25]] },
    });
  });

  it("🔴 data labels that do not follow the source decide the format, not the cell", async () => {
    const labelled = `<c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/></c:dLbls>`;
    const percent = answers({
      "S!$B$1": [text(0, "Share")],
      "S!$A$2:$A$2": [text(0, "North")],
      "S!$B$2:$B$2": [number(0, 0.25, "0%")],
    });
    expect(await read(barChart(bar("S!$B$1", "S!$A$2:$A$2", "S!$B$2:$B$2", labelled)), percent)).toMatchObject({
      data: { rows: [["North", "25.0%"]] },
    });
  });

  it("🔴 a point only at index 2 keeps its row, the cells before it stay empty", async () => {
    const sparse = answers({
      "S!$B$1": [text(0, "Sales")],
      "S!$A$2:$A$4": [text(2, "East")],
      "S!$B$2:$B$4": [number(2, 30)],
    });
    expect(await read(chart, sparse)).toMatchObject({ data: { rows: [["East", "30"]] } });
  });

  it("🔴 a reference the resolver refuses keeps the chart as the file wrote it, whole", async () => {
    const refusing = answers({ "S!$B$1": [text(0, "Sales")], "S!$A$2:$A$4": [text(0, "North")] });
    expect(await read(chart, refusing)).toMatchObject({ data: { headers: ["", ""], rows: [] } });
  });

  it("🔴 a chart with a multi-level category reference keeps its cache", async () => {
    const multi = barChart(
      `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat><c:multiLvlStrRef><c:f>S!$A$2:$B$4</c:f></c:multiLvlStrRef></c:cat>${val("S!$B$2:$B$4")}</c:ser>`,
    );
    expect(await read(multi, sheet)).toMatchObject({ data: { rows: [] } });
  });

  it("🔴 a value cached anywhere in the chart, even in a title reference, keeps the cache for every reference", async () => {
    const titled = chart.replace(
      "<c:chart>",
      '<c:chart><c:title><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Title</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>',
    );
    expect(await read(titled, sheet)).toMatchObject({ title: "Title", data: { rows: [] } });
  });

  it("🔴 a chart that holds literal data anywhere keeps its cache: nothing is read from the sheet", async () => {
    const literal = (tag: "c:numLit" | "c:strLit"): string =>
      `<${tag}><c:ptCount val="1"/><c:pt idx="0"><c:v>7</c:v></c:pt></${tag}>`;
    for (const [where, xml] of [
      [
        "values",
        barChart(
          `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}${cat("S!$A$2:$A$4")}<c:val>${literal("c:numLit")}</c:val></c:ser>`,
        ),
      ],
      [
        "categories",
        barChart(
          `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat>${literal("c:strLit")}</c:cat>${val("S!$B$2:$B$4")}</c:ser>`,
        ),
      ],
    ] as const) {
      const asked: string[] = [];
      const result = await read(xml, (address) => (asked.push(address), sheet(address)));
      expect(asked, where).toEqual([]);
      expect(result, where).toMatchObject({ data: { rows: [] } });
    }
  });

  it("🔴 points of a scatter series come from the sheet: X as the left column, Y under the series name", async () => {
    const scatter = chartOf(
      `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:xVal><c:numRef>${cached("S!$A$2:$A$3")}</c:numRef></c:xVal><c:yVal><c:numRef>${cached("S!$B$2:$B$3")}</c:numRef></c:yVal></c:ser></c:scatterChart>`,
    );
    const xy = answers({
      "S!$B$1": [text(0, "Sales")],
      "S!$A$2:$A$3": [number(0, 1), number(1, 2)],
      "S!$B$2:$B$3": [number(0, 10), number(1, 20)],
    });
    expect(await read(scatter, xy)).toMatchObject({
      data: {
        headers: ["", "Sales"],
        rows: [
          ["1", "10"],
          ["2", "20"],
        ],
      },
    });
  });
});

describe("a resolver's answer always tells its extent", () => {
  it("🔴 an answer without an extent does not type: the plan counts before any cell is read, never by reading", () => {
    // @ts-expect-error `extent` is required on every answer.
    const missing: ChartPoints = [text(0, "x")];
    expect([...missing]).toHaveLength(1);
  });
});

describe("a range is never read past the point bound", () => {
  /** A resolver whose ranges never end, each declaring `extent`; `pulled` counts the points handed out. */
  const endless = (extent: number): { resolver: ChartResolver; pulled: () => number } => {
    let pulled = 0;
    const resolver: ChartResolver = (address) =>
      address === "S!$B$1"
        ? fixed([text(0, "Sales")])
        : {
            *[Symbol.iterator]() {
              for (let index = 0; ; index++) {
                if (++pulled > 50_000) throw new Error("the range was read far past the bound");
                yield text(index, "x");
              }
            },
            extent,
          };
    return { resolver, pulled: () => pulled };
  };
  const chart = barChart(bar("S!$B$1", "S!$A$2:$A$99999", "S!$B$2:$B$99999"));

  it("🔴 a range filled to its end, far past the point bound, withholds the chart and is never read", async () => {
    const { resolver, pulled } = endless(99_998);
    expect(await read(chart, resolver)).toBe("withheld");
    expect(pulled()).toBe(0);
  });

  it("🔴 a resolver whose extent understates its points is read no further than the point bound, and the chart is withheld", async () => {
    const { resolver, pulled } = endless(2);
    expect(await read(chart, resolver)).toBe("withheld");
    // MAX_CHART_POINTS is 10 000: the first point at or past it is the last one pulled for each of the two ranges.
    expect(pulled()).toBeLessThanOrEqual(2 * 10_001);
    expect(pulled()).toBeGreaterThan(10_000);
  });
});

describe("the cost of reading a cache-less chart from its sheet is bounded by the chart's own bounds", () => {
  /** A resolver that counts the addresses it is asked for and answers every one with one text point. */
  const counting = (): { resolver: ChartResolver; asked: string[] } => {
    const asked: string[] = [];
    return { asked, resolver: (address) => (asked.push(address), fixed([text(0, "x")])) };
  };

  it("🔴 a chart past the series bound is withheld, as from its cache, and no reference is resolved", async () => {
    const { resolver, asked } = counting();
    const series = Array.from({ length: 256 }, (_, index) =>
      bar(`S!$B$${index + 1}`, "S!$A$2:$A$4", `S!$C$${index + 1}:$C$${index + 9}`),
    );
    expect(await read(barChart(...series), resolver)).toBe("withheld");
    expect(await read(barChart(...series), undefined)).toBe("withheld");
    expect(asked).toEqual([]);
  });

  it("🔴 an address repeated by several series (Google repeats its categories) is resolved once per chart", async () => {
    const { resolver, asked } = counting();
    await read(
      barChart(
        bar("S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4"),
        bar("S!$C$1", "S!$A$2:$A$4", "S!$C$2:$C$4"),
        bar("S!$D$1", "S!$A$2:$A$4", "S!$D$2:$D$4"),
      ),
      resolver,
    );
    expect(asked.filter((address) => address === "S!$A$2:$A$4")).toHaveLength(1);
    expect(asked).toHaveLength(7);
  });

  it("🔴 a blank column under thousands of kept rows costs at most (filled cells) + 2 x references x (point bound + 1) cell lookups", async () => {
    // Column A holds text on 5 000 rows; column B is blank; 200 series each read B over the whole height, from a different first row (no
    // address repeats), under a point bound of 10.
    const filledRows = 5_000;
    const seriesCount = 200;
    const pointBound = 10;
    let lookups = 0;
    const filledCells = Array.from({ length: filledRows }, (_, row) => [row, 0] as const);
    const grid = {
      filledCells,
      textAt: (row: number, column: number): string => (lookups++, column === 0 && row < filledRows ? "label" : ""),
      numberAt: (): undefined => void lookups++,
    };
    const series = Array.from({ length: seriesCount }, (_, index) =>
      bar("S!$C$1", "S!$A$1:$A$1", `S!$B$${index + 1}:$B$1048576`),
    );
    const pkg = { readXml: async () => parseXml(barChart(...series)) } as unknown as OpcPackage;
    await readChartOutcome(pkg, "xl/charts/chart1.xml", {
      picture: false,
      resolver: sheetChartResolver("S", grid),
      limits: { chartPoints: pointBound },
    });
    // Distinct addresses: 200 value ranges, one name, one category cell. One lookup per filled cell builds the sheet's lines once; each cell a
    // walk examines is a point and costs two (its text, its number).
    const references = seriesCount + 2;
    expect(lookups).toBeLessThanOrEqual(filledRows + 2 * references * (pointBound + 1));
  });
});

describe("a chart whose references can yield more cells than the bound is withheld, decided before any cell is read", () => {
  /** Column A holds the numbers 1..`rows` (zero-based rows 0..rows-1). `handed` counts the points the resolver gives out (one number lookup per
   *  point). */
  const filledColumn = (rows: number): { resolver: ChartResolver; handed: () => number } => {
    let handed = 0;
    const resolver = sheetChartResolver("S", {
      filledCells: Array.from({ length: rows }, (_, row) => [row, 0] as const),
      textAt: (row, column) => (column === 0 && row < rows ? String(row + 1) : ""),
      numberAt: (row, column) => (
        handed++,
        column === 0 && row < rows ? { value: row + 1, formatCode: undefined } : undefined
      ),
    });
    return { resolver, handed: () => handed };
  };
  const outcome = async (xml: string, resolver: ChartResolver | undefined, limits: ReaderLimits | undefined) => {
    const pkg = { readXml: async () => parseXml(xml) } as unknown as OpcPackage;
    return readChartOutcome(pkg, "xl/charts/chart1.xml", { picture: false, resolver, limits });
  };

  it("🔴 series that each name another range of one filled column add up past the cell bound: withheld, no point handed out", async () => {
    // Every series names its own range of column A for its categories (a string AND a number reference side by side) and its values: one
    // address per series, counted once, 1 + 300 cells. The fourth series brings the count to 1 204, over 1 000.
    const series = Array.from({ length: 40 }, (_, index) => {
      const range = `S!$A$1:$A$${300 + index}`;
      return `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat><c:strRef>${cached(range)}</c:strRef><c:numRef>${cached(range)}</c:numRef></c:cat>${val(range)}</c:ser>`;
    });
    const chart = barChart(...series);
    const limits = { chartCells: 1_000 };
    const { resolver, handed } = filledColumn(300);
    expect(await outcome(chart, resolver, limits)).toBe("withheld");
    expect(handed()).toBe(0);
  });

  it("🔴 at the cell bound the chart is drawn, one cell past it the chart is withheld", async () => {
    // One name cell (1) + 300 categories + 300 values = 601 cells the references can yield: the values write the same range another way, so
    // it is a second address, read a second time. The table holds 2 x 300 = 600 cells: at 600 only the references' count refuses.
    const chart = barChart(bar("S!$B$1", "S!$A$1:$A$300", "S!A1:A300"));
    expect(await outcome(chart, filledColumn(300).resolver, { chartCells: 601 })).toMatchObject({
      data: { rows: { length: 300 } },
    });
    const { resolver, handed } = filledColumn(300);
    expect(await outcome(chart, resolver, { chartCells: 600 })).toBe("withheld");
    expect(handed()).toBe(0);
  });

  it("🔴 a reference whose last point sits at or past the point bound withholds the chart, nothing read", async () => {
    const chart = barChart(bar("S!$B$1", "S!$A$1:$A$300", "S!$A$1:$A$300"));
    expect(await outcome(chart, filledColumn(300).resolver, { chartPoints: 300 })).toMatchObject({
      data: { rows: { length: 300 } },
    });
    const { resolver, handed } = filledColumn(300);
    expect(await outcome(chart, resolver, { chartPoints: 299 })).toBe("withheld");
    expect(handed()).toBe(0);
  });

  it("🔴 an address repeated by every series (Google repeats its categories) is counted once, as it is read once", async () => {
    // Three series, each naming column A for its categories AND its values: 3 name cells + 300 cells of one address = 303 counted, where a
    // count per series gave 3 x 601 = 1 803. The table itself holds (1 label + 3 series) x 300 = 1 200 cells: the bound is set there.
    const chart = barChart(...Array.from({ length: 3 }, () => bar("S!$B$1", "S!$A$1:$A$300", "S!$A$1:$A$300")));
    expect(await outcome(chart, filledColumn(300).resolver, { chartCells: 1_200 })).toMatchObject({
      data: { rows: { length: 300 } },
    });
  });

  it("🔴 a range is clipped to the last filled cell of its line: what it can yield, without reading a cell", () => {
    const { resolver, handed } = filledColumn(4);
    expect(resolver("S!$A$2:$A$50")?.extent).toBe(3);
    expect(resolver("S!$A$1:$A$1048576")?.extent).toBe(4);
    expect(resolver("S!$B$1:$B$50")?.extent).toBe(0);
    expect(resolver("S!$A$1:$C$1")?.extent).toBe(1);
    expect(handed()).toBe(0);
  });
});
