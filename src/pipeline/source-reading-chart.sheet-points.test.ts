/** The structure control's own reading of a cache-less workbook chart from its sheet (`source-reading-chart.ts`, `source-chart-reference.ts`)
 *  costs no more than the chart's own bounds allow: past the series bound nothing is resolved, a repeated address is resolved once, and a blank
 *  column under many kept rows is never walked row by row. Lookups are counted, never timed. */
import { describe, expect, it } from "vitest";

import { parseXml } from "../extractor/ooxml/xml";
import {
  keptCellsResolver,
  type KeptChartCell,
  type SheetCells,
  type SheetCellsResolver,
} from "./source-chart-reference";
import { chartReadOf } from "./source-reading-chart";
import type { N } from "./source-reading-package";

const C =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const name = (address: string): string => `<c:tx><c:strRef><c:f>${address}</c:f></c:strRef></c:tx>`;
const bar = (label: string, categories: string, values: string): string =>
  `<c:ser><c:idx val="0"/><c:order val="0"/>${name(label)}<c:cat><c:strRef><c:f>${categories}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${values}</c:f></c:numRef></c:val></c:ser>`;
const barChart = (...series: string[]): N =>
  parseXml(
    `<?xml version="1.0"?><c:chartSpace ${C}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series.join("")}</c:barChart></c:plotArea></c:chart></c:chartSpace>`,
  ) as unknown as N;

/** A resolver that counts the addresses it is asked for and answers every one with one text cell. */
const counting = (): { resolver: SheetCellsResolver; asked: string[] } => {
  const asked: string[] = [];
  return {
    asked,
    resolver: (address) => (
      asked.push(address),
      Object.assign([{ index: 0, text: "x", numeric: false }], { extent: 1 })
    ),
  };
};

describe("a resolver's answer always tells its extent", () => {
  it("🔴 an answer without an extent does not type: the plan counts before any cell is read, never by reading", () => {
    // @ts-expect-error `extent` is required on every answer.
    const missing: SheetCells = [{ index: 0, text: "x", numeric: false }];
    expect([...missing]).toHaveLength(1);
  });
});

describe("the control's cost of reading a cache-less chart from its sheet", () => {
  it("🔴 a chart past the series bound is withheld, as from its cache, and no reference is resolved", () => {
    const { resolver, asked } = counting();
    const chart = barChart(
      ...Array.from({ length: 256 }, (_, index) =>
        bar(`S!$B$${index + 1}`, "S!$A$2:$A$4", `S!$C$${index + 1}:$C$${index + 9}`),
      ),
    );
    expect(chartReadOf(chart, undefined, resolver).shape).toBeUndefined();
    expect(chartReadOf(chart).shape).toBeUndefined();
    expect(asked).toEqual([]);
  });

  it("🔴 an address repeated by several series is resolved once per chart", () => {
    const { resolver, asked } = counting();
    chartReadOf(
      barChart(
        bar("S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4"),
        bar("S!$C$1", "S!$A$2:$A$4", "S!$C$2:$C$4"),
        bar("S!$D$1", "S!$A$2:$A$4", "S!$D$2:$D$4"),
      ),
      undefined,
      resolver,
    );
    expect(asked.filter((address) => address === "S!$A$2:$A$4")).toHaveLength(1);
    expect(asked).toHaveLength(7);
  });

  it("🔴 a chart that holds literal data anywhere keeps its cache: nothing is read from the sheet", () => {
    const literal = (tag: "c:numLit" | "c:strLit"): string =>
      `<${tag}><c:ptCount val="1"/><c:pt idx="0"><c:v>7</c:v></c:pt></${tag}>`;
    const values = barChart(
      `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f></c:strRef></c:cat><c:val>${literal("c:numLit")}</c:val></c:ser>`,
    );
    const categories = barChart(
      `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat>${literal("c:strLit")}</c:cat><c:val><c:numRef><c:f>S!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>`,
    );
    for (const chart of [values, categories]) {
      const { resolver, asked } = counting();
      expect(chartReadOf(chart, undefined, resolver).shape).toEqual(chartReadOf(chart).shape);
      expect(asked).toEqual([]);
    }
  });

  it("🔴 a blank column under thousands of kept rows costs at most (kept cells) + references x (point bound + 1) cell reads", () => {
    // Column A (1) holds text on 5 000 rows; column B (2) holds only blanks there (kept cells that show nothing); 200 series each read B over
    // the whole height, from a different first row (no address repeats).
    const keptRows = 5_000;
    const seriesCount = 200;
    const pointBound = 10;
    let reads = 0;
    const counted = (rowNum: number, colNum: number, shown: string): KeptChartCell => ({
      rowNum,
      colNum,
      numeric: false,
      get text(): string {
        reads++;
        return shown;
      },
    });
    const cells = Array.from({ length: keptRows }, (_, index) => [
      counted(index + 1, 1, "label"),
      counted(index + 1, 2, " "),
    ]).flat();
    const chart = barChart(
      ...Array.from({ length: seriesCount }, (_, index) => bar("S!$C$1", "S!$A$1", `S!$B$${index + 1}:$B$1048576`)),
    );
    chartReadOf(chart, { chartPoints: pointBound }, keptCellsResolver("S", cells));
    // Distinct addresses: 200 value ranges, one name, one category cell.
    const references = seriesCount + 2;
    expect(reads).toBeLessThanOrEqual(cells.length + references * (pointBound + 1));
  });
});

describe("the control withholds a chart whose references can yield more cells than the bound, before any cell is read", () => {
  /** Column A (1) holds the numbers 1..`rows`; `handed` counts the cells the resolver gives out. */
  const filledColumn = (rows: number): { resolver: SheetCellsResolver; handed: () => number } => {
    let handed = 0;
    const inner = keptCellsResolver(
      "S",
      Array.from({ length: rows }, (_, index) => ({
        rowNum: index + 1,
        colNum: 1,
        text: String(index + 1),
        numeric: true,
      })),
    );
    const resolver: SheetCellsResolver = (address) => {
      const found = inner(address);
      if (found === undefined) return undefined;
      return {
        *[Symbol.iterator]() {
          for (const cell of found) {
            handed++;
            yield cell;
          }
        },
        extent: found.extent,
      };
    };
    return { resolver, handed: () => handed };
  };

  it("🔴 series that each name another range of one filled column add up past the cell bound: withheld, no cell handed out", () => {
    // One address per series (categories and values alike), counted once: 1 + 300 cells; the fourth series brings the count to 1 204.
    const chart = barChart(
      ...Array.from({ length: 40 }, (_, index) => {
        const range = `S!$A$1:$A$${300 + index}`;
        return `<c:ser><c:idx val="0"/><c:order val="0"/>${name("S!$B$1")}<c:cat><c:strRef><c:f>${range}</c:f></c:strRef><c:numRef><c:f>${range}</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>${range}</c:f></c:numRef></c:val></c:ser>`;
      }),
    );
    const { resolver, handed } = filledColumn(300);
    expect(chartReadOf(chart, { chartCells: 1_000 }, resolver)).toEqual({ words: [], shape: undefined });
    expect(handed()).toBe(0);
  });

  it("🔴 at the cell bound the chart is read from its sheet, one cell past it the chart is withheld", () => {
    // One name cell (1) + 300 categories + 300 values = 601: the values write the same range another way, a second address read a second
    // time. The table holds 2 x 300 = 600 cells: at 600 only the references' count refuses.
    const chart = barChart(bar("S!$B$1", "S!$A$1:$A$300", "S!A1:A300"));
    expect(chartReadOf(chart, { chartCells: 601 }, filledColumn(300).resolver).shape).toEqual({
      rows: 301,
      columns: 2,
    });
    const { resolver, handed } = filledColumn(300);
    expect(chartReadOf(chart, { chartCells: 600 }, resolver)).toEqual({ words: [], shape: undefined });
    expect(handed()).toBe(0);
  });

  it("🔴 an address repeated by every series is counted once, as it is read once", () => {
    // 3 name cells + 300 cells of one address = 303 counted (a count per series gave 1 803); the table holds (1 + 3) x 300 = 1 200 cells.
    const chart = barChart(...Array.from({ length: 3 }, () => bar("S!$B$1", "S!$A$1:$A$300", "S!$A$1:$A$300")));
    expect(chartReadOf(chart, { chartCells: 1_200 }, filledColumn(300).resolver).shape).toEqual({
      rows: 301,
      columns: 4,
    });
  });

  it("🔴 a reference whose last cell sits at or past the point bound withholds the chart, nothing read", () => {
    const chart = barChart(bar("S!$B$1", "S!$A$1:$A$300", "S!$A$1:$A$300"));
    expect(chartReadOf(chart, { chartPoints: 300 }, filledColumn(300).resolver).shape).toEqual({
      rows: 301,
      columns: 2,
    });
    const { resolver, handed } = filledColumn(300);
    expect(chartReadOf(chart, { chartPoints: 299 }, resolver)).toEqual({ words: [], shape: undefined });
    expect(handed()).toBe(0);
  });
});
