/** The closed grammar of the address a chart reads from its sheet (`xlsx-chart-reference.ts#parseChartReference`,
 *  docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md): every accepted form, every refused form. The forms are the ones Excel and Google Sheets write
 *  (fixtures/xlsx-reel/17-…, xlsx-heading-reel/03-…, 34-…) and the ones the grammar leaves out on purpose. */
import { describe, expect, it } from "vitest";

import { parseChartReference, sheetChartResolver, type ChartGrid } from "./xlsx-chart-reference";

const range = (firstRow: number, lastRow: number, firstColumn: number, lastColumn: number) => ({
  firstRow,
  lastRow,
  firstColumn,
  lastColumn,
});

describe("parseChartReference: accepted forms", () => {
  const accepted: [string, string, ReturnType<typeof range>][] = [
    ["a single cell, bare sheet name, absolute", "Bar!$B$1", range(0, 0, 1, 1)],
    ["a column range, as Google writes it", "Bar!$B$2:$B$50", range(1, 49, 1, 1)],
    ["a row range, quoted sheet name", "'Base LA Summary Data'!$C$10:$G$10", range(9, 9, 2, 6)],
    ["a relative address", "Bar!B2:B5", range(1, 4, 1, 1)],
    ["a mixed absolute/relative address", "Bar!$B2:B$5", range(1, 4, 1, 1)],
    ["a quoted name that needs no quotes", "'Bar'!A1", range(0, 0, 0, 0)],
    ["an apostrophe doubled inside a quoted name", "'Team''s'!A1:A3", range(0, 2, 0, 0)],
    ["a one-cell range", "Bar!A1:A1", range(0, 0, 0, 0)],
    ["the last cell Excel has", "Bar!XFD1048576", range(1_048_575, 1_048_575, 16_383, 16_383)],
    ["a sheet name that differs only by case", "BAR!A1", range(0, 0, 0, 0)],
  ];
  for (const [title, reference, expected] of accepted) {
    it(`🔴 ${title}: ${reference}`, () => {
      const sheet = reference.startsWith("'Base")
        ? "Base LA Summary Data"
        : reference.startsWith("'Team")
          ? "Team's"
          : "Bar";
      expect(parseChartReference(reference, sheet)).toEqual(expected);
    });
  }

  it("🔴 a bare name with accents, held by the sheet of that name", () => {
    expect(parseChartReference("Résumé!B2:B4", "Résumé")).toEqual(range(1, 3, 1, 1));
  });

  it("🔴 a quoted name with spaces and accents", () => {
    expect(parseChartReference("'Feuille Été'!A1", "Feuille Été")).toEqual(range(0, 0, 0, 0));
  });
});

describe("parseChartReference: refused forms (each one gives undefined)", () => {
  const refused: [string, string][] = [
    ["no sheet name", "$B$2:$B$50"],
    ["no sheet name, one cell", "B2"],
    ["another sheet", "Other!$B$2:$B$5"],
    ["another sheet, quoted", "'Other sheet'!B2"],
    ["an external reference", "[1]Bar!$B$2:$B$5"],
    ["an external reference, quoted", "'[1]Bar'!B2"],
    ["a defined name", "SalesFigures"],
    ["several areas", "(Bar!$B$2:$B$3,Bar!$B$5:$B$6)"],
    ["several areas without parentheses", "Bar!B2:B3,Bar!B5:B6"],
    ["a whole column", "Bar!$B:$B"],
    ["a whole row", "Bar!$2:$2"],
    ["a 3-D range", "Bar:Other!B2"],
    ["a 3-D range, quoted", "'Bar:Other'!B2"],
    ["a structured table reference", "Table1[Sales]"],
    ["a structured table reference behind a sheet", "Bar!Table1[Sales]"],
    ["a two-dimensional range", "Bar!A1:C5"],
    ["a range written from its bottom-right corner (column)", "Bar!B5:B2"],
    ["a range written from its bottom-right corner (row)", "Bar!G10:C10"],
    ["a formula", "=Bar!B2"],
    ["a function", "SUM(Bar!B2:B5)"],
    ["a space around the sheet name", " Bar!B2"],
    ["a space inside an address", "Bar!B2 :B5"],
    ["a lowercase column", "Bar!b2"],
    ["a row past the last one", "Bar!A1048577"],
    ["a column past the last one", "Bar!XFE1"],
    ["row zero", "Bar!A0"],
    ["a leading zero in the row", "Bar!A01"],
    ["a third cell in the range", "Bar!A1:A2:A3"],
    ["an empty address after the sheet", "Bar!"],
    ["an empty reference", ""],
    ["a missing sheet name before the mark", "!B2"],
    ["an unterminated quote", "'Bar!B2"],
  ];
  for (const [title, reference] of refused) {
    it(`🔴 ${title}: ${JSON.stringify(reference)}`, () => {
      expect(parseChartReference(reference, "Bar")).toBeUndefined();
    });
  }
});

describe("sheetChartResolver: which cells give a point", () => {
  // Zero-based: B1 "Period 1", A2 "Team 1", B2 24, B4 12; row 3 (A3, B3) is empty.
  const cells: Record<string, { text: string; number?: number; code?: string }> = {
    "0,1": { text: "Period 1" },
    "1,0": { text: "Team 1" },
    "1,1": { text: "24", number: 24, code: "0" },
    "3,1": { text: "12", number: 12 },
  };
  const grid: ChartGrid = {
    filledCells: [
      [0, 1],
      [1, 0],
      [1, 1],
      [3, 1],
    ],
    textAt: (row, column) => cells[`${row},${column}`]?.text ?? "",
    numberAt: (row, column) => {
      const cell = cells[`${row},${column}`];
      return cell?.number === undefined ? undefined : { value: cell.number, formatCode: cell.code };
    },
  };
  const resolve = sheetChartResolver("Bar", grid);

  it("🔴 a column range gives the filled cells at their place in the range; empty cells give none", () => {
    const points = [...resolve("Bar!$B$2:$B$50")!];
    expect(points.map((point) => [point.index, point.text, point.stored])).toEqual([
      [0, "24", "24"],
      [2, "12", "12"],
    ]);
    expect(points[0]!.formatCode).toBe("0");
  });

  it("🔴 a text cell has no stored number", () => {
    expect([...resolve("Bar!$A$2:$A$4")!]).toEqual([{ index: 0, text: "Team 1" }]);
  });

  it("🔴 a row range gives the cells of that row", () => {
    expect([...resolve("Bar!$A$2:$C$2")!].map((point) => [point.index, point.text])).toEqual([
      [0, "Team 1"],
      [1, "24"],
    ]);
  });

  it("🔴 a refused address gives undefined, not an empty list", () => {
    expect(resolve("Other!$B$2:$B$5")).toBeUndefined();
    expect(resolve("Bar!A1:C5")).toBeUndefined();
  });
});
