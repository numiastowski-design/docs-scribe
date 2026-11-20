/** A workbook chart whose file wrote no cached values is drawn from the sheet that holds it (docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md).
 *  Witness: fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx, Google's export: `Bar!A1:C5` holds « Period 1 »/« Period 2 » over four
 *  teams, the chart names `Bar!$B$2:$B$50` and `Bar!$C$2:$C$50` and caches nothing. The variants are that file with one part changed in memory. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractXlsx } from "./xlsx";

vi.setConfig({ testTimeout: 60_000 });

type Chart = Extract<IRBlock, { kind: "chart" }>;

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "../../fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx");
const SHEET = "xl/worksheets/sheet1.xml";
const CHART = "xl/charts/chart1.xml";

/** The fixture with `edits` applied to the text of its parts: `[part, from, to]`, each `from` present exactly once (a case that changes nothing would pass for the wrong reason). */
const patched = async (...edits: [part: string, from: string, to: string][]): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
  for (const [part, from, to] of edits) {
    const text = await zip.file(part)!.async("string");
    expect(text.split(from).length - 1, `${from} in ${part}`).toBe(1);
    zip.file(part, text.replace(from, to));
  }
  return zip.generateAsync({ type: "uint8array" });
};

const chartsOf = async (
  bytes: Uint8Array,
  limits?: Parameters<typeof extractXlsx>[2],
): Promise<{ charts: Chart[]; notices: unknown[] }> => {
  const { doc, notices } = await extractXlsx(bytes, "en", limits);
  const found: Chart[] = [];
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (typeof value === "object" && value !== null) {
      if ((value as { kind?: string }).kind === "chart") found.push(value as Chart);
      Object.values(value).forEach(walk);
    }
  };
  walk(doc.blocks);
  return { charts: found, notices };
};

/** What the chart of the unchanged file gave before the sheet was read: titles and empty cells, a drawing with no number. */
const TODAYS_CHART: Chart = {
  kind: "chart",
  data: { headers: ["", "", ""], rows: [] },
  drawing: { type: "hBar", values: [[], []], categoriesReversed: true },
  title: "Points scored",
};

describe("a workbook chart with no cached values is drawn from its sheet", () => {
  it("🔴 fixture 17: four rows, the sheet's names and numbers, drawn", async () => {
    const { charts } = await chartsOf(new Uint8Array(readFileSync(FIXTURE)));
    expect(charts).toHaveLength(1);
    expect(charts[0]!.data.rows).toEqual([
      ["Team 1", "24", "89"],
      ["Team 2", "36", "24"],
      ["Team 3", "12", "37"],
      ["Team 4", "38", "63"],
    ]);
    expect(charts[0]!.data.headers).toEqual(["", "Period 1", "Period 2"]);
    expect(charts[0]!.drawing).toEqual({
      type: "hBar",
      values: [
        [24, 36, 12, 38],
        [89, 24, 37, 63],
      ],
      categoriesReversed: true,
    });
    expect(charts[0]!.title).toBe("Points scored");
  });

  it("🔴 a hidden row inside the range gives no point, whatever plotVisOnly says", async () => {
    // Row 3 is Team 2. Both values of plotVisOnly are tried: the file's own (1) and the opposite, so the rule is not read from it.
    for (const visOnly of ["1", "0"]) {
      const bytes = await patched(
        [SHEET, '<row r="3">', '<row r="3" hidden="1">'],
        [CHART, '<c:plotVisOnly val="1"', `<c:plotVisOnly val="${visOnly}"`],
      );
      const { charts } = await chartsOf(bytes);
      expect(charts[0]!.data.rows, `plotVisOnly=${visOnly}`).toEqual([
        ["Team 1", "24", "89"],
        ["Team 3", "12", "37"],
        ["Team 4", "38", "63"],
      ]);
      expect(charts[0]!.drawing).toEqual({
        type: "hBar",
        values: [
          [24, 12, 38],
          [89, 37, 63],
        ],
        categoriesReversed: true,
      });
    }
  });

  it("🔴 a hidden column holding the values gives no point for that series", async () => {
    const sheetWithHiddenColumn = '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/>';
    const bytes = await patched([
      SHEET,
      sheetWithHiddenColumn,
      `${sheetWithHiddenColumn}<cols><col min="3" max="3" width="12" hidden="1" customWidth="1"/></cols>`,
    ]);
    const { charts } = await chartsOf(bytes);
    expect(charts[0]!.data.headers).toEqual(["", "Period 1", ""]);
    expect(charts[0]!.data.rows).toEqual([
      ["Team 1", "24", ""],
      ["Team 2", "36", ""],
      ["Team 3", "12", ""],
      ["Team 4", "38", ""],
    ]);
    expect(charts[0]!.drawing).toEqual({
      type: "hBar",
      values: [
        [24, 36, 12, 38],
        [null, null, null, null],
      ],
      categoriesReversed: true,
    });
  });

  it("🔴 a range reaching past the point bound with a filled cell withholds the chart, said by the notice, decided before any cell is read", async () => {
    // Row 20000 holds a value and the range now ends there: its index (19 998) is past MAX_CHART_POINTS.
    const bytes = await patched(
      [SHEET, "</sheetData>", '<row r="20000"><c r="B20000"><v>5</v></c></row></sheetData>'],
      [CHART, "Bar!$B$2:$B$50", "Bar!$B$2:$B$30000"],
    );
    const { charts, notices } = await chartsOf(bytes);
    expect(charts).toEqual([]);
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
  });

  it("🔴 a range reaching past the bound over empty cells is read: nothing is past it", async () => {
    const { charts } = await chartsOf(await patched([CHART, "Bar!$B$2:$B$50", "Bar!$B$2:$B$30000"]));
    expect(charts[0]!.data.rows).toHaveLength(4);
  });

  const abstentions: [string, [string, string, string]][] = [
    ["another sheet", [CHART, "Bar!$B$2:$B$50", "Other!$B$2:$B$50"]],
    ["an external reference", [CHART, "Bar!$B$2:$B$50", "[1]Bar!$B$2:$B$50"]],
    ["a defined name", [CHART, "Bar!$B$2:$B$50", "SalesFigures"]],
    ["a whole column", [CHART, "Bar!$B$2:$B$50", "Bar!$B:$B"]],
    ["a two-dimensional range", [CHART, "Bar!$B$2:$B$50", "Bar!$B$2:$C$50"]],
    ["several areas", [CHART, "Bar!$B$2:$B$50", "(Bar!$B$2:$B$3,Bar!$B$4:$B$5)"]],
  ];
  for (const [title, edit] of abstentions) {
    it(`🔴 ${title}: one reference outside the grammar leaves the whole chart as it was`, async () => {
      const { charts } = await chartsOf(await patched(edit));
      expect(charts).toEqual([TODAYS_CHART]);
    });
  }

  it("🔴 a chart that caches a value keeps its cache, whatever the sheet holds", async () => {
    // One series names a cache with a point (« 7 » at index 0): the sheet is not read for ANY reference of the chart.
    const bytes = await patched([
      CHART,
      "Bar!$B$2:$B$50</c:f><c:numCache/>",
      'Bar!$B$2:$B$50</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>7</c:v></c:pt></c:numCache>',
    ]);
    const { charts } = await chartsOf(bytes);
    expect(charts[0]!.data.rows).toEqual([["", "7", ""]]);
  });
});
