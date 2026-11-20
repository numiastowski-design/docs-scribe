/** The page and the structure control each write the grammar of a chart's address (`extractor/xlsx-chart-reference.ts`,
 *  `pipeline/source-chart-reference.ts`), so that neither can pass the defect of the other. This test is the ONE place that imports both: the
 *  same table of addresses, accepted or refused, with the same range for each accepted one, and the same cells for the same kept cells; then
 *  both real readers on the same workbook bytes, on the cells where they take their data from different sources. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractXlsx } from "../extractor/xlsx";
import { parseChartReference, sheetChartResolver } from "../extractor/xlsx-chart-reference";
import { keptCellsResolver, spanOf } from "./source-chart-reference";
import { readXlsxConservationSource, readXlsxConservationSourceUnder } from "./source-reading-xlsx";

const ADDRESSES = [
  "Bar!$B$1",
  "Bar!$B$2:$B$50",
  "'Base LA Summary Data'!$C$10:$G$10",
  "Bar!B2:B5",
  "Bar!$B2:B$5",
  "'Bar'!A1",
  "'Team''s'!A1:A3",
  "Bar!A1:A1",
  "Bar!XFD1048576",
  "BAR!A1",
  "Résumé!B2:B4",
  "$B$2:$B$50",
  "B2",
  "Other!$B$2:$B$5",
  "[1]Bar!$B$2:$B$5",
  "'[1]Bar'!B2",
  "SalesFigures",
  "(Bar!$B$2:$B$3,Bar!$B$5:$B$6)",
  "Bar!B2:B3,Bar!B5:B6",
  "Bar!$B:$B",
  "Bar!$2:$2",
  "Bar:Other!B2",
  "'Bar:Other'!B2",
  "Table1[Sales]",
  "Bar!A1:C5",
  "Bar!B5:B2",
  "Bar!G10:C10",
  "=Bar!B2",
  "SUM(Bar!B2:B5)",
  " Bar!B2",
  "Bar!B2 :B5",
  "Bar!b2",
  "Bar!A1048577",
  "Bar!XFE1",
  "Bar!A0",
  "Bar!A01",
  "Bar!A1:A2:A3",
  "Bar!",
  "",
  "!B2",
  "'Bar!B2",
];

describe("the page's grammar and the control's grammar agree", () => {
  for (const holder of ["Bar", "Base LA Summary Data", "Team's", "Résumé"]) {
    it(`🔴 every address of the table, holding sheet « ${holder} »`, () => {
      for (const address of ADDRESSES) {
        const page = parseChartReference(address, holder);
        const control = spanOf(address, holder);
        expect(
          control === undefined
            ? undefined
            : {
                firstRow: control.firstRow - 1,
                lastRow: control.lastRow - 1,
                firstColumn: control.firstCol - 1,
                lastColumn: control.lastCol - 1,
              },
          address,
        ).toEqual(page);
      }
    });
  }

  it("🔴 the table exercises both answers", () => {
    const answers = ADDRESSES.map((address) => parseChartReference(address, "Bar") !== undefined);
    expect(answers.filter(Boolean).length).toBeGreaterThan(5);
    expect(answers.filter((accepted) => !accepted).length).toBeGreaterThan(20);
  });

  it("🔴 the same kept cells give the same points, at the same indices", () => {
    // Zero-based for the page (row, column), one-based for the control; the same sheet: B1 "Period 1", A2 "Team 1", B2 24, B4 12; row 3 empty.
    const texts = new Map([
      ["0,1", "Period 1"],
      ["1,0", "Team 1"],
      ["1,1", "24"],
      ["3,1", "12"],
    ]);
    const numbers = new Map([
      ["1,1", 24],
      ["3,1", 12],
    ]);
    const page = sheetChartResolver("Bar", {
      filledCells: [
        [0, 1],
        [1, 0],
        [1, 1],
        [3, 1],
      ],
      textAt: (row, column) => texts.get(`${row},${column}`) ?? "",
      numberAt: (row, column) => {
        const value = numbers.get(`${row},${column}`);
        return value === undefined ? undefined : { value, formatCode: undefined };
      },
    });
    const control = keptCellsResolver("Bar", [
      { rowNum: 4, colNum: 2, text: "12", numeric: true },
      { rowNum: 1, colNum: 2, text: "Period 1", numeric: false },
      { rowNum: 2, colNum: 1, text: "Team 1", numeric: false },
      { rowNum: 2, colNum: 2, text: "24", numeric: true },
    ]);
    for (const address of ["Bar!$B$2:$B$50", "Bar!$A$2:$A$4", "Bar!$A$2:$C$2", "Bar!$B$1"]) {
      const fromPage = [...page(address)!].map((point) => [point.index, point.text, point.stored !== undefined]);
      const fromControl = [...control(address)!].map((cell) => [cell.index, cell.text, cell.numeric]);
      expect(fromControl, address).toEqual(fromPage);
      expect(fromPage.length, `${address} reads something`).toBeGreaterThan(0);
    }
  });
});

/** Fixture 17 (Google's export, `Bar!A1:C5`, a chart over `Bar!$A$2:$A$50` / `$B$2:$B$50` / `$C$2:$C$50` with no cache), changed in memory on
 *  the cells where the page and the control take their data from different sources: the page from the formatted grid text, the control from
 *  the raw cell and its own « hidden by format » rule. Both real readers read the SAME bytes. */
const FIXTURE = join(
  dirname(fileURLToPath(import.meta.url)),
  "../../fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx",
);
const SHEET = "xl/worksheets/sheet1.xml";
const STYLES = "xl/styles.xml";
const ROW_7 = '<row r="7"><c r="A7" s="5"/><c r="B7" s="5"/><c r="C7" s="5"/>';

/** The fixture with style 6 carrying `formatCode` (custom id 164) and row 7 starting with `cells`. */
const withRow7 = async (cells: string, formatCode = "General"): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
  const edit = async (part: string, from: string, to: string): Promise<void> => {
    const text = await zip.file(part)!.async("string");
    expect(text.split(from).length - 1, `${from} in ${part}`).toBe(1);
    zip.file(part, text.replace(from, to));
  };
  await edit(SHEET, ROW_7, `<row r="7">${cells}`);
  await edit(
    STYLES,
    '<fonts count="2">',
    `<numFmts count="1"><numFmt numFmtId="164" formatCode="${formatCode}"/></numFmts><fonts count="2">`,
  );
  await edit(STYLES, '<cellXfs count="6">', '<cellXfs count="7">');
  await edit(
    STYLES,
    "</cellXfs>",
    '<xf borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0" applyNumberFormat="1"/></cellXfs>',
  );
  return zip.generateAsync({ type: "uint8array" });
};

const patchedSheet = async (...edits: [from: string, to: string][]): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
  for (const [from, to] of edits) {
    const text = await zip.file(SHEET)!.async("string");
    expect(text.split(from).length - 1, `${from} in ${SHEET}`).toBe(1);
    zip.file(SHEET, text.replace(from, to));
  }
  return zip.generateAsync({ type: "uint8array" });
};

/** The rows and cells the control says the chart adds (header row included). */
const controlShape = async (bytes: Uint8Array): Promise<{ rows: number; cells: number }> => {
  const source = await readXlsxConservationSource(bytes);
  if ("readFailed" in source) throw new Error("the control could not read the workbook");
  const { added } = source.structure!.units[0]!;
  return { rows: added.rows, cells: added.cells };
};

/** The chart block the page publishes. */
const pageChart = async (bytes: Uint8Array): Promise<Extract<IRBlock, { kind: "chart" }>> => {
  const charts: Extract<IRBlock, { kind: "chart" }>[] = [];
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (typeof value === "object" && value !== null) {
      if ((value as { kind?: string }).kind === "chart") charts.push(value as Extract<IRBlock, { kind: "chart" }>);
      Object.values(value).forEach(walk);
    }
  };
  walk((await extractXlsx(bytes, "en")).doc.blocks);
  expect(charts).toHaveLength(1);
  return charts[0]!;
};

describe("the page and the control agree on which cells of a cache-less chart are points", () => {
  vi.setConfig({ testTimeout: 60_000 });
  // Each case: the workbook, then the page's rows (header excluded): the fixture gives four.
  const cases: [string, () => Promise<Uint8Array>, string[][]][] = [
    [
      "an empty-section format on zero (`0;-0;;@`), its category empty: the row shows nothing",
      () => withRow7('<c r="B7" s="6"><v>0</v></c>', "0;-0;;@"),
      [],
    ],
    [
      "a colour tag before an empty zero section (`[Blue]0;[Red]-0;;@`), its category empty: the row shows nothing",
      () => withRow7('<c r="B7" s="6"><v>0</v></c>', "[Blue]0;[Red]-0;;@"),
      [],
    ],
    [
      "whitespace-only text as a category, a number beside it",
      () => withRow7('<c r="A7" t="inlineStr"><is><t xml:space="preserve">   </t></is></c><c r="B7"><v>5</v></c>'),
      [["", "5", ""]],
    ],
    [
      "`;;;` on a text category: no point, no text",
      () => withRow7('<c r="A7" s="6" t="inlineStr"><is><t>Secret</t></is></c><c r="B7"><v>5</v></c>', ";;;"),
      [["", "5", ""]],
    ],
    [
      "`;;;` on a number beside a category",
      () => withRow7('<c r="A7" t="inlineStr"><is><t>Team 5</t></is></c><c r="B7" s="6"><v>5</v></c>', ";;;"),
      [["Team 5", "", ""]],
    ],
    ["a hidden row", () => patchedSheet(['<row r="3">', '<row r="3" hidden="1">']), []],
    [
      "a hidden column",
      () =>
        patchedSheet([
          '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/>',
          '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/><cols><col min="3" max="3" width="12" hidden="1" customWidth="1"/></cols>',
        ]),
      [],
    ],
    [
      "a filtered row",
      () =>
        patchedSheet(
          ['<row r="4">', '<row r="4" hidden="1">'],
          ["</sheetData>", '</sheetData><autoFilter ref="A1:C5"/>'],
        ),
      [],
    ],
  ];
  const FIXTURE_ROWS = [
    ["Team 1", "24", "89"],
    ["Team 2", "36", "24"],
    ["Team 3", "12", "37"],
    ["Team 4", "38", "63"],
  ];
  for (const [title, build, extra] of cases) {
    it(`🔴 ${title}`, async () => {
      const bytes = await build();
      const chart = await pageChart(bytes);
      const pageShape = {
        rows: chart.data.rows.length + 1,
        cells: (chart.data.rows.length + 1) * chart.data.headers!.length,
      };
      expect(await controlShape(bytes)).toEqual(pageShape);
      if (extra.length > 0) expect(chart.data.rows).toEqual([...FIXTURE_ROWS, ...extra]);
      expect(JSON.stringify(chart)).not.toContain("Secret");
    });
  }
});

describe("the page and the control decide alike whether a chart's references fit the cell bound", () => {
  vi.setConfig({ testTimeout: 60_000 });
  const bytes = new Uint8Array(readFileSync(FIXTURE));
  /** The chart blocks the page publishes under `limits`, and its notices. */
  const chartsUnder = async (limits: {
    chartCells: number;
  }): Promise<{ charts: Extract<IRBlock, { kind: "chart" }>[]; notices: unknown[] }> => {
    const charts: Extract<IRBlock, { kind: "chart" }>[] = [];
    const walk = (value: unknown): void => {
      if (Array.isArray(value)) value.forEach(walk);
      else if (typeof value === "object" && value !== null) {
        if ((value as { kind?: string }).kind === "chart") charts.push(value as Extract<IRBlock, { kind: "chart" }>);
        Object.values(value).forEach(walk);
      }
    };
    const { doc, notices } = await extractXlsx(bytes, "en", limits);
    walk(doc.blocks);
    return { charts, notices };
  };

  it("🔴 fixture 17: each range is clipped to its last filled cell, the chart's references can yield 14 cells", async () => {
    // Three ranges of 49 rows over four filled ones (A2:A5, B2:B5, C2:C5) and two name cells; the categories `Bar!$A$2:$A$50` are named by
    // both series and read once, so counted once: 2 + 4 + 4 + 4 = 14. Unclipped, 2 + 49 x 3 = 149.
    const page = sheetChartResolver("Bar", {
      filledCells: [
        [0, 1],
        [0, 2],
        ...[1, 2, 3, 4].flatMap(
          (row) =>
            [
              [row, 0],
              [row, 1],
              [row, 2],
            ] as const,
        ),
      ],
      textAt: () => "x",
      numberAt: () => undefined,
    });
    const control = keptCellsResolver(
      "Bar",
      [
        [1, 2],
        [1, 3],
        ...[2, 3, 4, 5].flatMap((row) => [
          [row, 1],
          [row, 2],
          [row, 3],
        ]),
      ].map(([rowNum, colNum]) => ({ rowNum: rowNum!, colNum: colNum!, text: "x", numeric: false })),
    );
    for (const address of ["Bar!$A$2:$A$50", "Bar!$B$2:$B$50", "Bar!$C$2:$C$50"]) {
      expect(page(address)?.extent, address).toBe(4);
      expect(control(address)?.extent, address).toBe(4);
    }
    // The chart is drawn at a cell bound of 14, by both readings (its table holds 3 x 4 = 12 cells: the references' count decides).
    const at = { chartCells: 14 };
    const { charts } = await chartsUnder(at);
    expect(charts).toHaveLength(1);
    expect(charts[0]!.data.rows).toHaveLength(4);
    const source = await readXlsxConservationSourceUnder(bytes, at);
    if ("readFailed" in source) throw new Error("the control could not read the workbook");
    expect(source.structure!.units[0]!.added).toMatchObject({ rows: 5, cells: 15 });
  });

  it("🔴 one cell under it, both readings withhold the chart, the page with its notice", async () => {
    const under = { chartCells: 13 };
    const { charts, notices } = await chartsUnder(under);
    expect(charts).toEqual([]);
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
    const source = await readXlsxConservationSourceUnder(bytes, under);
    if ("readFailed" in source) throw new Error("the control could not read the workbook");
    // No chart table expected: the sheet's own table alone.
    expect(source.structure!.units[0]!.added).toMatchObject({ tables: 1, rows: 0, cells: 0 });
  });
});
