/** The structure control counts the table of a workbook chart whose file cached no value from the SAME cells the page reads (docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md),
 *  through its own reading (`source-chart-reference.ts`): fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx, then that file with one part changed in memory.
 *  Each case also compares the control's count with the table the PAGE publishes (`extractXlsx`), so the two readings are held to each other. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { extractXlsx } from "../extractor/xlsx";
import type { IRBlock } from "doc-ir";
import { readXlsxConservationSource } from "./source-reading-xlsx";

vi.setConfig({ testTimeout: 60_000 });

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "../../fixtures/xlsx-reel/17-GSheets-Essai13-Graphique.xlsx");
const SHEET = "xl/worksheets/sheet1.xml";
const CHART = "xl/charts/chart1.xml";

const patched = async (...edits: [part: string, from: string, to: string][]): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
  for (const [part, from, to] of edits) {
    const text = await zip.file(part)!.async("string");
    expect(text.split(from).length - 1, `${from} in ${part}`).toBe(1);
    zip.file(part, text.replace(from, to));
  }
  return zip.generateAsync({ type: "uint8array" });
};

/** What the control declares the chart adds: its rows and cells (the sheet's own table adds one table and no row). */
const controlCount = async (bytes: Uint8Array): Promise<{ tables: number; rows: number; cells: number }> => {
  const source = await readXlsxConservationSource(bytes);
  if ("readFailed" in source) throw new Error("the control could not read the workbook");
  const { added } = source.structure!.units[0]!;
  return { tables: added.tables, rows: added.rows, cells: added.cells };
};

/** The same, counted on the chart block the page publishes: header row included. */
const pageCount = async (bytes: Uint8Array): Promise<{ tables: number; rows: number; cells: number }> => {
  const charts: Extract<IRBlock, { kind: "chart" }>[] = [];
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (typeof value === "object" && value !== null) {
      if ((value as { kind?: string }).kind === "chart") charts.push(value as Extract<IRBlock, { kind: "chart" }>);
      Object.values(value).forEach(walk);
    }
  };
  walk((await extractXlsx(bytes, "en")).doc.blocks);
  return {
    tables: 1 + charts.length,
    rows: charts.reduce((sum, chart) => sum + chart.data.rows.length + 1, 0),
    cells: charts.reduce((sum, chart) => sum + (chart.data.rows.length + 1) * chart.data.headers!.length, 0),
  };
};

describe("the control reads a chart without cache from the same cells as the page", () => {
  const cases: [string, () => Promise<Uint8Array>, { tables: number; rows: number; cells: number }][] = [
    [
      "fixture 17 as it is: four rows, three columns, header included",
      async () => new Uint8Array(readFileSync(FIXTURE)),
      { tables: 2, rows: 5, cells: 15 },
    ],
    [
      "a hidden row inside the range: three rows",
      () => patched([SHEET, '<row r="3">', '<row r="3" hidden="1">']),
      { tables: 2, rows: 4, cells: 12 },
    ],
    [
      "a hidden column holding one series: its column stays, empty",
      () =>
        patched([
          SHEET,
          '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/>',
          '<sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/><cols><col min="3" max="3" width="12" hidden="1" customWidth="1"/></cols>',
        ]),
      { tables: 2, rows: 5, cells: 15 },
    ],
    [
      "another sheet named by a reference: the chart stays as the file wrote it",
      () => patched([CHART, "Bar!$B$2:$B$50", "Other!$B$2:$B$50"]),
      { tables: 2, rows: 1, cells: 3 },
    ],
    [
      "a range past the point bound with a filled cell: the chart is withheld",
      () =>
        patched(
          [SHEET, "</sheetData>", '<row r="20000"><c r="B20000"><v>5</v></c></row></sheetData>'],
          [CHART, "Bar!$B$2:$B$50", "Bar!$B$2:$B$30000"],
        ),
      { tables: 1, rows: 0, cells: 0 },
    ],
  ];
  for (const [title, build, expected] of cases) {
    it(`🔴 ${title}`, async () => {
      const bytes = await build();
      expect(await controlCount(bytes)).toEqual(expected);
      expect(await pageCount(bytes)).toEqual(expected);
    });
  }
});

describe("one chart part anchored on two sheets", () => {
  /** Fixture 17 with a second sheet « Other », a copy of « Bar » whose drawing is the same part: the chart names `Bar!…`, so it reads its cells
   *  on « Bar » and keeps its (empty) cache on « Other ». */
  const twoSheets = async (): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
    const edit = async (part: string, from: string, to: string): Promise<void> => {
      const text = await zip.file(part)!.async("string");
      expect(text.split(from).length - 1, `${from} in ${part}`).toBe(1);
      zip.file(part, text.replace(from, to));
    };
    zip.file("xl/worksheets/sheet2.xml", await zip.file(SHEET)!.async("string"));
    zip.file(
      "xl/worksheets/_rels/sheet2.xml.rels",
      await zip.file("xl/worksheets/_rels/sheet1.xml.rels")!.async("string"),
    );
    await edit(
      "xl/workbook.xml",
      '<sheet state="visible" name="Bar" sheetId="1" r:id="rId5"/>',
      '<sheet state="visible" name="Bar" sheetId="1" r:id="rId5"/><sheet state="visible" name="Other" sheetId="2" r:id="rId6"/>',
    );
    await edit(
      "xl/_rels/workbook.xml.rels",
      "</Relationships>",
      '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
    );
    await edit(
      "[Content_Types].xml",
      "</Types>",
      '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet2.xml"/></Types>',
    );
    return zip.generateAsync({ type: "uint8array" });
  };

  it("🔴 each sheet gets its own chart table: from the cells on « Bar », from the cache on « Other »", async () => {
    const bytes = await twoSheets();
    const source = await readXlsxConservationSource(bytes);
    if ("readFailed" in source) throw new Error("the control could not read the workbook");
    const added = source.structure!.units.map((unit) => ({
      tables: unit.added.tables,
      rows: unit.added.rows,
      cells: unit.added.cells,
    }));
    expect(added).toEqual([
      { tables: 2, rows: 5, cells: 15 },
      { tables: 2, rows: 1, cells: 3 },
    ]);
    // The page, sheet by sheet: the chart in document order, « Bar » first.
    const charts: Extract<IRBlock, { kind: "chart" }>[] = [];
    const walk = (value: unknown): void => {
      if (Array.isArray(value)) value.forEach(walk);
      else if (typeof value === "object" && value !== null) {
        if ((value as { kind?: string }).kind === "chart") charts.push(value as Extract<IRBlock, { kind: "chart" }>);
        Object.values(value).forEach(walk);
      }
    };
    walk((await extractXlsx(bytes, "en")).doc.blocks);
    expect(charts.map((chart) => chart.data.rows.length + 1)).toEqual([5, 1]);
  });
});
