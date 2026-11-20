/**
 * Lot tableurs step 11 (witnesses written BEFORE the code, TDD pair of
 * 28/09/2026): `extractGoogleSheet` delivers the SAME shape as the `.xlsx`
 * reader — ONE `tabs` block, SHARED caps (`lib/reader-limits.ts`) counted
 * after the retreats — plus the two remaining Sheets specifics: the header
 * declared by frozen rows, and the OBJECT tab published as a single-line
 * tab. The chart line used to be « Chart "T" — source: 'Feuille'!A1:B10. »
 * (GridRange → A1 notation); since A3 of
 * docs/CONCEPTION-CARTE-ET-REMARQUES.md (29/09/2026) it is the
 * chart's TITLE ALONE, no "Chart" word, no source range.
 *
 * The real sheet these rules were first held on (a raw API answer of 298 rows × 6 columns, over the `MAX_SHEET_ROWS`
 * cap, refused whole) is not a document this package carries: its case is held by the publisher's application.
 */
import { describe, expect, it } from "vitest";
import type { IRBlock } from "doc-ir";
import { MAX_NOTICE_SHEET_NAME_CHARS, MAX_NOTICE_SHEET_NAMES } from "doc-ir/internal";
import { MAX_WORKBOOK_TABS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractGoogleSheet, type GoogleSheetMeta } from "./google-sheet";

const tabsOf = (doc: { blocks: IRBlock[] }): Extract<IRBlock, { kind: "tabs" }> | undefined =>
  doc.blocks[0]?.kind === "tabs" ? doc.blocks[0] : undefined;

/** A minimal GRID sheet: title, index, values injected by the caller. */
const grid = (
  title: string,
  index: number,
  over: Partial<GoogleSheetMeta["sheets"][number]["properties"]> = {},
): GoogleSheetMeta["sheets"][number] => ({
  properties: { title, index, sheetType: "GRID", ...over },
});

describe("extractGoogleSheet — one tabs block, the xlsx shape and caps (étape 11)", () => {
  it("🔴 a small grid publishes its table; frozenRowCount=1 marks the first row as header", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [grid("Ventes", 0, { gridProperties: { frozenRowCount: 1 } })],
    };
    const { doc } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "Ventes",
          [
            ["Mois", "CA"],
            ["Janvier", "18 400 €"],
          ],
        ],
      ]),
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.kind).toBe("table");
    expect(table.rows[0]!.cells.every((c) => c.header === true)).toBe(true);
    expect(table.rows[1]!.cells.every((c) => c.header === undefined)).toBe(true);
  });

  it("🔴 a frozen HEADER row whose only text sits in its first cell KEEPS its cells — never collapsed to one colspan: the file DECLARED it the header, its empty cells align with the data columns (revue 27/10/2026)", async () => {
    const meta: GoogleSheetMeta = { sheets: [grid("Stock", 0, { gridProperties: { frozenRowCount: 1 } })] };
    const { doc } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "Stock",
          [
            ["Inventory", "", ""],
            ["Nails", "12", "A3"],
            ["Screws", "40", "B1"],
          ],
        ],
      ]),
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells).toHaveLength(3);
    expect(table.rows[0]!.cells.every((c) => c.header === true)).toBe(true);
    expect(table.rows[0]!.cells[0]!.colspan).toBeUndefined();
  });

  it("🔴 no freeze, no Excel-style declaration: no header row", async () => {
    const { doc } = await extractGoogleSheet(
      { sheets: [grid("Brut", 0)] },
      new Map([["Brut", [["a"], ["b"]]]]),
      PAGE_SIZED_LIMITS,
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells.every((c) => c.header === undefined)).toBe(true);
  });

  it("🔴 an embedded chart becomes ONE line under the table: its title alone (A3, 29/09/2026 — no 'Chart' word, no source range)", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [
        {
          properties: { sheetId: 7, title: "Feuille", index: 0, sheetType: "GRID" },
          charts: [{ spec: { title: "T" } }],
        },
      ],
    };
    const { doc } = await extractGoogleSheet(meta, new Map([["Feuille", [["x", "y"]]]]), PAGE_SIZED_LIMITS);
    const blocks = tabsOf(doc)!.tabs[0]!.blocks;
    expect(blocks[0]!.kind).toBe("table");
    expect(blocks[1]).toEqual({ kind: "paragraph", inlines: [{ text: "T" }] });
  });

  it("🔴 a chart with no title publishes no line at all (reco of §1.3.3: 'sans titre, rien')", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [{ properties: { sheetId: 7, title: "Feuille", index: 0, sheetType: "GRID" }, charts: [{ spec: {} }] }],
    };
    const { doc } = await extractGoogleSheet(meta, new Map([["Feuille", [["x", "y"]]]]), PAGE_SIZED_LIMITS);
    expect(tabsOf(doc)!.tabs[0]!.blocks).toHaveLength(1); // the table alone
  });

  it("🔴 an OBJECT tab publishes as a tab carrying its chart line ALONE — no grid was ever fetched", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [
        grid("Données", 0, { sheetId: 1 }),
        {
          properties: { sheetId: 2, title: "Graphe seul", index: 1, sheetType: "OBJECT" },
          charts: [{ spec: { title: "Marge" } }],
        },
      ],
    };
    const { doc } = await extractGoogleSheet(meta, new Map([["Données", [["a"]]]]), PAGE_SIZED_LIMITS); // no values for the OBJECT tab, by design
    const tabs = tabsOf(doc)!.tabs;
    expect(tabs.map((t) => t.title)).toEqual(["Données", "Graphe seul"]);
    expect(tabs[1]!.blocks).toEqual([{ kind: "paragraph", inlines: [{ text: "Marge" }] }]);
  });

  it("an OBJECT tab whose charts have no title still holds one empty paragraph, as pptx pads an empty tab (review 30/09/2026, finding 8)", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [{ properties: { sheetId: 2, title: "Graphe", index: 0, sheetType: "OBJECT" }, charts: [{ spec: {} }] }],
    };
    const { doc } = await extractGoogleSheet(meta, new Map(), PAGE_SIZED_LIMITS);
    expect(tabsOf(doc)!.tabs[0]!.blocks).toEqual([{ kind: "paragraph", inlines: [] }]);
  });

  it("🔴 the visible sheet past the tab cap carries NO TAB AT ALL, the dashboard names it — the xlsx cap, shared (A7; cap 20 since 27/10/2026, was 6)", async () => {
    const count = MAX_WORKBOOK_TABS + 1;
    const meta: GoogleSheetMeta = { sheets: Array.from({ length: count }, (_, i) => grid(`F${i + 1}`, i)) };
    const { doc, notices } = await extractGoogleSheet(
      meta,
      new Map(Array.from({ length: count }, (_, i) => [`F${i + 1}`, [["x"]]] as [string, string[][]])),
      PAGE_SIZED_LIMITS,
    );
    const tabs = tabsOf(doc)!;
    expect(tabs.tabs).toHaveLength(MAX_WORKBOOK_TABS);
    expect(doc.blocks).toHaveLength(1); // no more paragraph under the tabs block
    expect(notices).toEqual([{ kind: "workbook-cap", sheets: [`F${count}`], more: 0 }]);
  });

  it("🔴 workbook-cap and sheet-withheld names are BOUNDED like xlsx's (item E): 5 names cut to 64 characters + `more`", async () => {
    const long = (i: number) => `${String(i).padStart(2, "0")}${"n".repeat(98)}`;
    // MAX_WORKBOOK_TABS + 30 sheets: 30 past the cap, so 5 named + 25 in `more`, whatever the cap's value.
    const meta: GoogleSheetMeta = {
      sheets: Array.from({ length: MAX_WORKBOOK_TABS + 30 }, (_, i) => grid(long(i + 1), i)),
    };
    const { notices } = await extractGoogleSheet(
      meta,
      new Map(meta.sheets.map((s) => [s.properties.title, [["x"]]] as [string, string[][]])),
      PAGE_SIZED_LIMITS,
    );
    const cap = notices.find((n) => n.kind === "workbook-cap");
    if (cap?.kind !== "workbook-cap") throw new Error("expected a workbook-cap notice");
    expect(cap.sheets).toHaveLength(MAX_NOTICE_SHEET_NAMES);
    expect(cap.more).toBe(30 - MAX_NOTICE_SHEET_NAMES);
    for (const name of cap.sheets) expect(name.length).toBeLessThanOrEqual(MAX_NOTICE_SHEET_NAME_CHARS + 1);
  });

  it("🔴 a withheld sheet's name is cut to 64 characters in its notice (item E)", async () => {
    const name = "w".repeat(100);
    // A second sheet that fits keeps the workbook publishable (a workbook with no such sheet is refused).
    const meta: GoogleSheetMeta = { sheets: [grid(name, 0), grid("ok", 1)] };
    const rows = Array.from({ length: 2000 }, () => ["x"]);
    const { notices } = await extractGoogleSheet(
      meta,
      new Map([
        [name, rows],
        ["ok", [["x"]]],
      ]),
      PAGE_SIZED_LIMITS,
    );
    const w = notices.find((n) => n.kind === "sheet-withheld");
    if (w?.kind !== "sheet-withheld") throw new Error("expected a sheet-withheld notice");
    expect(w.sheet.length).toBeLessThanOrEqual(MAX_NOTICE_SHEET_NAME_CHARS + 1);
  });
});

/** Same trimming as the `.xlsx` (edges and inside) and same bounding of merges, both before the caps. */
describe("extractGoogleSheet — empty lines out and merges bounded BEFORE the caps (revue 29/09, constat 5)", () => {
  it("🔴 data in A–C and M: the nine empty columns between drop, the sheet publishes 4 columns", async () => {
    const row = (a: string, m: string): string[] => [a, "b", "c", "", "", "", "", "", "", "", "", "", m];
    const { doc } = await extractGoogleSheet(
      { sheets: [grid("Large", 0)] },
      new Map([["Large", [row("a1", "m1"), row("a2", "m2")]]]),
      PAGE_SIZED_LIMITS,
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.kind).toBe("table");
    expect(table.rows.map((r) => r.cells.length)).toEqual([4, 4]);
  });

  it("🔴 a merge declared down to row 1000 never widens the grid: it spans the kept rows only", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [
        {
          ...grid("Fusion", 0),
          merges: [{ startRowIndex: 0, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 1 }],
        },
      ],
    };
    const { doc } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "Fusion",
          [
            ["titre", "x"],
            ["", "y"],
            ["", "z"],
          ],
        ],
      ]),
      PAGE_SIZED_LIMITS,
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.kind).toBe("table");
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0]!.cells[0]!.rowspan).toBe(3);
  });

  it("🔴 an interior empty row drops, and a row whose only text is in a hidden column drops too", async () => {
    const meta: GoogleSheetMeta = { sheets: [{ ...grid("T", 0), columnMetadata: [{}, { hiddenByUser: true }] }] };
    const { doc } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "T",
          [
            ["a", "x"],
            ["", ""],
            ["", "secret"],
            ["d", "y"],
          ],
        ],
      ]),
      PAGE_SIZED_LIMITS,
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows.map((r) => r.cells.map((c) => c.blocks.length))).toEqual([[1], [1]]);
  });
});

// The "chart source line never guesses, never names a hidden tab" witnesses
// (revue HIGH du 29/09, constat 8) tested the GridRange → A1 resolution:
// gone since A3 of docs/CONCEPTION-CARTE-ET-REMARQUES.md (29/09/2026), the
// chart line is the title alone, so there is no source, no sheetId, no
// hidden-tab guessing left to test here.

describe("extractGoogleSheet — a merge starting at A1 (start indices omitted by proto3) (29/09)", () => {
  it("🔴 the merge still spans its cells instead of vanishing", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [{ ...grid("Fusion", 0), merges: [{ endRowIndex: 1, endColumnIndex: 2 }] }],
    };
    const { doc } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "Fusion",
          [
            ["titre", ""],
            ["a", "b"],
          ],
        ],
      ]),
      PAGE_SIZED_LIMITS,
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells[0]!.colspan).toBe(2);
  });
});

// Audit of 07/10/2026: ONE header rule for every spreadsheet (28/09): the LAST frozen row is the header, as the
// .xlsx reader does — with a title in row 1 and the header in row 2, both frozen, the header is row 2, never the title.
describe("extractGoogleSheet — the header is the LAST frozen row, as in .xlsx", () => {
  it("🔴 two frozen rows (a title, then the header): only the second is the header", async () => {
    const { doc } = await extractGoogleSheet(
      { sheets: [grid("Ventes", 0, { gridProperties: { frozenRowCount: 2 } })] },
      new Map([["Ventes", [["Ventes 2026"], ["Mois", "CA"], ["Janvier", "18 400 €"]]]]),
    );
    const table = tabsOf(doc)!.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells.every((c) => c.header === undefined)).toBe(true);
    expect(table.rows[1]!.cells.every((c) => c.header === true)).toBe(true);
  });
});
