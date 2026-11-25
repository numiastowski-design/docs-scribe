import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import type { IRBlock } from "doc-ir";
import { formatCellReference } from "./xlsx-cells";
import { MAX_NOTICE_SHEET_NAME_CHARS, MAX_NOTICE_SHEET_NAMES } from "doc-ir/internal";
import { MAX_WORKBOOK_TABS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { cellText } from "./span-single-first";
import { extractXlsx } from "./xlsx";
import { anchoredPictureWorkbook } from "./xlsx-anchored-picture.test.support";
import { readDeclaredHeaderRows, readWorkbookSheets } from "./xlsx-parts";
import { formatNumberCell } from "./xlsx-number-format";

// A pass-through spy: which cells the extractor bothers to format.
vi.mock("./xlsx-number-format", async (importOriginal) => {
  const original = await importOriginal<typeof import("./xlsx-number-format")>();
  return { ...original, formatNumberCell: vi.fn(original.formatNumberCell) };
});

// Real fixture from the golden corpus (MIKSUDOC « Spécification fonctionnelle » §8 item 4: "un classeur
// multi-onglets ... dont un onglet volumineux"). Structural expectations below
// were cross-checked against the raw sheetN.xml (dimension ref, mergeCell,
// sheet state="hidden" in workbook.xml) before writing this test — never
// against cell prose, so no client content needed to travel through this
// session.
const here = dirname(fileURLToPath(import.meta.url));
const fixture = new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx/04-Effectifs-Atria.xlsx")));

// ---------------------------------------------------------------------------
// A minimal, hand-built .xlsx for the tests below that need one — SheetJS's
// own writer, retired with it 29/09/2026 (`xlsx-cells.parity.test.ts` before
// it, `xlsx-cells.parity.test.ts#fabricated` is the same idea at the cell
// layer). Only what these tests exercise: inline-string cells (no
// sharedStrings indirection needed), hidden rows/columns, several sheets, and
// — for two tests below — the exact raw `<sheetData>`/`<dimension>` a
// generate-then-patch approach used to reach (constats 3 and 10): authored
// directly here, no patching, nothing to prove the patch "took".
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const escapeXml = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

interface SyntheticSheet {
  name: string;
  /** row -> col -> text; `undefined`/empty cells are simply not written. */
  rows?: (string | undefined)[][];
  hiddenRows?: number[]; // zero-based
  hiddenCols?: number[]; // zero-based
  /** Literal `<sheetData>` content, overriding `rows` when given. */
  rawSheetDataXml?: string;
  /** Literal `dimension@ref`, else computed from `rows`. */
  dimension?: string;
  /** Merged ranges (e.g. "A2:D2"), written as `<mergeCells>` after `<sheetData>`. */
  merges?: string[];
}

const sheetXmlOf = (spec: SyntheticSheet): string => {
  const rows = spec.rows ?? [];
  const rowCount = rows.length;
  const colCount = Math.max(0, ...rows.map((r) => r.length));
  const dimension =
    spec.dimension ??
    (rowCount === 0 || colCount === 0 ? "A1" : `A1:${formatCellReference(rowCount - 1, colCount - 1)}`);
  const colsXml =
    spec.hiddenCols && spec.hiddenCols.length > 0
      ? `<cols>${spec.hiddenCols.map((c) => `<col min="${c + 1}" max="${c + 1}" width="9" hidden="1"/>`).join("")}</cols>`
      : "";
  const sheetData =
    spec.rawSheetDataXml ??
    rows
      .map((row, r) => {
        const hidden = spec.hiddenRows?.includes(r) ? ' hidden="1"' : "";
        const cells = row
          .map((text, c) =>
            text === undefined || text.length === 0
              ? ""
              : `<c r="${formatCellReference(r, c)}" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c>`,
          )
          .join("");
        return `<row r="${r + 1}"${hidden}>${cells}</row>`;
      })
      .join("");
  const mergesXml =
    spec.merges && spec.merges.length > 0
      ? `<mergeCells count="${spec.merges.length}">${spec.merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>`
      : "";
  return `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="${dimension}"/>${colsXml}<sheetData>${sheetData}</sheetData>${mergesXml}</worksheet>`;
};

/** A whole, minimal, valid .xlsx from its sheets — built directly with
 *  JSZip, no other library. */
const workbookOf = async (sheets: SyntheticSheet[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets
        .map(
          (_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join("") +
      "</Types>",
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets>` +
      sheets.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
      "</sheets></workbook>",
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheets
        .map(
          (_, i) => `<Relationship Id="rId${i + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join("") +
      "</Relationships>",
  );
  sheets.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXmlOf(s)));
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
};

describe("extractXlsx — hidden cells are never formatted (review finding 12)", () => {
  it("a number in a hidden row or a hidden column costs no formatting", async () => {
    const number = (r: number, c: number, v: number): string => `<c r="${formatCellReference(r, c)}"><v>${v}</v></c>`;
    const bytes = await workbookOf([
      {
        name: "S",
        hiddenRows: [1],
        hiddenCols: [1],
        rawSheetDataXml: `<row r="1">${number(0, 0, 11)}${number(0, 1, 22)}</row><row r="2" hidden="1">${number(1, 0, 33)}</row>`,
      },
    ]);
    vi.mocked(formatNumberCell).mockClear();
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    expect(vi.mocked(formatNumberCell).mock.calls.map((call) => call[0])).toContain(11);
    expect(vi.mocked(formatNumberCell).mock.calls.map((call) => call[0])).not.toContain(22);
    expect(vi.mocked(formatNumberCell).mock.calls.map((call) => call[0])).not.toContain(33);
    expect(doc.blocks).toHaveLength(1);
  });
});

describe("extractXlsx (real corpus fixture)", () => {
  // Since étape 3 of the lot tableurs (28/09/2026) the shape is ONE tabs
  // block — a sheet is a native tab (decisions (e)+(n)); these tests read it.
  const tabsOf = (doc: { blocks: IRBlock[] }): Extract<IRBlock, { kind: "tabs" }> =>
    doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;

  it("produces one TAB per visible, non-withheld sheet, in workbook order, dropping the hidden one", async () => {
    const { doc, notices } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    // workbook.xml lists 4 sheets: 2025, 2026, Grand-Onglet (visible, but
    // WITHHELD — see the row-counts test below), Archive-Interne
    // (state="hidden"). A7 of docs/CONCEPTION-CARTE-ET-REMARQUES.md
    // (29/09/2026): a withheld sheet carries no tab at all any more.
    expect(tabsOf(doc).tabs.map((t) => t.title)).toEqual(["2025", "2026"]);
    expect(notices).toContainEqual({ kind: "sheet-withheld", sheet: "Grand-Onglet", rows: 1999, cols: 4 });
  });

  it("each published sheet's tab holds exactly one native table block", async () => {
    const { doc } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    for (const tab of tabsOf(doc).tabs) {
      expect(tab.blocks).toHaveLength(1);
      expect(tab.blocks[0]?.kind).toBe("table");
    }
  });

  it("row counts: raw dimensions minus Grand-Onglet's two HIDDEN rows, minus '2026''s one fully EMPTY row (A1:D7 / A1:E9-1)", async () => {
    // STATUS.md 31/08/2026 (data-exposure fix): raw sheet3.xml (Grand-Onglet)
    // has rows r="5"/r="6" with `hidden="true"` — confirmed by direct XML
    // inspection, not assumed. Before the hidden-row/column guard, these 2
    // were leaking through (this test used to assert 2001).
    // '2026' also carries one fully empty row, removed. Grand-Onglet
    // is WITHHELD (its own test above): no tab, so no row count to read
    // here any more.
    const { doc } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    const published = tabsOf(doc).tabs.map((t) => (t.blocks[0]?.kind === "table" ? t.blocks[0].rows.length : -1));
    expect(published).toEqual([7, 8]);
  });

  // 27/10/2026: the note row (merged A9:D9, text in the first cell alone, empty E) now spans the WHOLE width (5), not the merge's 4
  // (`span-single-first.ts#spanSingleFirstRows`: a row whose only non-empty cell is the first one is one full-width cell). The
  // merge → colspan fact itself is pinned on a synthetic row that also holds text in E, right below.
  it("applies the merge in the '2026' tab (A9:D9) as a colspan cell on the last row — the whole width, E being empty", async () => {
    const { doc } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    const table = tabsOf(doc).tabs[1]!.blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    const lastRow = table.rows[table.rows.length - 1];
    expect(lastRow?.cells).toHaveLength(1); // was 2: the merge (colspan 4) + the empty column E
    expect(lastRow?.cells[0]?.colspan).toBe(5);
    expect(lastRow?.cells[0]?.rowspan).toBeUndefined();
    expect(lastRow?.cells[0]?.blocks[0]).toEqual({
      kind: "paragraph",
      inlines: [{ text: "Note : effectifs mesurés au dernier jour de chaque mois." }],
    });
  });

  it("applies a merge (A2:D2) inside the table as a colspan-4 cell on that data row", async () => {
    const bytes = await workbookOf([
      {
        name: "S",
        rows: [
          ["h1", "h2", "h3", "h4", "h5"],
          ["merged", "", "", "", "e2"],
          ["a", "b", "c", "d", "e"],
        ],
        merges: ["A2:D2"],
      },
    ]);
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const table = tabsOf(doc).tabs[0]!.blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    const row = table.rows[1];
    expect(row?.cells).toHaveLength(2); // 5 columns minus the 3 absorbed by the merge
    expect(row?.cells[0]?.colspan).toBe(4);
    expect(row?.cells[0]?.rowspan).toBeUndefined();
  });

  it("is deterministic: extracting the same bytes twice yields the same IRDoc", async () => {
    expect(await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS)).toEqual(
      await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS),
    );
  });

  it("EXTRACT_FAILED when SheetJS itself throws (e.g. an unsupported encrypted zip)", async () => {
    // SheetJS is lenient with plain garbage bytes (falls back to a 1-cell CSV
    // read rather than throwing) — a ZIP-signature prefix it can't actually
    // open is what reliably triggers its own parse error.
    const fakeEncryptedZip = new TextEncoder().encode(`PK\x03\x04${"x".repeat(200)}`);
    try {
      await extractXlsx(fakeEncryptedZip, undefined, PAGE_SIZED_LIMITS);
      expect.unreachable("expected extractXlsx to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(ReaderError);
      expect((error as ReaderError).code).toBe("EXTRACT_FAILED");
    }
  });
});

describe("extractXlsx (synthetic, structural properties not present in the current fixture)", () => {
  // STATUS.md 31/08/2026 (Gemini consultation, data-exposure finding): a
  // hidden row/column within an otherwise-visible tab must be ignored, same
  // treatment as an already-hidden tab (xlsx.ts's own module doc) — a
  // spreadsheet author who hides a salary column doesn't expect it
  // republished elsewhere. This reproduces the CURRENT (pre-fix) leak: the
  // hidden row/column's content is extracted anyway.
  it("ignores a hidden row and a hidden column, same treatment as a hidden tab", async () => {
    const bytes = await workbookOf([
      {
        name: "Sheet1",
        rows: [
          ["a1", "b1", "c1"],
          ["a2-secret", "b2-secret", "c2-secret"], // row index 1, to be hidden
          ["a3", "b3", "c3"],
        ],
        hiddenRows: [1],
        hiddenCols: [1],
      },
    ]);

    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tabs = doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;
    const table = tabs.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    const allText = table.rows.flatMap((r) =>
      r.cells.flatMap((c) => c.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text) : []))),
    );
    expect(allText).not.toContain("a2-secret");
    expect(allText).not.toContain("b2-secret");
    expect(allText).not.toContain("c2-secret");
    expect(allText).not.toContain("b1"); // hidden column, visible row
    expect(allText).not.toContain("b3");
    expect(allText).toEqual(["a1", "c1", "a3", "c3"]);
  });

  it("an empty CELL on a living row and column stays an empty-blocks cell — only fully empty LINES are dropped (étape 2)", async () => {
    // Two rows: the middle column lives through "y", so row 1's gap stays an
    // empty CELL. Before étape 2 the single-row version of this test proved
    // the same point; with one row the middle column was fully empty and is
    // now rightly dropped — the second row pins the distinction.
    const bytes = await workbookOf([
      {
        name: "Sheet1",
        rows: [
          ["a", "", "c"],
          ["x", "y", "z"],
        ],
      },
    ]);

    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tabs = doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;
    const table = tabs.tabs[0]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]?.cells.map((c) => c.blocks)).toEqual([
      [{ kind: "paragraph", inlines: [{ text: "a" }] }],
      [],
      [{ kind: "paragraph", inlines: [{ text: "c" }] }],
    ]);

    // And the single-row gap version: the fully empty middle column drops.
    const oneBytes = await workbookOf([{ name: "Sheet1", rows: [["a", "", "c"]] }]);
    const { doc: doc1 } = await extractXlsx(oneBytes, undefined, PAGE_SIZED_LIMITS);
    const table1 = (doc1.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!.blocks[0] as Extract<
      IRBlock,
      { kind: "table" }
    >;
    expect(table1.rows[0]?.cells.map((c) => c.blocks)).toEqual([
      [{ kind: "paragraph", inlines: [{ text: "a" }] }],
      [{ kind: "paragraph", inlines: [{ text: "c" }] }],
    ]);
  });
});

/** Spreadsheet lot, step 2 (decision (b), 25-27/09): rows and columns that are ENTIRELY empty are removed, at the
 *  edges AND inside — "empty" = no value to display, a cell that is only styled included. The two shapes were first
 *  measured on two real workbooks (28/09/2026), which stay in the publisher's private repository and are held by the
 *  application's tests: a sheet that declares A1:Z1000 while a few dozen rows carry text, in columns B to G only (the
 *  phantom 1000 x 26 grid), and a sheet whose column C is empty INSIDE the table. They are written here by hand. */
describe("extractXlsx — fully empty rows and columns are dropped, edges AND interior (étape 2)", () => {
  const tableOf = (doc: { blocks: IRBlock[] }, tab: number): Extract<IRBlock, { kind: "table" }> => {
    const t = doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;
    return t.tabs[tab]!.blocks[0] as Extract<IRBlock, { kind: "table" }>;
  };
  const widthOf = (table: Extract<IRBlock, { kind: "table" }>): number =>
    Math.max(...table.rows.map((r) => r.cells.reduce((n, c) => n + (c.colspan ?? 1), 0)));

  it("🔴 a declared 1000×26 grid publishes its 4 real rows and 6 real columns", async () => {
    const bytes = await workbookOf([
      {
        name: "Report",
        dimension: "A1:Z1000",
        rows: [
          [undefined, "Mileage report"],
          [],
          [undefined, "Date", "Description", "Use", "Start", "End", "Distance"],
          [undefined, "Monday", "Customer", "Work", "ten", "fifty", "forty"],
          [],
          [undefined, "Tuesday", "Training", "Work", "fifty", "sixty", "ten"],
        ],
      },
    ]);
    const table = tableOf((await extractXlsx(bytes)).doc, 0);
    // Failure values: 6 rows (the two empty ones kept), 7 or 26 columns (column A, or the declared grid, kept).
    expect(table.rows.length).toBe(4);
    // 27/10/2026: the title row (text in the first cell alone) stays row 0, as ONE cell over the 6 columns.
    expect(table.rows[0]!.cells).toHaveLength(1);
    expect(table.rows[0]!.cells[0]!.colspan).toBe(6);
    expect(table.rows[0]!.cells[0]!.blocks).toEqual([{ kind: "paragraph", inlines: [{ text: "Mileage report" }] }]);
    expect(widthOf(table)).toBe(6);
  });

  it("🔴 an interior empty column C is dropped (5 columns left), no row lost", async () => {
    const bytes = await workbookOf([
      {
        name: "Comparison",
        rows: [
          ["Cover", "First offer", undefined, "Second offer", "Third offer", "Note"],
          ["Hospital", "full", undefined, "partial", "full", "none"],
          ["Dental", "partial", undefined, "full", "none", "ceiling"],
        ],
      },
    ]);
    const table = tableOf((await extractXlsx(bytes)).doc, 0);
    // Failure value: 6 columns (the empty column kept as empty cells).
    expect(table.rows.length).toBe(3);
    expect(widthOf(table)).toBe(5);
  });
});

/** One `tabs` block, a sheet being a native tab; the display caps are counted after the removals. A sheet too large
 *  keeps its tab, which says why; the sheets past the cap are named under the block. */
describe("extractXlsx — one tabs block, display caps said in place (étape 3)", () => {
  it("🔴 fixture 04: one tabs block, two tabs named by their sheets; Grand-Onglet (1999×4) carries no tab at all, the dashboard says why (A7, 29/09/2026)", async () => {
    const { doc, notices } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    expect(doc.blocks).toHaveLength(1);
    const tabs = doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;
    expect(tabs.kind).toBe("tabs");
    expect(tabs.tabs.map((t) => t.title)).toEqual(["2025", "2026"]);
    expect(tabs.tabs[0]!.blocks[0]!.kind).toBe("table");
    expect(notices).toContainEqual({ kind: "sheet-withheld", sheet: "Grand-Onglet", rows: 1999, cols: 4 });
  });

  it("🔴 the sheet past the tab cap carries no tab at all, the dashboard names it: the first MAX_WORKBOOK_TABS publish (A7; cap 20 since 27/10/2026, was 6)", async () => {
    const count = MAX_WORKBOOK_TABS + 1;
    const bytes = await workbookOf(
      Array.from({ length: count }, (_, i) => ({ name: `Feuille ${i + 1}`, rows: [[`v${i + 1}`]] })),
    );

    const { doc, notices } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tabs = doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>;
    expect(tabs.tabs).toHaveLength(MAX_WORKBOOK_TABS);
    expect(doc.blocks).toHaveLength(1); // no more paragraph under the tabs block
    expect(notices).toContainEqual({ kind: "workbook-cap", sheets: [`Feuille ${count}`], more: 0 });
  });

  it("🔴 workbook-cap is BOUNDED (item E): 30 sheets with 100-character names past the cap give at most 5 names cut to 64 characters + a `more` count", async () => {
    const long = (i: number) => `${String(i).padStart(2, "0")}${"n".repeat(98)}`;
    // MAX_WORKBOOK_TABS + 30 sheets: 30 past the cap, so 5 named + 25 in `more`, whatever the cap's value (20 since 27/10/2026, was 6).
    const bytes = await workbookOf(
      Array.from({ length: MAX_WORKBOOK_TABS + 30 }, (_, i) => ({ name: long(i + 1), rows: [[`v${i + 1}`]] })),
    );
    const { notices } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const cap = notices.find((n) => n.kind === "workbook-cap");
    expect(cap).toBeDefined();
    if (cap?.kind !== "workbook-cap") return;
    expect(cap.sheets).toHaveLength(MAX_NOTICE_SHEET_NAMES);
    expect(cap.more).toBe(30 - MAX_NOTICE_SHEET_NAMES);
    for (const name of cap.sheets) expect(name.length).toBeLessThanOrEqual(MAX_NOTICE_SHEET_NAME_CHARS + 1); // + the « … »
    expect(JSON.stringify(cap).length).toBeLessThan(1000);
  });
});

/** A formula with no cached value gives an empty cell, never an invented « 0 », and the page says so under the table.
 *  Fixture 12 is made up without cached values. */
describe("extractXlsx — a formula without a cached value is EMPTY and said (étape 4)", () => {
  it("🔴 fixture 12: the five Total cells are empty, and a dashboard notice counts them (lot A, 29/09/2026 — no more line under the table)", async () => {
    const bytes = new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx/12-FABRIQUE-Formules-Sans-Valeur.xlsx")));
    const { doc, notices } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tab = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!;
    const table = tab.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    // Column D (« Total ») on the five data rows: empty blocks, no invented 0.
    // 27/10/2026: the fourth data row is « Grand total » alone in column A (its Total formula has no value): the row is ONE cell
    // over the 5 columns, so that empty cell is no longer published as a cell of its own (it holds nothing; the notice still counts it).
    const totals = table.rows.slice(1).map((r) => r.cells[3]?.blocks);
    expect(totals).toEqual([[], [], [], undefined, []]);
    expect(table.rows.map((r) => (r.cells.length === 1 ? r.cells[0]!.colspan : 1))).toEqual([1, 1, 1, 1, 5, 1]);
    expect(tab.blocks).toHaveLength(1); // the table alone
    expect(notices).toContainEqual({ kind: "formula-cells-empty", sheet: tab.title, count: 5 });
  });
});

/** `header: true` only on the row the file declares: the first `headerRowCount` rows of each Excel table's `ref`
 *  (`xl/tables/`), else the last frozen row (`ySplit`). Never the first published row by default (a title may sit
 *  above the real header), never deduced from the content. */
describe("extractXlsx — the declared header row, never a guessed one (étape 6)", () => {
  type TableBlock = Extract<IRBlock, { kind: "table" }>;
  const tableOf = (doc: { blocks: IRBlock[] }, tab: number): TableBlock =>
    (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[tab]!.blocks[0] as TableBlock;
  /** The text of a row that is ONE cell spanning `width` columns, else undefined (27/10/2026: a row whose only non-empty cell is
   *  the first one publishes so — `span-single-first.ts#spanSingleFirstRows`, shared with Google Sheets). */
  const spanningRow = (row: TableBlock["rows"][number] | undefined, width: number): string | undefined =>
    row !== undefined && row.cells.length === 1 && row.cells[0]!.colspan === width
      ? cellText(row.cells[0]!)
      : undefined;
  /** Row index → texts of its header cells, for the rows that carry any. */
  const headerRowsOf = (table: TableBlock): Map<number, string[]> => {
    const out = new Map<number, string[]>();
    table.rows.forEach((row, i) => {
      if (row.cells.some((c) => c.header === true))
        out.set(i, row.cells.filter((c) => c.header === true).map(cellText));
    });
    return out;
  };

  it("🔴 fixture 11 « Ventes 2026 » (Excel table + frozen pane): first published row is a header row", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    expect([...headerRowsOf(tableOf(doc, 0)).keys()]).toEqual([0]);
    // « Synthèse » declares nothing: no header.
    expect(headerRowsOf(tableOf(doc, 1)).size).toBe(0);
  });

  describe("synthetic declarations", () => {
    const TABLE_REL = `${REL_NS}/table`;
    /** Adds a pane and/or table parts to sheet1 of a built workbook. */
    const declared = async (
      bytes: Uint8Array,
      opts: { ySplit?: number | string; sheetViews?: string; tables?: string[] },
    ): Promise<Uint8Array> => {
      const zip = await JSZip.loadAsync(bytes);
      let xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
      if (opts.ySplit !== undefined)
        xml = xml.replace(
          "<dimension",
          `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${opts.ySplit}" state="frozen"/></sheetView></sheetViews><dimension`,
        );
      if (opts.sheetViews !== undefined)
        xml = xml.replace("<dimension", `<sheetViews>${opts.sheetViews}</sheetViews><dimension`);
      if (opts.tables) {
        xml = xml.replace(
          "</worksheet>",
          `<tableParts count="${opts.tables.length}">${opts.tables.map((_, i) => `<tablePart r:id="rIdT${i + 1}"/>`).join("")}</tableParts></worksheet>`,
        );
        zip.file(
          "xl/worksheets/_rels/sheet1.xml.rels",
          `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${opts.tables
            .map((_, i) => `<Relationship Id="rIdT${i + 1}" Type="${TABLE_REL}" Target="../tables/table${i + 1}.xml"/>`)
            .join("")}</Relationships>`,
        );
        opts.tables.forEach((attrs, i) =>
          zip.file(
            `xl/tables/table${i + 1}.xml`,
            `${XML_DECL}<table xmlns="${MAIN_NS}" id="${i + 1}" name="T${i + 1}" displayName="T${i + 1}" ${attrs}/>`,
          ),
        );
      }
      zip.file("xl/worksheets/sheet1.xml", xml);
      return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
    };
    const grid = [["Title"], ["a", "b"], ["1", "2"], ["3", "4"]];

    it("🔴 headerRowCount=2 marks the first TWO rows of the ref; headerRowCount=0 marks none", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      const two = await extractXlsx(await declared(base, { tables: ['ref="A2:B4" headerRowCount="2"'] }));
      expect([...headerRowsOf(tableOf(two.doc, 0)).keys()]).toEqual([1, 2]);
      // 27/10/2026: « Title » (first cell alone) is row 0, ONE cell over the 2 columns, and no header.
      expect(spanningRow(tableOf(two.doc, 0).rows[0], 2)).toBe("Title");
      const none = await extractXlsx(await declared(base, { tables: ['ref="A2:B4" headerRowCount="0"'] }));
      expect(headerRowsOf(tableOf(none.doc, 0)).size).toBe(0);
    });

    it("🔴 a frozen pane alone marks its LAST frozen row (ySplit=2 → row 2), not the first", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      const { doc } = await extractXlsx(await declared(base, { ySplit: 2 }));
      expect([...headerRowsOf(tableOf(doc, 0)).keys()]).toEqual([1]);
      expect(headerRowsOf(tableOf(doc, 0)).get(1)).toEqual(["a", "b"]);
      expect(spanningRow(tableOf(doc, 0).rows[0], 2)).toBe("Title"); // 27/10/2026: the title row spans, it is not the header
    });

    it("🔴 a frozen pane whose last row is hidden abstains (nothing promoted in its place)", async () => {
      const base = await workbookOf([{ name: "S", rows: grid, hiddenRows: [1] }]);
      const { doc } = await extractXlsx(await declared(base, { ySplit: 2 }));
      expect(headerRowsOf(tableOf(doc, 0)).size).toBe(0);
    });

    // Revue 02/10/2026 S1/S2/L10. ECMA-376 / ISO 29500-1 §18.3.1.66: ySplit = rows visible in the top pane,
    // counted from sheetView@topLeftCell; pane@topLeftCell is the BOTTOM pane's scroll position, never the split.
    // https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.pane?view=openxml-3.0.1
    const frozenView = (ySplit: string, view = "", pane = ""): string =>
      `<sheetView workbookViewId="0"${view}><pane ySplit="${ySplit}"${pane} state="frozen"/></sheetView>`;
    const bigGrid = Array.from({ length: 10 }, (_, i) => [`r${i + 1}`, "x"]);

    it("🔴 a ySplit that is not a plain decimal integer in 1..1048576 declares nothing (no row walk, no header)", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      for (const ySplit of ["10000000", "1048577", "1e3", "2.0", "0x2", " 2"]) {
        const started = performance.now();
        const { doc } = await extractXlsx(await declared(base, { ySplit }));
        expect(headerRowsOf(tableOf(doc, 0)).size, `ySplit=${JSON.stringify(ySplit)}`).toBe(0);
        expect(
          performance.now() - started,
          `ySplit=${JSON.stringify(ySplit)}: no walk of the row numbers`,
        ).toBeLessThan(1000);
      }
    });

    it("🔴 ySplit=1048576 (the Excel row limit) is accepted and lands on the last kept row", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      const { doc } = await extractXlsx(await declared(base, { ySplit: 1048576 }));
      expect([...headerRowsOf(tableOf(doc, 0)).keys()]).toEqual([3]);
    });

    it("🔴 the top pane starts at sheetView@topLeftCell: topLeftCell=A4 with ySplit=2 freezes rows 4-5, so the header is row 5", async () => {
      const base = await workbookOf([{ name: "S", rows: bigGrid }]);
      const { doc } = await extractXlsx(await declared(base, { sheetViews: frozenView("2", ' topLeftCell="A4"') }));
      expect([...headerRowsOf(tableOf(doc, 0)).keys()]).toEqual([4]);
    });

    it("🔴 a scrolled bottom pane (pane@topLeftCell=A8) does not move the split: ySplit=2 still ends on row 2", async () => {
      const base = await workbookOf([{ name: "S", rows: bigGrid }]);
      const { doc } = await extractXlsx(await declared(base, { sheetViews: frozenView("2", "", ' topLeftCell="A8"') }));
      expect([...headerRowsOf(tableOf(doc, 0)).keys()]).toEqual([1]);
    });

    it("🔴 a table whose ref runs past the last Excel row, or a headerRowCount far beyond it, is never expanded", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      const started = performance.now();
      const huge = await extractXlsx(await declared(base, { tables: ['ref="A1:A2000000" headerRowCount="2000000"'] }));
      expect(headerRowsOf(tableOf(huge.doc, 0)).size, "ref beyond row 1048576: ignored").toBe(0);
      expect(performance.now() - started).toBeLessThan(1000);
      const clamped = await extractXlsx(await declared(base, { tables: ['ref="A2:B3" headerRowCount="9999999"'] }));
      expect([...headerRowsOf(tableOf(clamped.doc, 0)).keys()], "count clamped to the ref's height").toEqual([1, 2]);
    });

    it("🔴 readDeclaredHeaderRows hands back spans, never one entry per row", async () => {
      const base = await workbookOf([{ name: "S", rows: grid }]);
      const { pkg, sheets } = await readWorkbookSheets(
        await declared(base, { tables: ['ref="A1:A1048576" headerRowCount="9999999"'] }),
      );
      const got = await readDeclaredHeaderRows(pkg, sheets[0]!.partPath);
      expect(got.tableSpans).toEqual([{ start: 0, count: 1048576 }]);
    });

    it("🔴 a declared header row that is not published (hidden) abstains: no header at all", async () => {
      const base = await workbookOf([{ name: "S", rows: grid, hiddenRows: [1] }]);
      const { doc } = await extractXlsx(await declared(base, { tables: ['ref="A2:B4"'] }));
      expect(headerRowsOf(tableOf(doc, 0)).size).toBe(0);
    });
  });
});

/** Lot tableurs step 7: cell links — the EXTERNAL target published RAW
 *  (the scheme filter lives at the door, `storage.ts#PUBLISHED_SCHEME`,
 *  arbitrage 5: no copy here); the internal form `#'Feuille'!A1` refused by
 *  `lib/ir.ts#INTERNAL_HREF_FORM` — plain text. Witness: fixture 11, E2
 *  external and E3 internal (generator). */
describe("extractXlsx — cell links: external raw, internal form refused (étape 7)", () => {
  it("🔴 fixture 11: E2 carries its external href verbatim, E3 stays plain text", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const table = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!.blocks[0] as Extract<
      IRBlock,
      { kind: "table" }
    >;
    const inlineOf = (row: number, col: number) =>
      (table.rows[row]!.cells[col]!.blocks[0] as Extract<IRBlock, { kind: "paragraph" }>).inlines[0]!;
    expect(inlineOf(1, 4)).toEqual({ text: "Voir le rapport", href: "https://example.com/rapport-janvier" });
    expect(inlineOf(2, 4)).toEqual({ text: "Aller à la synthèse" });
  });
});

/** A `.xlsx` chart becomes a `chart` block (title and table of the cached data) under the table; a chart sheet becomes
 *  a tab holding its title and chart, and counts towards the tab cap. */
describe("extractXlsx — charts under the table, chart-only tabs published (étape 8)", () => {
  it("🔴 fixture 11: « Ventes 2026 » carries its chart under the table; « Graphique » is a chart-only tab", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tabs = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs;
    expect(tabs.map((t) => t.title)).toEqual(["Ventes 2026", "Synthèse", "Graphique"]);

    const ventes = tabs[0]!.blocks;
    expect(ventes[0]!.kind).toBe("table");
    const posed = ventes.find<Extract<IRBlock, { kind: "chart" }>>((b) => b.kind === "chart");
    expect(posed?.title).toBe("Chiffre d'affaires mensuel");
    expect((posed?.data.rows.length ?? 0) > 0).toBe(true);

    const graphique = tabs[2]!.blocks;
    expect(graphique).toHaveLength(1);
    const only = graphique[0] as Extract<IRBlock, { kind: "chart" }>;
    expect(only.kind).toBe("chart");
    expect(only.title).toBe("Marge mensuelle");
    expect(only.data.rows.length > 0).toBe(true);
  });
});

/** Spreadsheet lot, step 9 (decision (h)): the anchored pictures of the workbook — `image` blocks UNDER the table,
 *  with no alternative text, through the shared collector. Witness: the made-up workbook with one anchored picture
 *  (`xlsx-anchored-picture.test.support.ts`); the real one the step was written on (a signature anchored in a cell)
 *  stays in the publisher's private repository, held by the application's tests. */
describe("extractXlsx — anchored pictures under the table (étape 9)", () => {
  it("🔴 an anchored picture reaches the page as an image block, and its bytes ride in images[]", async () => {
    const { doc, images } = await extractXlsx(await anchoredPictureWorkbook(), undefined, PAGE_SIZED_LIMITS);
    const tab = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!;
    // The table comes first, the picture under it.
    expect(tab.blocks[0]?.kind).toBe("table");
    const image = tab.blocks.find<Extract<IRBlock, { kind: "image" }>>((b) => b.kind === "image");
    expect(image).toBeDefined();
    expect(image!.alt).toBeUndefined();
    expect(images).toHaveLength(1);
    expect(images[0]!.sha256).toBe(image!.ref.sha256);
    expect(images[0]!.bytes.length > 0).toBe(true);
  });
});

/** An encrypted Office file is a CFB container, the same whatever Office format it wraps. */
describe("extractXlsx — a password-protected workbook is NAMED, never « damaged » (constat 2)", () => {
  it("🔴 the real encrypted fixture refuses as DOCUMENT_PASSWORD_PROTECTED", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(here, "../../fixtures/docx/refuses/09-Fiche-Contact-Protegee.docx")),
    );
    const error = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS).then(
      () => undefined,
      (e) => e as ReaderError,
    );
    expect(error?.code).toBe("DOCUMENT_PASSWORD_PROTECTED");
  });
});

/** Counted on the whole XML, hidden rows and removed columns included, the line would lie and reveal the structure of
 *  what is hidden. */
describe("extractXlsx — the formula line counts only what the PUBLISHED grid shows (constat 8)", () => {
  it("🔴 a value-less formula alone in a fully empty column: the column drops, and NO line is printed", async () => {
    // B1 is a formula, no cached `<v>` — its column holds nothing else.
    const bytes = await workbookOf([
      {
        name: "Sheet1",
        dimension: "A1:B2",
        rawSheetDataXml:
          '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1"><f>1+1</f></c></row>' +
          '<row r="2"><c r="A2" t="inlineStr"><is><t>b</t></is></c></row>',
      },
    ]);

    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tab = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!;
    expect(tab.blocks).toHaveLength(1); // the table alone — no phantom count under it
    const table = tab.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells).toHaveLength(1); // column B dropped
  });
});

describe("extractXlsx — computed empty strings, hidden-column rows, huge declared ranges (revue 29/09)", () => {
  const firstTab = (doc: { blocks: IRBlock[] }) => (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!;

  it('🔴 constat 3: a t="str" formula with an EMPTY cached <v> is a computed empty string — no « no computed value » line', async () => {
    // The shape Excel itself saves for `=IF(…,"",…)` answering "".
    // 27/10/2026: a row whose only non-empty cell is the first one is published as ONE spanning cell, which would absorb the
    // computed-empty cell and leave nothing to observe. The formula row therefore carries text in C too: its empty B cell is
    // still a cell of the table (expected [] differs from a « 0 » or a notice).
    const bytes = await workbookOf([
      {
        name: "Sheet1",
        dimension: "A1:C2",
        rawSheetDataXml:
          '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c>' +
          '<c r="B1" t="str"><f>IF(A1="a","","")</f><v></v></c>' +
          '<c r="C1" t="inlineStr"><is><t>c</t></is></c></row>' +
          '<row r="2"><c r="A2" t="inlineStr"><is><t>b</t></is></c>' +
          '<c r="B2" t="inlineStr"><is><t>y</t></is></c>' +
          '<c r="C2" t="inlineStr"><is><t>z</t></is></c></row>',
      },
    ]);
    const { doc, notices } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tab = firstTab(doc);
    expect(tab.blocks).toHaveLength(1); // the table alone
    const table = tab.blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows[0]!.cells).toHaveLength(3);
    expect(table.rows[0]!.cells[1]!.blocks).toEqual([]);
    expect(notices.some((n) => n.kind === "formula-cells-empty")).toBe(false);
  });

  it("🔴 constat 4: a row whose only text sits in a HIDDEN column is not published (not even empty)", async () => {
    const bytes = await workbookOf([
      {
        name: "Sheet1",
        rows: [
          ["a", "x"],
          ["", "secret"],
        ],
        hiddenCols: [1],
      },
    ]);
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const table = firstTab(doc).blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows).toHaveLength(1);
  });

  // By construction, not by measurement alone: our reader (`xlsx-cells.ts`)
  // walks the cells the sheet's XML actually HAS — the `<row>`/`<c>`
  // elements of `<sheetData>` — never the declared `dimension`, so a huge
  // declared range costs it nothing to begin with (no SheetJS-style walk of
  // ~27 M addresses to protect against any more). The time bound is kept as
  // a regression witness, not as the live defence.
  it("🔴 constat 10: a huge declared range with two real cells publishes its two cells, fast", async () => {
    const bytes = await workbookOf([{ name: "Sheet1", rows: [["a"], ["b"]], dimension: "A1:Z1048576" }]);
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const table = firstTab(doc).blocks[0] as Extract<IRBlock, { kind: "table" }>;
    expect(table.rows.map((r) => r.cells.length)).toEqual([1, 1]);
  });
});

// Audit of 07/10/2026: one damaged auxiliary part must not refuse the whole workbook — a presentation degrades a damaged
// chart to its marker, a workbook refused everything. The sheets' own text is published; the damaged chart is not invented.
describe("extractXlsx — a damaged chart part refuses nothing else", () => {
  it("🔴 a chart XML that does not parse leaves the workbook published, its cells included", async () => {
    const real = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const zip = await JSZip.loadAsync(real);
    zip.file("xl/charts/chart1.xml", '<?xml version="1.0"?<broken><c:a/>');
    const damagedBytes = await zip.generateAsync({ type: "uint8array" });
    const { doc } = await extractXlsx(damagedBytes, undefined, PAGE_SIZED_LIMITS);
    expect(JSON.stringify(doc.blocks)).toContain("Janvier");
  });

  // 07/10/2026: the loss is said, with the notice a Word document already uses for a chart kept out.
  it("🔴 the damaged chart is said: one chart-withheld notice", async () => {
    const real = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const zip = await JSZip.loadAsync(real);
    zip.file("xl/charts/chart1.xml", '<?xml version="1.0"?<broken><c:a/>');
    const { notices } = await extractXlsx(await zip.generateAsync({ type: "uint8array" }));
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
  });

  // 07/10/2026: a MISSING part is said like a damaged one, the same notice in every format.
  it("🔴 a chart whose part is missing is said: one chart-withheld notice per chart", async () => {
    const real = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const zip = await JSZip.loadAsync(real);
    zip.remove("xl/charts/chart1.xml");
    const { notices } = await extractXlsx(await zip.generateAsync({ type: "uint8array" }));
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
  });

  it("🔴 a chart whose relationship does not resolve is said too", async () => {
    const real = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const zip = await JSZip.loadAsync(real);
    zip.file(
      "xl/drawings/_rels/drawing1.xml.rels",
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
    );
    const { notices } = await extractXlsx(await zip.generateAsync({ type: "uint8array" }));
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
  });
});

// Audit of 07/10/2026: a tab whose only content is a chart left out (damaged or withheld) is never an empty frame: like a
// Google Sheet tab and an empty pptx slide, it carries one empty paragraph.
describe("extractXlsx — a chart sheet whose chart is left out is not an empty tab", () => {
  it("🔴 every published tab has at least one block", async () => {
    const real = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx")),
    );
    const zip = await JSZip.loadAsync(real);
    for (const part of ["xl/charts/chart1.xml", "xl/charts/chart2.xml"])
      zip.file(part, '<?xml version="1.0"?<broken><c:a/>');
    const { doc } = await extractXlsx(await zip.generateAsync({ type: "uint8array" }));
    const tabs = doc.blocks.find((b): b is Extract<IRBlock, { kind: "tabs" }> => b.kind === "tabs")!;
    expect(tabs.tabs.map((t) => [t.title, t.blocks.length > 0])).toEqual(tabs.tabs.map((t) => [t.title, true]));
  });
});

// 14/10/2026: « 2 formula cells have no computed value » looked false, since Excel shows 35 and 8. The file
// (written by openpyxl 3.1.5) does hold `<f>SUM(B2:B4)</f><v></v>` in Metrics B5 and C5: Excel recomputes on opening,
// the file itself carries no value. openpyxl only hides the formulas when loaded with `data_only=True`.
describe("extractXlsx — a workbook saved by openpyxl with two SUM and no cached value (fixtures/xlsx-reel 18)", () => {
  it("says the two cells of Metrics, and nothing for the three sheets without formula", async () => {
    const fixture = readFileSync(
      join(
        dirname(fileURLToPath(import.meta.url)),
        "../../fixtures/xlsx-reel/18-Operations-openpyxl-SUM-sans-valeur.xlsx",
      ),
    );
    const { notices } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    expect(notices.filter((n) => n.kind === "formula-cells-empty")).toEqual([
      { kind: "formula-cells-empty", sheet: "Metrics", count: 2 },
    ]);
  });
});

// 14/10/2026: the « Total » row of « 07 Operations.xlsx » is bold in Excel and Google Sheets (cellXfs 3 → font 2,
// `<b val="1"/>`) and was published plain: the reader never looked at a cell's font.
describe("extractXlsx — a bold cell is published bold (fixtures/xlsx-reel 19)", () => {
  it("the Total row of Metrics carries the strong mark, a plain row does not", async () => {
    const fixture = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), "../../fixtures/xlsx-reel/19-Operations-Excel-recalcule.xlsx"),
    );
    const { doc } = await extractXlsx(fixture, undefined, PAGE_SIZED_LIMITS);
    const metrics = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs.find((t) => t.title === "Metrics")!;
    const table = metrics.blocks.find((b): b is Extract<IRBlock, { kind: "table" }> => b.kind === "table")!;
    const inlinesOf = (row: number) =>
      table.rows[row]!.cells.map((c) => (c.blocks[0] as Extract<IRBlock, { kind: "paragraph" }>).inlines[0]!);
    const total = table.rows.findIndex((r) => JSON.stringify(r).includes('"Total"'));
    expect(inlinesOf(total).map((i) => [i.text, i.marks ?? []])).toEqual([
      ["Total", ["strong"]],
      ["35", ["strong"]],
      ["8", ["strong"]],
    ]);
    expect(inlinesOf(total - 1).map((i) => i.marks ?? [])).toEqual([[], [], []]);
  });
});

describe("language sample — tab slots mirror the page's (revue 27/10/2026, tour 7)", () => {
  it("🔴 a sheet whose part is GONE frees its slot in the language sample too: the sheet sitting at the cap still feeds it", async () => {
    // FAILS WITH (before the fix): the gone sheet consumed a sample slot, the English text of the last sheet was never
    // sampled, the page language (fr) stood — the boolean printed VRAI although the workbook is English.
    // The sample reads SHARED strings only, so the English text goes through xl/sharedStrings.xml (added by hand below).
    const english =
      "The annual report was approved by the board after the meeting and the plan covers the next year of work. " +
      "Every department wrote down what it shipped during the quarter and what remains open for the coming months, " +
      "and the committee agreed to publish the whole document on the internal site before the end of the week.";
    const last = MAX_WORKBOOK_TABS + 1;
    const sheets: SyntheticSheet[] = [
      { name: "Gone", rows: [["x"]] },
      { name: "Bool", rawSheetDataXml: '<row r="1"><c r="A1" t="b"><v>1</v></c></row>' },
      ...Array.from({ length: MAX_WORKBOOK_TABS - 2 }, (_, i) => ({
        name: `Mid${i}`,
        rawSheetDataXml: `<row r="1"><c r="A1"><v>${i + 1}</v></c></row>`,
      })),
      { name: "Last", rawSheetDataXml: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' },
    ];
    const zip = await JSZip.loadAsync(await workbookOf(sheets));
    zip.remove("xl/worksheets/sheet1.xml");
    zip.file(
      "xl/sharedStrings.xml",
      `${XML_DECL}<sst xmlns="${MAIN_NS}" count="1" uniqueCount="1"><si><t>${english}</t></si></sst>`,
    );
    const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
    zip.file(
      "xl/_rels/workbook.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rId${last + 1}" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
      ),
    );
    const { doc } = await extractXlsx(new Uint8Array(await zip.generateAsync({ type: "uint8array" })), "fr");
    const tabs = doc.blocks[0];
    if (tabs?.kind !== "tabs") throw new Error("expected tabs");
    const bool = tabs.tabs.find((t) => t.title === "Bool")!.blocks[0];
    if (bool?.kind !== "table") throw new Error("expected table");
    expect(cellText(bool.rows[0]!.cells[0]!)).toBe("TRUE");
    expect(tabs.tabs.map((t) => t.title)).toContain("Last");
  });
});
