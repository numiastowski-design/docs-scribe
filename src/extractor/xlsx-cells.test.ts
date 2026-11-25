/**
 * `xlsx-cells.ts`, our cell reader, asserted directly: no comparison with
 * SheetJS any more (removed 29/09/2026) — this file restates what
 * `xlsx-cells.parity.test.ts` used to prove by difference. Kept from the
 * parity study (29/09/2026, point 2): the cell count per real fixture
 * (31 264 in all, of which 23 442 on seventeen of the study's original 19 — none has
 * moved since; the FAQ file of 27/10/2026 adds 7 822; the two other workbooks of the
 * study, 378 cells, stay in the publisher's private repository and are held there), and the two fixes
 * SheetJS did not bring (built-in format 14, non-ASCII hyperlink), now
 * pinned one by one instead of read off a diff (sources: MIKSUDOC
 * « Spécification fonctionnelle » §9.4, § Sources de référence).
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { publicFixtures } from "./public-fixtures.test.support";
import { extractXlsx } from "./xlsx";
import {
  isColumnHidden,
  findHyperlinkTargetAt,
  parseCellReference,
  readWorksheetCells,
  readSharedStringsAndStyles,
  type WorksheetCells,
  type SharedStringsAndStyles,
} from "./xlsx-cells";
import { readWorkbookSheets, type WorkbookSheet } from "./xlsx-parts";
import { type OpcPackage } from "./ooxml/package";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "../../fixtures");
const FIXTURES = publicFixtures(".xlsx")
  .filter((path) => path.startsWith("fixtures/xlsx/") || path.startsWith("fixtures/xlsx-reel/"))
  .map((path) => path.slice("fixtures/".length));

/** Every cell our reader finds across the workbook's worksheets (chartsheets
 *  carry no grid). */
const cellsOf = async (bytes: Uint8Array): Promise<number> => {
  const { pkg, sheets } = await readWorkbookSheets(bytes);
  const workbook = await readSharedStringsAndStyles(pkg);
  let total = 0;
  for (const info of sheets) {
    if (info.kind !== "worksheet") continue;
    const xml = await pkg.readXml(info.partPath);
    if (!xml) continue;
    total += (await readWorksheetCells(pkg, info.partPath, xml, workbook)).cells.size;
  }
  return total;
};

/** One named sheet, read on its own — for the two spot-checks below. */
const sheetOf = async (
  bytes: Uint8Array,
  name: string,
): Promise<{
  pkg: OpcPackage;
  sheets: WorkbookSheet[];
  workbook: SharedStringsAndStyles;
  sheet: Awaited<ReturnType<typeof readWorksheetCells>>;
}> => {
  const { pkg, sheets } = await readWorkbookSheets(bytes);
  const workbook = await readSharedStringsAndStyles(pkg);
  const info = sheets.find((s) => s.name === name)!;
  const xml = await pkg.readXml(info.partPath);
  return { pkg, sheets, workbook, sheet: await readWorksheetCells(pkg, info.partPath, xml!, workbook) };
};

// Cells holding a value, per fixture — guards the count itself: an empty
// walk would compare nothing. 31 264 in all (23 442 over seventeen of the original 19 fixtures, study of 29/09/2026, point 2; the FAQ adds 7 822,
// widened 9e3d3d3/9ede124 with real, open-data and Google Sheets exports).
const FIXTURE_CELLS: Record<string, number> = {
  "xlsx/04-Effectifs-Atria.xlsx": 10_065,
  "xlsx/09-Effectifs-Dates.xlsx": 9,
  "xlsx/10-Commentaires-Cellules.xlsx": 6,
  "xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx": 42,
  "xlsx/12-FABRIQUE-Formules-Sans-Valeur.xlsx": 20,
  "xlsx-reel/15-Collectivites-Elus-Locaux.xlsx": 3_413,
  "xlsx-reel/16-Collectivites-Fiscalite-Locale.xlsx": 4_501,
  "xlsx-reel/17-GSheets-Essai13-Graphique.xlsx": 14,
  "xlsx-reel/18-Operations-openpyxl-SUM-sans-valeur.xlsx": 55,
  "xlsx-reel/19-Operations-Excel-recalcule.xlsx": 57,
  "xlsx-reel/18-GSheets-Essai11-Graphique.xlsx": 14,
  "xlsx-reel/21-GSheets-Planning-Formations.xlsx": 42,
};

describe("xlsx-cells — cell counts on the eighteen fixtures (was: parity with SheetJS)", () => {
  it("the eighteen fixtures are all here", () => {
    expect(FIXTURES).toHaveLength(18);
  });

  for (const f of FIXTURES) {
    it(`${f}: reads exactly ${FIXTURE_CELLS[f]} cells`, async () => {
      expect(await cellsOf(new Uint8Array(readFileSync(join(root, f))))).toBe(FIXTURE_CELLS[f]);
    });
  }
});

// Why each correction below is OURS being right (sources:
// `docs/SOURCES-OFFICIELLES.md`). Formerly two
// entries in the parity diff list; pinned directly now that there is
// nothing left to diff against.
describe("xlsx-cells — the two corrections, pinned directly", () => {
  it("🔴 fixture 09 Effectifs-Dates!Effectifs B2/B3: built-in format 14 as ECMA-376 §18.8.30 writes it (mm-dd-yy) — DECIDED (29/09): the reading language's short date, rendered US today", async () => {
    const { sheet } = await sheetOf(
      new Uint8Array(readFileSync(join(root, "xlsx/09-Effectifs-Dates.xlsx"))),
      "Effectifs",
    );
    // `format` stays the faithful ECMA-376 code; `isBuiltinNumFmt14` is the ID-based
    // flag `extractor/xlsx.ts#writeCellText` reads to render `m/d/yy` instead.
    expect(sheet.cells.get("B2")?.formatCode).toBe("mm-dd-yy");
    expect(sheet.cells.get("B2")?.isBuiltinNumFmt14).toBe(true);
    expect(sheet.cells.get("B3")?.formatCode).toBe("mm-dd-yy");
    expect(sheet.cells.get("B3")?.isBuiltinNumFmt14).toBe(true);
  });

  it("🔴 fixture 11 Ventes 2026!E3: the internal hyperlink's non-ASCII sheet name is read once, not double-decoded (Open XML SDK, Hyperlink.Location)", async () => {
    const { sheet } = await sheetOf(
      new Uint8Array(readFileSync(join(root, "xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx"))),
      "Ventes 2026",
    );
    const at = parseCellReference("E3")!;
    expect(findHyperlinkTargetAt(sheet, at.rowIndex, at.columnIndex)).toBe("#'Synthèse'!A1");
  });

  it("🔴 a workbook that REDEFINES built-in id 14 keeps its own code: builtinDate14 is false, the id alone never triggers the override", async () => {
    const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        "</Types>",
    );
    zip.file(
      "_rels/.rels",
      `${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      `${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
    );
    zip.file(
      "xl/workbook.xml",
      `${decl}<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Feuille" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    // Redefines id 14 with an explicit, different code — the workbook's own
    // choice, never confused with the built-in default.
    zip.file(
      "xl/styles.xml",
      `${decl}<styleSheet xmlns="${main}"><numFmts count="1"><numFmt numFmtId="14" formatCode="yyyy-mm-dd"/></numFmts>` +
        '<cellXfs count="1"><xf numFmtId="14"/></cellXfs></styleSheet>',
    );
    zip.file(
      "xl/worksheets/sheet1.xml",
      `${decl}<worksheet xmlns="${main}" xmlns:r="${rel}"><dimension ref="A1"/><sheetData>` +
        '<row r="1"><c r="A1"><v>46173</v></c></row></sheetData></worksheet>',
    );
    const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
    const { sheet } = await sheetOf(bytes, "Feuille");
    expect(sheet.cells.get("A1")?.formatCode).toBe("yyyy-mm-dd");
    expect(sheet.cells.get("A1")?.isBuiltinNumFmt14).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// What the fixtures do not carry, fabricated in memory: phonetic runs, the
// `_xHHHH_` escape, a date cell (`t="d"`), the 1904 date system, a column
// span hidden, a range hyperlink, a location-only one, one with both.

interface Namespaces {
  main: string;
  rel: string;
  relType: string;
}
const TRANSITIONAL: Namespaces = {
  main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  relType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

const fabricated = async (ns: Namespaces): Promise<Uint8Array> => {
  const zip = new JSZip();
  const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  zip.file(
    "[Content_Types].xml",
    `${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      "</Types>",
  );
  const rels = (items: string): string =>
    `${decl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  zip.file(
    "_rels/.rels",
    rels(`<Relationship Id="rId1" Type="${ns.relType}/officeDocument" Target="xl/workbook.xml"/>`),
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    rels(
      `<Relationship Id="rId1" Type="${ns.relType}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${ns.relType}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${ns.relType}/sharedStrings" Target="sharedStrings.xml"/>`,
    ),
  );
  zip.file(
    "xl/workbook.xml",
    `${decl}<workbook xmlns="${ns.main}" xmlns:r="${ns.rel}"><workbookPr date1904="1"/>` +
      '<sheets><sheet name="Feuille" sheetId="1" r:id="rId1"/></sheets></workbook>',
  );
  zip.file(
    "xl/styles.xml",
    `${decl}<styleSheet xmlns="${ns.main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>` +
      '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="4"/><xf numFmtId="49"/></cellXfs></styleSheet>',
  );
  zip.file(
    "xl/sharedStrings.xml",
    `${decl}<sst xmlns="${ns.main}" count="3" uniqueCount="3">` +
      '<si><r><t>Tō</t></r><r><t xml:space="preserve">kyō </t></r><rPh sb="0" eb="1"><t>トウ</t></rPh><phoneticPr fontId="0"/></si>' +
      "<si><t>avant_x0009_après _x005F_x0041_</t></si>" +
      "<si><t>Libellé</t></si></sst>",
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `${decl}<worksheet xmlns="${ns.main}" xmlns:r="${ns.rel}"><dimension ref="A1:D4"/>` +
      '<cols><col min="5" max="7" width="9" hidden="1"/></cols><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>46294</v></c><c r="C1" t="b"><v>1</v></c>' +
      '<c r="D1" t="inlineStr"><is><r><t>en </t></r><r><t>ligne</t></r></is></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="d" s="1"><v>2026-09-29T12:00:00</v></c><c r="C2" t="e"><v>#N/A</v></c>' +
      '<c r="D2" s="2"><v>-1234.5</v></c></row>' +
      '<row r="3" hidden="1"><c r="A3" t="s" s="3"><v>2</v></c><c r="C3" t="str"><f>IF(1,"","")</f><v></v></c>' +
      '<c r="D3"><f>A1*2</f></c></row>' +
      '<row r="4"><c r="A4" s="2"/><c r="B4"><v>0.1</v></c></row>' +
      '</sheetData><mergeCells count="1"><mergeCell ref="C4:D4"/></mergeCells>' +
      '<hyperlinks><hyperlink ref="A1:B1" r:id="rId1"/><hyperlink ref="D1" location="Feuille!A1"/>' +
      '<hyperlink ref="C1" r:id="rId1" location="frag"/></hyperlinks></worksheet>',
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    rels(
      `<Relationship Id="rId1" Type="${ns.relType}/hyperlink" Target="https://example.com/x" TargetMode="External"/>`,
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
};

describe("xlsx-cells — a fabricated transitional workbook, every feature read as the norm writes it", () => {
  it("reads each feature as the norm writes it", async () => {
    const { pkg, sheets } = await readWorkbookSheets(await fabricated(TRANSITIONAL));
    const workbook = await readSharedStringsAndStyles(pkg);
    const xml = await pkg.readXml(sheets[0]!.partPath);
    const sheet = await readWorksheetCells(pkg, sheets[0]!.partPath, xml!, workbook);
    expect(sheet.cells.size).toBe(11);
    expect(workbook.date1904).toBe(true);
    expect(sheet.cells.get("A1")?.value).toEqual({ t: "s", v: "Tōkyō " }); // rPh excluded
    expect(sheet.cells.get("A2")?.value).toEqual({ t: "s", v: "avant\taprès _x0041_" }); // §22.9.2.19
    // 2026-09-29T12:00 in the 1904 system: 45 732 days after 1904-01-01, plus half a day.
    expect(sheet.cells.get("B2")?.value).toEqual({ t: "n", v: 44832.5 });
    expect(sheet.cells.get("B2")?.formatCode).toBe("dd/mm/yyyy");
    expect(sheet.cells.get("C3")?.value).toEqual({ t: "s", v: "" }); // computed empty string, decision (f)
    expect(sheet.cells.has("D3")).toBe(false); // a formula without a cached value
    expect(sheet.formulasWithoutValue).toEqual([{ rowIndex: 2, columnIndex: 3 }]); // …counted by the same walk (review finding 12)
    expect(sheet.cells.has("A4")).toBe(false); // a style alone
    expect(sheet.mergeCells).toEqual([
      { start: { rowIndex: 3, columnIndex: 2 }, end: { rowIndex: 3, columnIndex: 3 } },
    ]); // C4:D4
    expect(findHyperlinkTargetAt(sheet, 0, 1)).toBe("https://example.com/x"); // B1, inside A1:B1
    expect(findHyperlinkTargetAt(sheet, 0, 2)).toBe("https://example.com/x#frag");
    expect(findHyperlinkTargetAt(sheet, 0, 3)).toBe("#Feuille!A1");
    expect([4, 5, 6, 7].map((c) => isColumnHidden(sheet, c))).toEqual([true, true, true, false]);
    expect([...sheet.hiddenRows]).toEqual([2]);
  });
});

// ---------------------------------------------------------------------------
// The same workbook in the ECMA-376 STRICT conformance class (Part 1 §A,
// namespaces under http://purl.oclc.org/ooxml/): `ooxml/xml.ts` maps
// both classes to the same canonical names, so our reader sees no difference.
const STRICT: Namespaces = {
  main: "http://purl.oclc.org/ooxml/spreadsheetml/main",
  rel: "http://purl.oclc.org/ooxml/officeDocument/relationships",
  relType: "http://purl.oclc.org/ooxml/officeDocument/relationships",
};

describe("xlsx-cells — isColHidden over many spans (review finding 21)", () => {
  const sheetWith = (hiddenCols: { min: number; max: number }[]): WorksheetCells => ({
    cells: new Map(),
    dimensionRange: undefined,
    mergeCells: [],
    hiddenRows: new Set(),
    hiddenColumns: hiddenCols,
    hyperlinks: [],
    formulasWithoutValue: [],
  });

  it("unsorted, overlapping and adjacent spans answer like a plain scan", () => {
    const spans = [
      { min: 10, max: 12 },
      { min: 2, max: 4 },
      { min: 3, max: 6 },
      { min: 7, max: 7 },
      { min: 20, max: 20 },
    ];
    const sheet = sheetWith(spans);
    for (let c = 0; c < 30; c++) expect(isColumnHidden(sheet, c)).toBe(spans.some((s) => c >= s.min && c <= s.max));
  });

  it("8 000 spans x 200 000 lookups stay fast (not spans x cells)", () => {
    const sheet = sheetWith(Array.from({ length: 8000 }, (_, i) => ({ min: i * 2, max: i * 2 })));
    const start = performance.now();
    let hidden = 0;
    for (let i = 0; i < 200_000; i++) if (isColumnHidden(sheet, i % 16_384)) hidden++;
    expect(hidden).toBeGreaterThan(0);
    expect(performance.now() - start).toBeLessThan(500);
  });
});

describe("xlsx-cells — a Strict workbook", () => {
  it("our reader reads it exactly as the transitional one", async () => {
    const readAll = async (ns: Namespaces) => {
      const { pkg, sheets } = await readWorkbookSheets(await fabricated(ns));
      const workbook = await readSharedStringsAndStyles(pkg);
      const xml = await pkg.readXml(sheets[0]!.partPath);
      return { sheets, workbook, sheet: await readWorksheetCells(pkg, sheets[0]!.partPath, xml!, workbook) };
    };
    const strict = await readAll(STRICT);
    expect(strict.sheet.cells.size).toBe(11);
    expect(strict).toEqual(await readAll(TRANSITIONAL));
  });

  it("the whole pipeline (extractXlsx) publishes its tab", async () => {
    const { doc } = await extractXlsx(await fabricated(STRICT));
    const tabs = doc.blocks[0];
    expect(tabs?.kind === "tabs" ? tabs.tabs.map((t) => t.title) : []).toEqual(["Feuille"]);
  });
});
