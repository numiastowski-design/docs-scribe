/** Behaviour tests for `readXlsxConservationSource`
 *  (MIKSUDOC « Spécification technique » T11): a real xlsx yields a
 *  usable declared-format source, a read that cannot be trusted becomes a
 *  named failure — never a thrown error, never a reason not to publish
 *  (21/09/2026) — and the declared removals (tab cap, size-cap
 *  withhold, hidden sheet) never surface as missing words.
 *  `readXlsxSource` itself is exercised exhaustively, against every real
 *  fixture, by the app's `conservation/no-word-lost.test.ts`. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  MAX_SHEET_COLS,
  MAX_SHEET_ROWS,
  MAX_WORKBOOK_TABS,
  PAGE_SIZED_LIMITS,
} from "../extractor/page-caps.test.support";
import { extractXlsx } from "../extractor/xlsx";
import { isDeclaredFormatSource, isReadFailure } from "./conservation-source";
import { readXlsxConservationSource, readXlsxSource } from "./source-reading-xlsx";

const root = join(dirname(fileURLToPath(import.meta.url)), "../..");
const fixture = (relative: string): Uint8Array => new Uint8Array(readFileSync(join(root, relative)));

// ---------------------------------------------------------------------------
// Made-up workbooks — a minimal OPC package, one worksheet part per sheet.

const workbookOf = async (
  sheets: readonly { name: string; state?: "hidden" | "veryHidden"; cellsXml: string }[],
  stylesXml?: string,
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const RELS = (inner: string): string =>
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      `xmlns:r="${REL}"><sheets>` +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"${s.state ? ` state="${s.state}"` : ""}/>`,
        )
        .join("") +
      "</sheets></workbook>",
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  if (stylesXml !== undefined) zip.file("xl/styles.xml", stylesXml);
  sheets.forEach((s, i) => {
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData>${s.cellsXml}</sheetData></worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const colLetter = (n: number): string => {
  let out = "";
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) out = String.fromCharCode(65 + ((x - 1) % 26)) + out;
  return out;
};

/** ONE row, ONE inline-string cell (`t="inlineStr"`: no shared-strings part
 *  needed for these made-up fixtures). */
const rowOf = (rowNum: number, col: number, word: string): string =>
  `<row r="${rowNum}"><c r="${colLetter(col)}${rowNum}" t="inlineStr"><is><t>${word}</t></is></c></row>`;

const allWords = (source: Awaited<ReturnType<typeof readXlsxConservationSource>>): string[] => {
  if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
  return source.units.flatMap((u) => u.tokens.map((t) => t.text));
};

describe("readXlsxConservationSource — real xlsx fixtures", () => {
  it("a real workbook: one unit per visible sheet, words present", async () => {
    const source = await readXlsxConservationSource(
      fixture("fixtures/xlsx-reel/16-Collectivites-Fiscalite-Locale.xlsx"),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source, got a read failure");
    expect(source.units.length).toBeGreaterThan(0);
    expect(source.units.flatMap((u) => u.tokens).length).toBeGreaterThan(20);
  });

  it("🔴 GUARD BITES: bytes that are not a zip at all — a named failure, never a thrown error", async () => {
    const source = await readXlsxConservationSource(new TextEncoder().encode("not an xlsx"), PAGE_SIZED_LIMITS);
    expect(isReadFailure(source)).toBe(true);
  });

  it("🔴 GUARD BITES: a zip with no xl/workbook.xml — a named failure", async () => {
    const zip = new JSZip();
    zip.file("xl/worksheets/sheet1.xml", "<worksheet/>");
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const source = await readXlsxConservationSource(bytes, PAGE_SIZED_LIMITS);
    expect(isReadFailure(source)).toBe(true);
  });
});

describe("readXlsxConservationSource — the declared removals never surface as missing", () => {
  it("a sheet past the tab cap: its unit is excluded, never a source of tokens", async () => {
    const sheets = Array.from({ length: MAX_WORKBOOK_TABS + 2 }, (_, i) => ({
      name: `Sheet${i + 1}`,
      cellsXml: rowOf(1, 1, `word${i + 1}unique`),
    }));
    const source = await readXlsxConservationSource(await workbookOf(sheets), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(MAX_WORKBOOK_TABS + 2);
    expect([...source.excludedUnits].sort((a, b) => a - b)).toEqual([MAX_WORKBOOK_TABS, MAX_WORKBOOK_TABS + 1]);
    expect(source.units[MAX_WORKBOOK_TABS]!.tokens).toEqual([]);
    expect(allWords(source)).toContain("word1unique");
    expect(allWords(source)).not.toContain(`word${MAX_WORKBOOK_TABS + 1}unique`);
  });

  it("a sheet over MAX_SHEET_ROWS: withheld whole, its unit excluded and empty", async () => {
    const rows = Array.from({ length: MAX_SHEET_ROWS + 2 }, (_, i) => rowOf(i + 1, 1, `r${i}unique`)).join("");
    const sheets = [
      { name: "Big", cellsXml: rows },
      { name: "Small", cellsXml: rowOf(1, 1, "next sheet unique") },
    ];
    const source = await readXlsxConservationSource(await workbookOf(sheets), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(2);
    expect([...source.excludedUnits]).toEqual([0]);
    expect(source.units[0]!.tokens).toEqual([]);
    expect(allWords(source)).not.toContain("r0unique");
    expect(allWords(source)).toContain("unique");
  });

  it("a sheet over MAX_SHEET_COLS: withheld whole, its unit excluded and empty", async () => {
    const cells = Array.from(
      { length: MAX_SHEET_COLS + 2 },
      (_, i) => `<c r="${colLetter(i + 1)}1" t="inlineStr"><is><t>c${i}unique</t></is></c>`,
    ).join("");
    const sheets = [
      { name: "Wide", cellsXml: `<row r="1">${cells}</row>` },
      { name: "Small", cellsXml: rowOf(1, 1, "next sheet unique") },
    ];
    const source = await readXlsxConservationSource(await workbookOf(sheets), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).not.toContain("c0unique");
    expect(allWords(source)).toContain("unique");
  });

  it("a hidden sheet: a unit of its own, excluded, never counted toward the cap", async () => {
    const sheets = [
      { name: "Shown1", cellsXml: rowOf(1, 1, "shown one unique") },
      { name: "Hidden", state: "hidden" as const, cellsXml: rowOf(1, 1, "hidden words unique") },
      { name: "Shown2", cellsXml: rowOf(1, 1, "shown two unique") },
    ];
    const source = await readXlsxConservationSource(await workbookOf(sheets), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(3);
    expect([...source.excludedUnits]).toEqual([1]);
    expect(source.units[1]!.label).toBe("Hidden");
    expect(source.units[1]!.tokens).toEqual([]);
    expect(allWords(source)).not.toContain("hidden");
    expect(allWords(source)).toContain("one");
    expect(allWords(source)).toContain("two");
  });

  it("a hidden column: its word never read, the rest of the row still is", async () => {
    const sheetXml =
      `<row r="1"><c r="A1" t="inlineStr"><is><t>kept</t></is></c>` +
      `<c r="B1" t="inlineStr"><is><t>masked</t></is></c></row>`;
    // `<cols>` sits before `<sheetData>` in a real part; the reader looks
    // for it among the worksheet's own children, order aside.
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>',
    );
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS = (inner: string): string =>
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
    zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
    zip.file(
      "xl/workbook.xml",
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        `xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      RELS(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`),
    );
    zip.file(
      "xl/worksheets/sheet1.xml",
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<cols><col min="2" max="2" hidden="1"/></cols><sheetData>${sheetXml}</sheetData></worksheet>`,
    );
    const source = await readXlsxConservationSource(await zip.generateAsync({ type: "uint8array" }), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    const words = allWords(source);
    expect(words).toContain("kept");
    expect(words).not.toContain("masked");
  });
});

// ---------------------------------------------------------------------------
// Numbers and dates (§1, lot E addendum, 29/09/2026): checked by
// POSITION on the published page (`pipeline/xlsx-page-cells.ts`), never as
// text here — these tests only pin what THIS reading itself produces: the
// `numericCells` list (position, ref, cause), and the scope limits (a
// formula's empty `<v>`, a boolean, an error cell never enter it).

describe("readXlsxConservationSource — numeric and date source cells", () => {
  it('a plain number cell (no style at all): cause "number", ref quoted only when the sheet name needs it', async () => {
    const cellsXml = '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1"><v>42</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "Sheet1", cellsXml }]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "Sheet1!B1", row: 0, col: 1, cause: "number" }]);
  });

  it("a sheet name that is not a bare identifier is Excel-quoted in the ref", async () => {
    const cellsXml = '<row r="1"><c r="A1"><v>1</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "Budget 2026", cellsXml }]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "'Budget 2026'!A1", row: 0, col: 0, cause: "number" }]);
  });

  it('a t="d" cell is always "date", style aside (ECMA-376 §18.3.1.4)', async () => {
    const cellsXml = '<row r="1"><c r="A1" t="d"><v>2026-01-01T00:00:00Z</v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "date" }]);
  });

  it('a built-in date numFmtId (14, ECMA-376 §18.8.30): "date", read through the style — never the literal mask', async () => {
    const styles =
      '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<cellXfs count="1"><xf numFmtId="14"/></cellXfs></styleSheet>';
    const cellsXml = '<row r="1"><c r="A1" s="0"><v>45658</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styles),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "date" }]);
  });

  it('a CUSTOM numFmtId whose code carries a date/time letter run: "date", independently of the writer', async () => {
    const styles =
      '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>' +
      '<cellXfs count="1"><xf numFmtId="164"/></cellXfs></styleSheet>';
    const cellsXml = '<row r="1"><c r="A1" s="0"><v>45658</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styles),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "date" }]);
  });

  it('a CUSTOM numFmtId with no date/time letter (a currency mask): "number"', async () => {
    const styles =
      '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;€&quot;#,##0.00"/></numFmts>' +
      '<cellXfs count="1"><xf numFmtId="164"/></cellXfs></styleSheet>';
    const cellsXml = '<row r="1"><c r="A1" s="0"><v>19.9</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styles),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "number" }]);
  });

  it("🔴 GUARD BITES: a formula's EMPTY <v></v> (no cached value, decision f) is never a numeric source cell", async () => {
    const cellsXml = '<row r="1"><c r="A1"><f>1+1</f><v></v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toBeUndefined();
  });

  it("a formula WITH a cached value is an ordinary numeric source cell", async () => {
    const cellsXml = '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "number" }]);
  });

  it('a boolean cell is never a numeric source cell (out of scope: booleans, not "types n, d")', async () => {
    const cellsXml = '<row r="1"><c r="A1" t="b"><v>1</v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toBeUndefined();
  });

  it("an error cell is never a numeric source cell (it stays a TEXT token, unchanged)", async () => {
    const cellsXml = '<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toBeUndefined();
    expect(allWords(source)).toEqual(expect.arrayContaining(["div", "0"]));
  });

  it("the OUTPUT position is the cell's rank among KEPT rows/columns, not its raw A1 address", async () => {
    // Row 1 and row 5 are the only ones with anything; column A and column B
    // the only ones — the kept grid compacts to a 2×2 table, output indices
    // 0/1 each, whatever the raw row/column numbers were.
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>row5</t></is></c><c r="B5"><v>7</v></c></row>';
    const source = await readXlsxConservationSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!B5", row: 1, col: 1, cause: "number" }]);
  });

  it("a withheld sheet (over the size cap) carries no numeric cell either — same as its words", async () => {
    const rows = Array.from(
      { length: MAX_SHEET_ROWS + 2 },
      (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`,
    ).join("");
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "Big", cellsXml: rows }]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Review findings A, B, 3, 4 (30/09/2026).

const withPart = async (bytes: Uint8Array, path: string, content: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(bytes);
  zip.file(path, content);
  return zip.generateAsync({ type: "uint8array" });
};

const SS_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

describe("readXlsxSource — hardened reads of the hard-coded parts (finding A)", () => {
  it("🔴 a DTD in xl/sharedStrings.xml is refused: an abstention, not a reading", async () => {
    const base = await workbookOf([{ name: "S", cellsXml: rowOf(1, 1, "alpha") }]);
    const bytes = await withPart(
      base,
      "xl/sharedStrings.xml",
      `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><sst ${SS_NS}/>`,
    );
    expect(isReadFailure(await readXlsxConservationSource(bytes, PAGE_SIZED_LIMITS))).toBe(true);
  });

  it("a styles part inflating past PART_MAX_BYTES abstains", async () => {
    const base = await workbookOf([{ name: "S", cellsXml: rowOf(1, 1, "alpha") }]);
    const zip = await JSZip.loadAsync(base);
    zip.file("xl/styles.xml", `<styleSheet ${SS_NS}><!--${" ".repeat(65_000_000)}--></styleSheet>`, {
      compression: "DEFLATE",
    });
    const bomb = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    expect(bomb.byteLength).toBeLessThan(1_000_000);
    expect(isReadFailure(await readXlsxConservationSource(bomb, PAGE_SIZED_LIMITS))).toBe(true);
  }, 60_000);
});

describe("readXlsxSource — hidden column spans (finding B)", () => {
  it("🔴 a hidden span up to col@max=2000000000 costs nothing: the visible words are still read", async () => {
    const base = await workbookOf([
      {
        name: "S",
        cellsXml:
          '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c><c r="C1" t="inlineStr"><is><t>beta</t></is></c></row>',
      },
    ]);
    const zip = await JSZip.loadAsync(base);
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file(
      "xl/worksheets/sheet1.xml",
      sheet.replace("<sheetData>", '<cols><col min="3" max="2000000000" hidden="1"/></cols><sheetData>'),
    );
    const source = await readXlsxConservationSource(await zip.generateAsync({ type: "uint8array" }), PAGE_SIZED_LIMITS);
    expect(allWords(source)).toEqual(["alpha"]);
  });
});

describe("readXlsxSource — optional row/cell references (finding 3)", () => {
  it("🔴 rows and cells without `r` take the next position after their predecessor (§18.3.1.73, §18.3.1.4)", async () => {
    const cellsXml = "<row><c><v>1</v></c><c><v>2</v></c></row>" + "<row><c><v>3</v></c><c><v>4</v></c></row>";
    const { sheets } = await readXlsxSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    expect(sheets[0]!.numericCells.map((c) => [c.row, c.col])).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
  });
});

describe("readXlsxSource — string cells' text (finding 4)", () => {
  it("🔴 a shared string's _xHHHH_ escape is unescaped (§22.9.2.19) and its phonetic run (rPh) is not text", async () => {
    const base = await workbookOf([{ name: "S", cellsXml: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' }]);
    const bytes = await withPart(
      base,
      "xl/sharedStrings.xml",
      `<?xml version="1.0"?><sst ${SS_NS}><si><t>caf_x00E9_</t><rPh sb="0" eb="4"><t>phonetik</t></rPh></si></sst>`,
    );
    const { sheets } = await readXlsxSource(bytes, PAGE_SIZED_LIMITS);
    expect(sheets[0]!.text).toBe("café");
  });

  it("an inline string's rich runs are joined, its phonetic run left out", async () => {
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><r><t>foo</t></r><r><t>bar</t></r><rPh sb="0" eb="1"><t>zzz</t></rPh></is></c></row>';
    const { sheets } = await readXlsxSource(await workbookOf([{ name: "S", cellsXml }]), PAGE_SIZED_LIMITS);
    expect(sheets[0]!.text).toBe("foobar");
  });
});

// A number the format HIDES (`;;;`, or an empty section for that value) is
// published as nothing (`extractor/xlsx-number-format.ts#formatNumberCell`
// returns ""), and the extractor's kept grid drops a cell that shows nothing:
// the reading must not expect it — no false « missing number » — nor keep its
// row/column, or every later position would shift.
describe("readXlsxConservationSource — a number hidden by its format is not expected on the page", () => {
  const styleOf = (code: string): string =>
    '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<numFmts count="1"><numFmt numFmtId="164" formatCode="${code}"/></numFmts>` +
    '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>';

  it("🔴 `;;;`: the hidden cell is no numeric source cell, and its column no longer shifts the next one", async () => {
    // B1 is hidden by `;;;`: the published table has A (label) then C (9) as columns 0 and 1.
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1" s="1"><v>7</v></c><c r="C1"><v>9</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf(";;;")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!C1", row: 0, col: 1, cause: "number" }]);
  });

  it("🔴 `0;;`: a positive shows, zero and a negative are hidden (their own section is empty)", async () => {
    const cellsXml =
      '<row r="1"><c r="A1" s="1"><v>5</v></c></row><row r="2"><c r="A2" s="1"><v>0</v></c><c r="B2" t="inlineStr"><is><t>x</t></is></c></row>' +
      '<row r="3"><c r="A3" s="1"><v>-3</v></c><c r="B3" t="inlineStr"><is><t>y</t></is></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf("0;;")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "number" }]);
  });

  it("a visible format keeps its number expected (the guard is not a blanket skip)", async () => {
    const cellsXml = '<row r="1"><c r="A1" s="1"><v>0</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf("0;-0;0")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "number" }]);
  });

  it("🔴 a TEXT cell under an empty 4th (text) section `0;0;0;`: its word is never expected, its column leaves the kept grid", async () => {
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>secretword</t></is></c><c r="C1"><v>9</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf("0;0;0;")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(allWords(source)).toEqual(["label"]);
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!C1", row: 0, col: 1, cause: "number" }]);
  });

  it('🔴 a DATE under `;;;` (numeric date style, and t="d") is hidden too', async () => {
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1" s="1"><v>45658</v></c><c r="C1" s="1" t="d"><v>2026-01-01T00:00:00Z</v></c><c r="D1"><v>9</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf(";;;")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!D1", row: 0, col: 1, cause: "number" }]);
  });

  it("🔴 a colour tag is no condition: zero under `[Blue]0;[Red]-0;;@` is hidden, as the page publishes it", async () => {
    // B1 is 0, its zero section is empty: the published table has A (label) then C (9), and no cell for B.
    const cellsXml =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1" s="1"><v>0</v></c><c r="C1"><v>9</v></c></row>';
    // The page finds the styles through the workbook's relationship (this file's `workbookOf` writes none: the control reads the part by its path).
    const zip = await JSZip.loadAsync(await workbookOf([{ name: "S", cellsXml }], styleOf("[Blue]0;[Red]-0;;@")));
    const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
    zip.file(
      "xl/_rels/workbook.xml.rels",
      rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
      ),
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    const source = await readXlsxConservationSource(bytes, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!C1", row: 0, col: 1, cause: "number" }]);
    // The page, read from the same bytes: two cells, the zero shows nothing.
    expect(JSON.stringify((await extractXlsx(bytes, "en")).doc.blocks)).not.toMatch(/"text":"0"/);
  });

  it("a condition section (`[>100]0;0`) is not decided: the number stays expected (abstain, never a guess)", async () => {
    const cellsXml = '<row r="1"><c r="A1" s="1"><v>5</v></c></row>';
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "S", cellsXml }], styleOf("[&gt;100]0;")),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.numericCells).toEqual([{ unit: 0, ref: "S!A1", row: 0, col: 0, cause: "number" }]);
  });
});

/** Audit 07/10, P4-4.2. The extractor resolves an OPC relationship target with `ooxml/package.ts#resolveTarget` (percent-decoding, `..`
 *  segments); this independent reading did `xl/${target}` verbatim, so a workbook the extractor reads fine had its sheet part `not found` here:
 *  an empty sheet that is not excluded, no word read, a bias toward « ok ». */
describe("a worksheet target is resolved like the extractor resolves it", () => {
  const withTarget = async (target: string, partPath: string): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(
      await workbookOf([
        { name: "S", cellsXml: '<row r="1"><c r="A1" t="inlineStr"><is><t>wordfromsheet</t></is></c></row>' },
      ]),
    );
    const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
    zip.file("xl/_rels/workbook.xml.rels", rels.replace("worksheets/sheet1.xml", target));
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.remove("xl/worksheets/sheet1.xml");
    zip.file(partPath, xml);
    return zip.generateAsync({ type: "uint8array" });
  };

  it("a target with a parent segment (../xl/worksheets/sheet1.xml) finds its part", async () => {
    const { sheets } = await readXlsxSource(
      await withTarget("../xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),
      PAGE_SIZED_LIMITS,
    );
    expect(sheets[0]!.text).toBe("wordfromsheet"); // failure value: "" (part not found)
  });

  it("a percent-encoded target (my%20sheet.xml) finds its part", async () => {
    const { sheets } = await readXlsxSource(
      await withTarget("worksheets/my%20sheet.xml", "xl/worksheets/my sheet.xml"),
      PAGE_SIZED_LIMITS,
    );
    expect(sheets[0]!.text).toBe("wordfromsheet");
  });
});

/** Audit 07/10, P4-4.2. The extractor (`sheet-workbook.ts`) and this independent reading each count the visible sheets against
 *  `MAX_WORKBOOK_TABS` and the kept grid against `MAX_SHEET_ROWS`/`MAX_SHEET_COLS`, by design not sharing code. Nothing compared their
 *  verdicts on one workbook: here the sheets the extractor PUBLISHES are exactly the sheets this reading does not exclude. A divergence would
 *  be a false « missing word » (reading expects what the page omits) or a lost word read as ok. */
describe("the extractor and the independent reading exclude the same sheets", () => {
  it("hidden sheet, sheet over the row cap, sheet over the column cap, and a visible sheet past the tab cap", async () => {
    const bigRows = Array.from({ length: MAX_SHEET_ROWS + 2 }, (_, i) => rowOf(i + 1, 1, `r${i}`)).join("");
    const wideCells = Array.from(
      { length: MAX_SHEET_COLS + 2 },
      (_, i) => `<c r="${colLetter(i + 1)}1" t="inlineStr"><is><t>c${i}</t></is></c>`,
    ).join("");
    const sheets: { name: string; state?: "hidden"; cellsXml: string }[] = [
      { name: "Hid", state: "hidden", cellsXml: rowOf(1, 1, "hiddenword") },
      { name: "Big", cellsXml: bigRows },
      { name: "Wide", cellsXml: `<row r="1">${wideCells}</row>` },
      ...Array.from({ length: MAX_WORKBOOK_TABS + 1 }, (_, i) => ({
        name: `V${i + 1}`,
        cellsXml: rowOf(1, 1, `visible${i + 1}`),
      })),
    ];
    const bytes = await workbookOf(sheets);
    const { doc } = await extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS);
    const tabs = doc.blocks[0]?.kind === "tabs" ? doc.blocks[0].tabs.map((t) => t.title) : [];
    const { sheets: read } = await readXlsxSource(bytes, PAGE_SIZED_LIMITS);
    const kept = read.filter((s) => !s.hidden && !s.overCap && !s.withheld).map((s) => s.name);
    // Failure value: a sheet published by one and excluded by the other.
    expect(tabs).toEqual(kept);
    // A withheld sheet still takes one of the slots (A7, 29/09/2026), a hidden one none: the fixture builds
    // MAX_WORKBOOK_TABS + 1 visible sheets, so cap - 2 withheld survive whatever the cap's value (20 since 27/10/2026, was 6).
    expect(tabs).toEqual(Array.from({ length: MAX_WORKBOOK_TABS - 2 }, (_, i) => `V${i + 1}`));
  });
});
