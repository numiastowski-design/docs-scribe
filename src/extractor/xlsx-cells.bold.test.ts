// A cell is bold when its style's font says so (ECMA-376 Part 1 §18.8.2 CT_Font/b, CT_BooleanProperty: val defaults to true),
// or, for a rich string, when its runs do (a run's own rPr replaces the cell font).
//   npx vitest run --project=main src/extractor/xlsx-cells.bold.test.ts
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { readWorkbookSheets } from "./xlsx-parts";
import { readWorksheetCells, readSharedStringsAndStyles } from "./xlsx-cells";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** fonts: the inner XML of each <font>; xfs: the fontId attribute text of each cellXfs <xf> (empty string = no fontId). */
const build = async (fonts: string[], xfs: string[], sst: string, rows: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  const rels = (items: string): string =>
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  zip.file(
    "[Content_Types].xml",
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
  );
  zip.file("_rels/.rels", rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/_rels/workbook.xml.rels",
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`,
    ),
  );
  zip.file(
    "xl/workbook.xml",
    `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/styles.xml",
    `${DECL}<styleSheet xmlns="${MAIN}"><fonts count="${fonts.length}">${fonts.map((f) => `<font>${f}</font>`).join("")}</fonts>` +
      `<cellXfs count="${xfs.length}">${xfs.map((f) => `<xf numFmtId="0"${f}/>`).join("")}</cellXfs></styleSheet>`,
  );
  zip.file("xl/sharedStrings.xml", `${DECL}<sst xmlns="${MAIN}">${sst}</sst>`);
  zip.file("xl/worksheets/sheet1.xml", `${DECL}<worksheet xmlns="${MAIN}"><sheetData>${rows}</sheetData></worksheet>`);
  return zip.generateAsync({ type: "uint8array" });
};

const boldOf = async (bytes: Uint8Array): Promise<Record<string, boolean | undefined>> => {
  const { pkg, sheets } = await readWorkbookSheets(bytes);
  const workbook = await readSharedStringsAndStyles(pkg);
  const sheet = await readWorksheetCells(pkg, sheets[0]!.partPath, (await pkg.readXml(sheets[0]!.partPath))!, workbook);
  return Object.fromEntries([...sheet.cells].map(([ref, cell]) => [ref, cell.isBold]));
};

const FONTS = ['<sz val="11"/>', "<b/>", '<b val="0"/>', '<b val="false"/>', '<b val="1"/>', '<b val="true"/>'];
const XFS = ["", ' fontId="1"', ' fontId="2"', ' fontId="3"', ' fontId="4"', ' fontId="5"', ' fontId="99"'];
const cell = (ref: string, s: number): string => `<c r="${ref}" s="${s}" t="inlineStr"><is><t>x</t></is></c>`;

describe("xlsx-cells — bold from the cell style's font", () => {
  // Review of 14/10/2026: Number("") is 0, so an empty fontId used to read font 0.
  it("an empty fontId is no font at all, even when font 0 is bold", async () => {
    const rows = `<row r="1">${cell("A1", 0)}${cell("B1", 1)}</row>`;
    expect(await boldOf(await build(["<b/>"], [' fontId=""', ' fontId="0"'], "", rows))).toEqual({
      A1: false,
      B1: true,
    });
  });

  it("reads <b/>, <b val=1> and <b val=true> as bold; <b val=0>, <b val=false>, no <b>, no fontId and a fontId out of range as plain", async () => {
    const rows = `<row r="1">${[0, 1, 2, 3, 4, 5, 6].map((s) => cell(`${"ABCDEFG"[s]}1`, s)).join("")}</row>`;
    expect(await boldOf(await build(FONTS, XFS, "", rows))).toEqual({
      A1: false,
      B1: true,
      C1: false,
      D1: false,
      E1: true,
      F1: true,
      G1: false,
    });
  });

  it("a style index the workbook does not define is plain", async () => {
    expect(await boldOf(await build(FONTS, XFS, "", `<row r="1">${cell("A1", 40)}</row>`))).toEqual({ A1: false });
  });
});

describe("xlsx-cells — bold from a rich string's runs", () => {
  const sst =
    "<si><r><rPr><b/></rPr><t>all</t></r><r><rPr><b/></rPr><t> bold</t></r></si>" + // 0: every run bold
    '<si><r><rPr><b/></rPr><t>one</t></r><r><rPr><sz val="11"/></rPr><t> plain</t></r></si>' + // 1: mixed
    "<si><r><t>inherits</t></r></si>" + // 2: a run without rPr takes the cell font
    '<si><r><rPr><b val="0"/></rPr><t>own plain</t></r></si>'; // 3: the run's own plain wins over a bold cell
  const shared = (ref: string, idx: number, s: number): string => `<c r="${ref}" s="${s}" t="s"><v>${idx}</v></c>`;

  it("a string whose runs are all bold is bold even under a plain cell font", async () => {
    expect((await boldOf(await build(FONTS, XFS, sst, `<row r="1">${shared("A1", 0, 0)}</row>`))).A1).toBe(true);
  });

  it("a run's own properties win over the cell font, both ways", async () => {
    const rows = `<row r="1">${shared("A1", 3, 1)}${shared("B1", 2, 1)}${shared("C1", 2, 0)}</row>`;
    expect(await boldOf(await build(FONTS, XFS, sst, rows))).toEqual({ A1: false, B1: true, C1: false });
  });

  it("runs of mixed weight give a plain cell (one inline per cell: the limit, not a guess)", async () => {
    expect((await boldOf(await build(FONTS, XFS, sst, `<row r="1">${shared("A1", 1, 1)}</row>`))).A1).toBe(false);
  });

  it("an inline string with bold runs is read the same way", async () => {
    const rows = '<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r></is></c></row>';
    expect((await boldOf(await build(FONTS, XFS, "", rows))).A1).toBe(true);
  });
});
