/** Characterisation of `readXlsxSource` (`source-reading-xlsx.ts`), stage by stage, on made-up workbooks: which sheets are read and
 *  which are declared empty, how the hidden columns and rows and the optional references shape the kept grid, how the merges and the
 *  columns that show no text place the numeric cells. Written before the function was cut into one function per stage; every
 *  expectation was seen green on the uncut function.
 *  NOT here: the number and date classification, the format that hides a value (`source-reading-xlsx.test.ts`), the merge rules
 *  themselves (`xlsx-merges-cost.test.ts`), the real fixtures (`no-word-lost.test.ts`). */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_WORKBOOK_TABS, PAGE_SIZED_LIMITS } from "../extractor/page-caps.test.support";
import { readXlsxSource } from "./source-reading-xlsx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

interface SheetSpec {
  /** Absent: the sheet element carries no name. */
  name?: string;
  state?: string;
  /** The relationship target; default `worksheets/sheet<N>.xml`. `null`: no relationship at all for this sheet. */
  target?: string | null;
  /** The worksheet part written at the target; `null`: no part. Default: the sheet has `sheetData` from `data` and the rest from `before`/`after`. */
  part?: string | null;
  /** Elements between `<worksheet>` and `<sheetData>` (`<cols>`). */
  before?: string;
  data?: string;
  /** Elements after `<sheetData>` (`<mergeCells>`). */
  after?: string;
}

const workbookOf = async (
  sheets: readonly SheetSpec[],
  extra: { noSheetsElement?: boolean } = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  const sheetEl = (s: SheetSpec, i: number): string =>
    `<sheet${s.name !== undefined ? ` name="${s.name}"` : ""} sheetId="${i + 1}" r:id="rId${i + 1}"${s.state ? ` state="${s.state}"` : ""}/>`;
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">${extra.noSheetsElement ? "" : `<sheets>${sheets.map(sheetEl).join("")}</sheets>`}</workbook>`,
  );
  const targetOf = (s: SheetSpec, i: number): string | null =>
    s.target === undefined ? `worksheets/sheet${i + 1}.xml` : s.target;
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      sheets
        .map((s, i) =>
          targetOf(s, i) === null
            ? ""
            : `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="${targetOf(s, i)}"/>`,
        )
        .join(""),
    ),
  );
  sheets.forEach((s, i) => {
    const target = targetOf(s, i);
    if (target === null || s.part === null) return;
    zip.file(
      `xl/${target}`,
      s.part ??
        `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${s.before ?? ""}<sheetData>${s.data ?? ""}</sheetData>${s.after ?? ""}</worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const text = (ref: string, word: string): string => `<c r="${ref}" t="inlineStr"><is><t>${word}</t></is></c>`;
const num = (ref: string, value = "1"): string => `<c r="${ref}"><v>${value}</v></c>`;
const row = (r: number | string, cells: string, attrs = ""): string => `<row r="${r}"${attrs}>${cells}</row>`;

const withoutStructure = (reading: object): object =>
  Object.fromEntries(Object.entries(reading).filter(([key]) => key !== "structure"));

const emptyReading = (name: string, flags: { hidden?: boolean; overCap?: boolean } = {}) => ({
  name,
  text: "",
  numericCells: [],
  hidden: flags.hidden ?? false,
  overCap: flags.overCap ?? false,
  withheld: false,
});

describe("readXlsxSource — which sheets are read", () => {
  it("a workbook with no sheets element has no sheet", async () => {
    expect(await readXlsxSource(await workbookOf([], { noSheetsElement: true }), PAGE_SIZED_LIMITS)).toEqual({
      sheets: [],
    });
  });

  it("hidden and very hidden sheets, chartsheets, unresolved relationships and missing parts are empty; the others are read, in workbook order", async () => {
    const visible = (name: string, word: string): SheetSpec => ({ name, data: row(1, text("A1", word)) });
    const bytes = await workbookOf([
      visible("first", "alpha"),
      { name: "hid", state: "hidden", data: row(1, text("A1", "secret")) },
      { name: "veryhid", state: "veryHidden", data: row(1, text("A1", "secret")) },
      { name: "chart", target: "chartsheets/sheet4.xml", part: "<chartsheet/>" },
      { name: "norel", target: null },
      { name: "nopart", part: null },
      { target: "worksheets/unnamed.xml", data: row(1, text("A1", "unnamed")) },
      { name: "shown", state: "visible", data: row(1, text("A1", "omega")) },
    ]);
    const { sheets } = await readXlsxSource(bytes, PAGE_SIZED_LIMITS);
    // The structure counts of each sheet have their own tests (the app's
    // `src/conservation/structure-conservation.sheets.test.ts`); here only which sheets are read.
    expect(sheets.map(withoutStructure)).toEqual([
      { ...emptyReading("first"), text: "alpha" },
      emptyReading("hid", { hidden: true }),
      emptyReading("veryhid", { hidden: true }),
      emptyReading("chart"),
      emptyReading("norel"),
      emptyReading("nopart"),
      { ...emptyReading(""), text: "unnamed" },
      { ...emptyReading("shown"), text: "omega" },
    ]);
  });

  // Revue 27/10/2026, tour 9: the missing part no longer takes a slot (it takes none on the page either, `extractor/xlsx.ts#open`
  // returns "part-missing" and `publishWorkbook` gives it no tab), so the sheet that used to sit just past the cap now fits.
  it("the tab cap counts the visible sheets only, the chartsheet included, the MISSING PART excluded; a hidden sheet after the cap stays hidden, not over the cap", async () => {
    const specs: SheetSpec[] = [
      { name: "h0", state: "hidden" },
      { name: "chart", target: "chartsheets/c.xml", part: "<chartsheet/>" },
      { name: "nopart", part: null },
    ];
    for (let i = specs.length - 1; i < MAX_WORKBOOK_TABS - 1; i++)
      specs.push({ name: `v${i}`, data: row(1, text("A1", `w${i}`)) });
    specs.push({ name: "last", data: row(1, text("A1", "inside")) });
    specs.push(
      { name: "h1", state: "veryHidden" },
      { name: "over1", data: row(1, text("A1", "outside")) },
      { name: "over2", data: row(1, text("A1", "outside")) },
    );
    const { sheets } = await readXlsxSource(await workbookOf(specs), PAGE_SIZED_LIMITS);
    expect(sheets.map((s) => [s.name, s.hidden, s.overCap, s.text === ""])).toEqual([
      ["h0", true, false, true],
      ["chart", false, false, true],
      ["nopart", false, false, true],
      ...specs
        .slice(3, MAX_WORKBOOK_TABS + 1)
        .map((s): [string, boolean, boolean, boolean] => [s.name!, false, false, false]),
      ["h1", true, false, true],
      ["over1", false, false, false], // the slot the missing part no longer takes
      ["over2", false, true, true],
    ]);
    // MAX_WORKBOOK_TABS sheets hold a tab slot, plus the slotless missing part, which is neither hidden nor over the cap.
    expect(sheets.filter((s) => !s.hidden && !s.overCap)).toHaveLength(MAX_WORKBOOK_TABS + 1);
  });
});

describe("readXlsxSource — the workbook part", () => {
  it("a package whose main part is not xl/workbook.xml, and that holds none, is refused", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    );
    zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/other.xml"/>`));
    zip.file("xl/other.xml", "<other/>");
    await expect(readXlsxSource(await zip.generateAsync({ type: "uint8array" }), PAGE_SIZED_LIMITS)).rejects.toThrow(
      "no workbook part",
    );
  });
});

describe("readXlsxSource — hidden columns and rows, references", () => {
  const words = (cells: string[]): string => row(1, cells.map((c, i) => text(`${"ABCDEF"[i]}1`, c)).join(""));

  it("columns hidden by `1` or `true`, in every cols element, by their min/max span; a visible span and an unhidden one stay", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        before: `<cols><col min="2" max="2" hidden="1"/><col min="3" max="3" hidden="true"/><col min="4" max="4" hidden="0"/></cols><cols><col min="6" max="7" hidden="1"/></cols>`,
        data: words(["a", "b", "c", "d", "e", "f"]),
      },
    ]);
    expect((await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets[0]!.text).toBe("a d e");
  });

  it("a hidden span starting far below the first column costs nothing (clamped to column 1)", async () => {
    const bytes = await workbookOf([
      { name: "s", before: `<cols><col min="-2000000000" max="2" hidden="1"/></cols>`, data: words(["a", "b", "c"]) },
    ]);
    expect((await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets[0]!.text).toBe("c");
  }, 4000);

  it("a hidden col with no max hides nothing; one with no min starts at the first column", async () => {
    const a = await workbookOf([
      { name: "s", before: `<cols><col hidden="1"/><col min="2" hidden="1"/></cols>`, data: words(["a", "b", "c"]) },
    ]);
    expect((await readXlsxSource(a, PAGE_SIZED_LIMITS)).sheets[0]!.text).toBe("a b c");
    const b = await workbookOf([
      { name: "s", before: `<cols><col max="2" hidden="1"/></cols>`, data: words(["a", "b", "c"]) },
    ]);
    expect((await readXlsxSource(b, PAGE_SIZED_LIMITS)).sheets[0]!.text).toBe("c");
  });

  it("a row hidden by `1` or `true` is not read and keeps no line: the next row is the first of the kept grid", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        data:
          row(1, text("A1", "x") + num("B1"), ' hidden="1"') +
          row(2, text("A2", "y"), ' hidden="true"') +
          row(3, num("A3", "7") + text("B3", "kept"), ' hidden="0"'),
      },
    ]);
    const [sheet] = (await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets;
    expect(sheet!.text).toBe("kept");
    expect(sheet!.numericCells).toEqual([{ ref: "s!A3", row: 0, col: 0, cause: "number" }]);
  });

  it("a cell with no value keeps neither its row nor its column", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        data:
          row(
            1,
            `<c r="A1"/><c r="B1" t="inlineStr"><is><t></t></is></c><c r="C1" t="s"><v>99</v></c><c r="D1"><v></v></c>`,
          ) + row(2, num("E2", "4")),
      },
    ]);
    const [sheet] = (await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets;
    expect(sheet!.numericCells).toEqual([{ ref: "s!E2", row: 0, col: 0, cause: "number" }]);
  });

  it("a cell in a hidden column keeps neither its row nor its column: the number below sits on the first kept row", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        before: `<cols><col min="1" max="1" hidden="1"/></cols>`,
        data: row(1, text("A1", "hid")) + row(2, num("B2", "5")),
      },
    ]);
    const [sheet] = (await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets;
    expect(sheet!.numericCells).toEqual([{ ref: "s!B2", row: 0, col: 0, cause: "number" }]);
    expect(sheet!.text).toBe("");
  });

  it("a row reference that is absent, zero, fractional or not a number takes the row after its predecessor", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        data: [
          row(5, num("A5")),
          `<row r="0">${num("A0")}</row>`,
          `<row>${num("A1")}</row>`,
          `<row r="2.5">${num("A2")}</row>`,
          row("x", num("A3")),
        ].join(""),
      },
    ]);
    const [sheet] = (await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets;
    expect(sheet!.numericCells.map((c) => [c.ref, c.row])).toEqual([
      ["s!A5", 0],
      ["s!A6", 1],
      ["s!A7", 2],
      ["s!A8", 3],
      ["s!A9", 4],
    ]);
  });

  it("the words come in source order, one space apart, empty cells left out", async () => {
    const bytes = await workbookOf([
      {
        name: "s",
        data:
          row(1, text("A1", "one") + `<c r="B1" t="inlineStr"><is><t></t></is></c>` + text("C1", "two")) +
          row(2, text("A2", "three")),
      },
    ]);
    expect((await readXlsxSource(bytes, PAGE_SIZED_LIMITS)).sheets[0]!.text).toBe("one two three");
  });
});

describe("readXlsxSource — merges and the columns that show no text place the numeric cells", () => {
  const sheetWith = async (data: string, after: string) =>
    (await readXlsxSource(await workbookOf([{ name: "s", data, after }]), PAGE_SIZED_LIMITS)).sheets[0]!;
  const merges = (...refs: string[]): string =>
    `<mergeCells>${refs.map((r) => `<mergeCell ref="${r}"/>`).join("")}</mergeCells>`;

  it("the text of a covered cell is not expected; the anchor's and the others' are", async () => {
    const sheet = await sheetWith(
      row(1, text("A1", "anchor") + text("B1", "covered") + text("C1", "free")),
      merges("A1:B1"),
    );
    expect(sheet.text).toBe("anchor free");
  });

  it("a merge written from its bottom row to its top row covers the same cells", async () => {
    const sheet = await sheetWith(row(1, text("A1", "top")) + row(2, text("A2", "below")), merges("A2:A1"));
    expect(sheet.text).toBe("top");
  });

  it("the rows and columns are ranked in numeric order, not in the order the file lists them", async () => {
    const sheet = await sheetWith(row(5, num("C5") + num("A5")) + row(3, num("B3")), "");
    expect(sheet.numericCells.map((c) => [c.ref, c.row, c.col])).toEqual([
      ["s!C5", 1, 2],
      ["s!A5", 1, 0],
      ["s!B3", 0, 1],
    ]);
  });

  it("a merge written backwards covers the same cells; a single-cell reference covers none; an unreadable one is ignored", async () => {
    const sheet = await sheetWith(
      row(1, text("A1", "anchor") + text("B1", "covered") + text("C1", "free")),
      merges("B1:A1", "C1", "??:A1", "A1:??"),
    );
    expect(sheet.text).toBe("anchor free");
  });

  it("a covered number is placed where its merge's first cell is, and its own column does not count", async () => {
    const sheet = await sheetWith(row(1, text("A1", "anchor") + num("B1", "5") + text("C1", "x")), merges("A1:B1"));
    expect(sheet.numericCells).toEqual([{ ref: "s!B1", row: 0, col: 0, cause: "number" }]);
  });

  it("a column where no uncovered cell shows text is dropped: the numbers after it shift left by one", async () => {
    const blank = `<c r="A1" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>`;
    const sheet = await sheetWith(row(1, blank + text("B1", "x") + num("C1", "2")), "");
    expect(sheet.numericCells).toEqual([{ ref: "s!C1", row: 0, col: 1, cause: "number" }]);
  });

  it("a number shows: it keeps its column among the text columns", async () => {
    const blank = `<c r="A1" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>`;
    const sheet = await sheetWith(row(1, blank + text("B1", "x") + num("C1", "2")) + row(2, num("A2", "3")), "");
    expect(sheet.numericCells).toEqual([
      { ref: "s!C1", row: 0, col: 2, cause: "number" },
      { ref: "s!A2", row: 1, col: 0, cause: "number" },
    ]);
  });

  it("a merge's first cell in a dropped column leaves the covered number past the table's extent", async () => {
    const blank = `<c r="A1" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>`;
    const sheet = await sheetWith(row(1, blank + num("B1", "5") + text("C1", "x")), merges("A1:B1"));
    expect(sheet.numericCells).toEqual([{ ref: "s!B1", row: 0, col: 1, cause: "number" }]);
  });

  it("a merge over several rows and columns: every covered number is placed at its anchor", async () => {
    const sheet = await sheetWith(row(1, num("A1", "1") + num("B1", "2")) + row(2, num("A2", "3")), merges("A1:B2"));
    expect(sheet.numericCells.map((c) => [c.ref, c.row, c.col])).toEqual([
      ["s!A1", 0, 0],
      ["s!B1", 0, 0],
      ["s!A2", 0, 0],
    ]);
  });
});
