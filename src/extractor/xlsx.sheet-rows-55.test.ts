import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_SHEET_ROWS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractXlsx } from "./xlsx";

// 08/10/2026: ONE row cap for worksheet tabs (MAX_SHEET_ROWS), 55 instead of 100.
// Run: npx vitest run --project=main src/extractor/xlsx.sheet-rows-55.test.ts

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

/** `n` filled rows (one inline-string cell each, word `${prefix}${i}x`), then `empty` rows holding no cell. */
const rowsOf = (n: number, prefix: string, empty = 0): string =>
  Array.from(
    { length: n },
    (_, i) =>
      `<row r="${i + 1}"><c r="${colLetter(1)}${i + 1}" t="inlineStr"><is><t>${prefix}${i + 1}x</t></is></c></row>`,
  ).join("") + Array.from({ length: empty }, (_, i) => `<row r="${n + i + 1}"/>`).join("");

type Run = { titles: string[]; json: string; notices: readonly { kind: string }[] };
const run = async (sheets: { name: string; cellsXml: string }[]): Promise<Run> => {
  const { doc, notices } = await extractXlsx(await workbookOf(sheets), undefined, PAGE_SIZED_LIMITS);
  const tabs = doc.blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  return { titles: tabs.tabs.map((t) => t.title), json: JSON.stringify(doc), notices };
};
const small = { name: "Small", cellsXml: rowsOf(2, "s") };

describe("xlsx — a tab of more than MAX_SHEET_ROWS kept rows is withheld (sheet-withheld), a tab AT the cap is published", () => {
  it("a tab of MAX_SHEET_ROWS filled rows is published whole, no remark", async () => {
    const r = await run([{ name: "Long", cellsXml: rowsOf(MAX_SHEET_ROWS, "l") }, small]);
    expect(r.titles).toEqual(["Long", "Small"]);
    expect(r.json).toContain(`l${MAX_SHEET_ROWS}x`);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });

  it("a tab of MAX_SHEET_ROWS + 1 filled rows carries no tab, with the existing remark (rows 56); the next tab stays", async () => {
    const r = await run([{ name: "Long", cellsXml: rowsOf(MAX_SHEET_ROWS + 1, "l") }, small]);
    // Failure value (before, cap 100): the tab is published.
    expect(r.titles).toEqual(["Small"]);
    expect(r.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Long", rows: MAX_SHEET_ROWS + 1, cols: 1 });
  });

  it("non-regression: entirely empty rows do not count (MAX_SHEET_ROWS filled + 40 empty rows is published)", async () => {
    const r = await run([{ name: "Gaps", cellsXml: rowsOf(MAX_SHEET_ROWS, "g", 40) }]);
    expect(r.titles).toEqual(["Gaps"]);
    expect(r.json).toContain(`g${MAX_SHEET_ROWS}x`);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });
});

describe("the sheet row cap constant", () => {
  it("MAX_SHEET_ROWS is 55 (the fixtures above are built from it; at the old 100 they would pass vacuously)", () => {
    expect(MAX_SHEET_ROWS).toBe(55);
  });
});
