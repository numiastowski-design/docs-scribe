import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_TABLE_COLS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractXlsx } from "./xlsx";

// 08/10/2026: ONE column cap for tables and for worksheet tabs (MAX_TABLE_COLS was 10).
// Run: npx vitest run --project=main src/extractor/xlsx.sheet-cols-12.test.ts

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

const wideRow = (n: number, prefix: string): string =>
  `<row r="1">${Array.from({ length: n }, (_, i) => `<c r="${colLetter(i + 1)}1" t="inlineStr"><is><t>${prefix}${i + 1}</t></is></c>`).join("")}</row>`;
type Tabs = { titles: string[]; wideCols: number | undefined; notices: readonly { kind: string }[] };
const run = async (sheets: { name: string; cellsXml: string }[]): Promise<Tabs> => {
  const { doc, notices } = await extractXlsx(await workbookOf(sheets), undefined, PAGE_SIZED_LIMITS);
  const tabs = doc.blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  const table = tabs.tabs[0]?.blocks[0];
  return {
    titles: tabs.tabs.map((t) => t.title),
    wideCols: table?.kind === "table" ? table.rows[0]?.cells.length : undefined,
    notices,
  };
};
const small = { name: "Small", cellsXml: wideRow(2, "s") };

describe("xlsx — a tab of more than the column cap is withheld (sheet-withheld), a tab AT the cap is published", () => {
  it("a tab of MAX_TABLE_COLS filled columns is published (it was withheld from 11 on with the cap of 10)", async () => {
    const r = await run([{ name: "Wide", cellsXml: wideRow(MAX_TABLE_COLS, "w") }, small]);
    expect(r.titles).toEqual(["Wide", "Small"]);
    expect(r.wideCols).toBe(MAX_TABLE_COLS);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });

  it("a tab of MAX_TABLE_COLS + 1 filled columns carries no tab, with the existing remark; the next tab stays", async () => {
    const r = await run([{ name: "Wide", cellsXml: wideRow(MAX_TABLE_COLS + 1, "w") }, small]);
    expect(r.titles).toEqual(["Small"]);
    expect(r.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Wide", rows: 1, cols: MAX_TABLE_COLS + 1 });
  });

  it("entirely empty columns do not count: MAX_TABLE_COLS + 1 columns of which one is empty is published", async () => {
    const cells = Array.from({ length: MAX_TABLE_COLS + 1 }, (_, i) =>
      i === 4 ? "" : `<c r="${colLetter(i + 1)}1" t="inlineStr"><is><t>e${i + 1}</t></is></c>`,
    ).join("");
    const r = await run([{ name: "Gap", cellsXml: `<row r="1">${cells}</row>` }]);
    expect(r.titles).toEqual(["Gap"]);
    expect(r.wideCols).toBe(MAX_TABLE_COLS);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });
});
