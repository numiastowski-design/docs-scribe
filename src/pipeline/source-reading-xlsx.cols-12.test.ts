import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_TABLE_COLS, PAGE_SIZED_LIMITS } from "../extractor/page-caps.test.support";
import { isDeclaredFormatSource } from "./conservation-source";
import { readXlsxConservationSource } from "./source-reading-xlsx";

// 08/10/2026: the independent control reading follows the 12-column cap (MAX_TABLE_COLS).
// Run: npx vitest run --project=main src/pipeline/source-reading-xlsx.cols-12.test.ts

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

const wideSheet = (n: number): string =>
  `<row r="1">${Array.from({ length: n }, (_, i) => `<c r="${colLetter(i + 1)}1" t="inlineStr"><is><t>c${i + 1}unique</t></is></c>`).join("")}</row>`;

describe("readXlsxConservationSource — the column cap is 12", () => {
  it("a 13-column tab: withheld whole, none of its words expected", async () => {
    const source = await readXlsxConservationSource(
      await workbookOf([
        { name: "Wide", cellsXml: wideSheet(MAX_TABLE_COLS + 1) },
        { name: "Small", cellsXml: rowOf(1, 1, "next sheet unique") },
      ]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).not.toContain("c1unique");
    expect(allWords(source)).not.toContain(`c${MAX_TABLE_COLS + 1}unique`);
  });

  it("a 12-column tab: its words are expected (with the cap of 10 it was withheld)", async () => {
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "Wide", cellsXml: wideSheet(MAX_TABLE_COLS) }]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([]);
    expect(allWords(source)).toContain("c1unique");
    expect(allWords(source)).toContain(`c${MAX_TABLE_COLS}unique`);
  });
});
