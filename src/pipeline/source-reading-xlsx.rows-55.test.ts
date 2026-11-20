import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_SHEET_ROWS, PAGE_SIZED_LIMITS } from "../extractor/page-caps.test.support";
import { isDeclaredFormatSource } from "./conservation-source";
import { readXlsxConservationSource } from "./source-reading-xlsx";

// 08/10/2026: the independent control reading follows the 55-row cap (MAX_SHEET_ROWS).
// Run: npx vitest run --project=main src/pipeline/source-reading-xlsx.rows-55.test.ts

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

/** `n` filled rows, one inline-string cell each, word `r${i}unique`. */
const longSheet = (n: number): string =>
  Array.from({ length: n }, (_, i) => rowOf(i + 1, 1, `r${i + 1}unique`)).join("");

describe("readXlsxConservationSource — the row cap is 55", () => {
  it("a tab of MAX_SHEET_ROWS + 1 rows: withheld whole, none of its words expected", async () => {
    const source = await readXlsxConservationSource(
      await workbookOf([
        { name: "Long", cellsXml: longSheet(MAX_SHEET_ROWS + 1) },
        { name: "Small", cellsXml: rowOf(1, 1, "next sheet unique") },
      ]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    // Failure value (before, cap 100): the unit is kept and its words are expected.
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).not.toContain("r1unique");
    expect(allWords(source)).not.toContain(`r${MAX_SHEET_ROWS + 1}unique`);
  });

  it("a tab of MAX_SHEET_ROWS rows: its words are expected", async () => {
    const source = await readXlsxConservationSource(
      await workbookOf([{ name: "Long", cellsXml: longSheet(MAX_SHEET_ROWS) }]),
      PAGE_SIZED_LIMITS,
    );
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([]);
    expect(allWords(source)).toContain("r1unique");
    expect(allWords(source)).toContain(`r${MAX_SHEET_ROWS}unique`);
  });
});

describe("the sheet row cap constant", () => {
  it("MAX_SHEET_ROWS is 55 (the fixtures above are built from it; at the old 100 they would pass vacuously)", () => {
    expect(MAX_SHEET_ROWS).toBe(55);
  });
});
