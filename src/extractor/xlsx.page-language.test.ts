/**
 * `extractXlsx` writes a cell's number or date in the PAGE language
 * (features/page-language.feature): built-in id 14 is read as the
 * page language's short date, « General » and masks follow its separators, a
 * hide-in-place code still hides, and the independent second reading takes
 * no language at all (§5: it compares positions, never the text of a number).
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readXlsxConservationSource } from "../pipeline/source-reading-xlsx";
import { extractXlsx } from "./xlsx";
import { PAGE_SIZED_LIMITS } from "./page-caps.test.support";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** One sheet, one row per cell: A1 id 14, A2 `#,##0.00`, A3 General, A4 `[$-40C]dddd d mmmm yyyy`, A5 `;;;`. */
const workbook = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.file(
    "xl/styles.xml",
    `${DECL}<styleSheet xmlns="${MAIN}"><numFmts count="3">` +
      '<numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="[$-40C]dddd d mmmm yyyy"/><numFmt numFmtId="166" formatCode=";;;"/></numFmts>' +
      '<cellXfs count="5"><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="0"/><xf numFmtId="165"/><xf numFmtId="166"/></cellXfs></styleSheet>',
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `${DECL}<worksheet xmlns="${MAIN}"><dimension ref="A1:A5"/><sheetData>` +
      '<row r="1"><c r="A1" s="0"><v>46053</v></c></row>' +
      '<row r="2"><c r="A2" s="1"><v>1234.5</v></c></row>' +
      '<row r="3"><c r="A3" s="2"><v>0.3333333333333333</v></c></row>' +
      '<row r="4"><c r="A4" s="3"><v>45000</v></c></row>' +
      '<row r="5"><c r="A5" s="4"><v>77</v></c></row>' +
      "</sheetData></worksheet>",
  );
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
};

const textOf = async (language?: "fr" | "en"): Promise<string> =>
  JSON.stringify((await extractXlsx(await workbook(), language, PAGE_SIZED_LIMITS)).doc);

describe("extractXlsx — the page language", () => {
  it("an English page: id 14 is m/d/yy, separators are English, [$-40C] still gives French words", async () => {
    const page = await textOf("en");
    for (const expected of ["1/31/26", "1,234.50", "0.333333333", "mercredi 15 mars 2023"])
      expect(page, expected).toContain(expected);
    expect(page).not.toContain("77");
  });

  it("a French page: id 14 is dd/mm/yyyy, separators are French, the code's own [$-40C] words stay French", async () => {
    const page = await textOf("fr");
    for (const expected of ["31/01/2026", "1 234,50", "0,333333333", "mercredi 15 mars 2023"])
      expect(page, expected).toContain(expected);
    expect(page).not.toContain("77");
  });

  it("no language given is English (today's pages)", async () => {
    expect(await textOf()).toBe(await textOf("en"));
  });

  it("the second reading takes no language: its signature is fixed and it never reads a number's text", async () => {
    // Arity 2 since 09/11/2026, and the second parameter is `limits`, never a language: a TypeScript `?` adds no
    // default value, so an optional parameter still counts in `Function.length`. What this pins is unchanged — the
    // second reading takes NO language — and the two identical readings below are the real proof.
    expect(readXlsxConservationSource.length).toBe(2);
    const a = await readXlsxConservationSource(await workbook(), PAGE_SIZED_LIMITS);
    const b = await readXlsxConservationSource(await workbook(), PAGE_SIZED_LIMITS);
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });
});
