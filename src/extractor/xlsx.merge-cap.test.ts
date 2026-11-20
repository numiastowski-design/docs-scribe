import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_MERGE_CELLS_PER_SHEET } from "./xlsx-cells";
import { extractXlsx } from "./xlsx";

// A worksheet declaring more merged ranges than any real workbook (172 at most on 68 real ones, ~1.8 M in the hostile
// file of 27/10/2026) is refused for safety. The hostile workbook is built here: this exercises a defence, not a rendering.
// Run: npx vitest run --project=main src/extractor/xlsx.merge-cap.test.ts

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

/** A one-sheet workbook whose sheet declares `merges` two-cell merged ranges, one per row (A{i}:B{i}). */
const workbookWithMerges = async (merges: number): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`),
  );
  const mergeXml = Array.from({ length: merges }, (_, i) => `<mergeCell ref="A${i + 1}:B${i + 1}"/>`).join("");
  zip.file(
    "xl/worksheets/sheet1.xml",
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>kept</t></is></c></row></sheetData>' +
      `<mergeCells count="${merges}">${mergeXml}</mergeCells></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

describe("extractXlsx — merged ranges per sheet are capped", () => {
  it("🔴 a sheet with exactly MAX_MERGE_CELLS_PER_SHEET merges is read", async () => {
    const { doc } = await extractXlsx(await workbookWithMerges(MAX_MERGE_CELLS_PER_SHEET));
    expect(doc.blocks.length).toBeGreaterThan(0);
  }, 30_000);

  it("🔴 a sheet with MAX_MERGE_CELLS_PER_SHEET + 1 merges is refused for safety", async () => {
    const outcome = await extractXlsx(await workbookWithMerges(MAX_MERGE_CELLS_PER_SHEET + 1)).then(
      () => "read",
      (e: unknown) => e,
    );
    expect(outcome).toMatchObject({
      code: "FILE_UNSAFE",
      message: "the sheet declares more merged ranges than any real workbook",
    });
  }, 30_000);
});
