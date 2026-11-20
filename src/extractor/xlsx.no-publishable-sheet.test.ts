import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_SHEET_ROWS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractXlsx } from "./xlsx";

// 06/11/2026: a workbook none of whose sheets fits the display caps is REFUSED (no page at all), not published
// empty (28 of 68 real workbooks). A sheet is publishable when it FITS (at most MAX_SHEET_ROWS rows and MAX_SHEET_COLS
// columns); one such sheet is enough to publish the workbook exactly as before. Built workbooks: no real one among
// the 28 sits in fixtures/.
// Run: npx vitest run --project=main src/extractor/xlsx.no-publishable-sheet.test.ts

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const workbookOf = async (rowCounts: readonly number[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets>` +
      rowCounts.map((_, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
      "</sheets></workbook>",
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      rowCounts
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  rowCounts.forEach((n, i) => {
    const rows = Array.from(
      { length: n },
      (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="inlineStr"><is><t>w${r + 1}x</t></is></c></row>`,
    ).join("");
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const tooBig = MAX_SHEET_ROWS + 1;

describe("extractXlsx — a workbook with no publishable sheet is refused", () => {
  it("🔴 a small ordinary workbook is published: no refusal (guard against an inverted rule)", async () => {
    const { doc } = await extractXlsx(await workbookOf([5]), undefined, PAGE_SIZED_LIMITS);
    expect(doc.blocks).toHaveLength(1);
  });

  it("🔴 every sheet over the row cap: refused with our own code", async () => {
    const outcome = await extractXlsx(await workbookOf([tooBig, tooBig + 10]), undefined, PAGE_SIZED_LIMITS).then(
      () => "read",
      (e: unknown) => e,
    );
    expect(outcome).toMatchObject({
      code: "NO_PUBLISHABLE_SHEET",
      message: "the workbook has no sheet that fits the page limits",
    });
  });

  it("🔴 one sheet that fits among too-large ones: published as before, the large ones withheld and said", async () => {
    const { doc, notices } = await extractXlsx(await workbookOf([tooBig, 5]), undefined, PAGE_SIZED_LIMITS);
    expect(doc.blocks).toHaveLength(1);
    expect(notices.some((n) => n.kind === "sheet-withheld")).toBe(true);
  });
});
