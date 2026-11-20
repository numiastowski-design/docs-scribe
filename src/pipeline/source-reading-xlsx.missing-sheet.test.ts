/** Decision of 08/10/2026: a sheet whose part is absent from the archive is said (`sheet-unreadable`). The
 *  independent control reading must not expect words from it (nothing to expect, nothing to lose) and must not
 *  throw; the words of the other sheets are still expected.
 *  Runs in the "main" vitest project:
 *    npx vitest run --project=main src/pipeline/source-reading-xlsx.missing-sheet.test.ts */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { isDeclaredFormatSource } from "./conservation-source";
import { readXlsxConservationSource, readXlsxSource } from "./source-reading-xlsx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const workbookOf = async (sheets: readonly { name: string; word: string; absent?: boolean }[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets>` +
      sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
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
  sheets.forEach((s, i) => {
    if (s.absent) return;
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${s.word}</t></is></c></row></sheetData></worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const SHEETS = [
  { name: "Alpha", word: "alphaword" },
  { name: "Gone", word: "goneword", absent: true },
  { name: "Gamma", word: "gammaword" },
];

describe("readXlsxConservationSource / readXlsxSource — a sheet whose part is absent", () => {
  it("non-regression: the conservation source is a declared-format source (no read failure, no throw) and expects the other sheets' words (X6)", async () => {
    const source = await readXlsxConservationSource(await workbookOf(SHEETS));
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source, got a read failure");
    const words = source.units.flatMap((u) => u.tokens.map((t) => t.text));
    expect(words).toContain("alphaword");
    expect(words).toContain("gammaword");
  });

  it("non-regression: no unit expects a word for the absent sheet (nothing to expect)", async () => {
    const source = await readXlsxConservationSource(await workbookOf(SHEETS));
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    const words = source.units.flatMap((u) => u.tokens.map((t) => t.text));
    expect(words).not.toContain("goneword");
    for (const unit of source.units.filter((u) => u.label === "Gone")) expect(unit.tokens).toEqual([]);
  });

  it("non-regression: readXlsxSource does not throw and reads the other sheets' text; the absent sheet has no text", async () => {
    const { sheets } = await readXlsxSource(await workbookOf(SHEETS));
    const joined = sheets.map((s) => s.text).join("|");
    expect(joined).toContain("alphaword");
    expect(joined).toContain("gammaword");
    expect(joined).not.toContain("goneword");
  });
});
