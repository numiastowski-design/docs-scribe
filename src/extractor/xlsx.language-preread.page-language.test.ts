/** 08/10/2026, defect D6 (code review), observable guard WITH the page language given (the production call; the tester's
 *  file calls `extractXlsx(bytes, undefined, PAGE_SIZED_LIMITS)` alone, which skips the language pre-read). Sheets past the tab cap must not be decoded
 *  by the language pre-read. Run: npx vitest run --project=main src/extractor/xlsx.language-preread.page-language.test.ts */
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { MAX_WORKBOOK_TABS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractXlsx } from "./xlsx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const markedWorkbookOf = async (inline: boolean): Promise<Uint8Array> => {
  const total = 40;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      `xmlns:r="${REL}"><sheets>` +
      Array.from({ length: total }, (_, i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join(
        "",
      ) +
      "</sheets></workbook>",
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      Array.from(
        { length: total },
        (_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      ).join(""),
    ),
  );
  for (let i = 0; i < total; i++) {
    const cell = inline ? '<c r="A1" t="inlineStr"><is><t>w</t></is></c>' : '<c r="A1"><v>1</v></c>';
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData><!--SHEET${i + 1}--><row r="1">${cell}</row></sheetData></worksheet>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

describe("xlsx — with the page language given, the language pre-read never opens a sheet the tab cap hides", () => {
  it("only the sheets inside the cap are decoded (numbers-only AND short inline strings)", async () => {
    for (const inline of [false, true]) {
      const seen = new Set<number>();
      // eslint-disable-next-line @typescript-eslint/unbound-method -- taken before the spy replaces it, then called through `apply`, which binds `this`.
      const original = TextDecoder.prototype.decode as unknown as (this: unknown, ...args: unknown[]) => string;
      const spy = vi.spyOn(TextDecoder.prototype, "decode").mockImplementation(function (
        this: unknown,
        ...args: unknown[]
      ) {
        const text = original.apply(this, args);
        const m = /<!--SHEET(\d+)-->/.exec(text);
        if (m) seen.add(Number(m[1]));
        return text;
      });
      try {
        await extractXlsx(await markedWorkbookOf(inline), "en", PAGE_SIZED_LIMITS);
      } finally {
        spy.mockRestore();
      }
      // Failure value (before the fix): 1..40.
      expect(
        [...seen].sort((x, y) => x - y),
        `inline ${inline}`,
      ).toEqual(Array.from({ length: MAX_WORKBOOK_TABS }, (_, k) => k + 1));
    }
  });
});
