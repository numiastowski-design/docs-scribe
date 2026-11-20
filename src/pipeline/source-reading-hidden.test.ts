import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { GoogleSheetMeta } from "../extractor/google-sheet";
import { isDeclaredFormatSource } from "./conservation-source";
import { readDocxConservationSource } from "./source-reading-docx";
import { readGoogleSheetConservationSource } from "./source-reading-gsheet";

// Decision of 08/10/2026: hidden shapes (docx) and filtered rows (Google Sheet) are not published, so the
// independent control readings must not expect their words (else « mot perdu » for what was withheld on purpose).
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/pipeline/source-reading-hidden.test.ts

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const textBox = (words: string, id: number, hidden: string): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="${id}" name="b${id}"${hidden}/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${words}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const wordsOfSource = (source: Awaited<ReturnType<typeof readDocxConservationSource>>): string[] => {
  if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
  return source.units.flatMap((u) => u.tokens.map((t) => t.text));
};

describe("readDocxConservationSource — hidden drawings", () => {
  it.each(["1", "true", "TRUE", " 1 "])(
    '🔴 the words of a text box hidden="%s" are not expected; the visible box and host words are',
    async (value) => {
      const body = `<w:p><w:r><w:t>hostword</w:t></w:r>${textBox("hiddenboxword", 1, ` hidden="${value}"`)}${textBox("visibleboxword", 2, "")}</w:p>`;
      const words = wordsOfSource(await readDocxConservationSource(await docxOf(body)));
      // Failure value (before): "hiddenboxword" is expected.
      expect(words).not.toContain("hiddenboxword");
      expect(words).toContain("visibleboxword");
      expect(words).toContain("hostword");
    },
  );

  it.each([' hidden="0"', ""])("non-regression: a text box with docPr%s is still expected", async (attr) => {
    const words = wordsOfSource(
      await readDocxConservationSource(await docxOf(`<w:p>${textBox("shownboxword", 1, attr)}</w:p>`)),
    );
    expect(words).toContain("shownboxword");
  });
});

type RowFlags = { hiddenByUser?: boolean; hiddenByFilter?: boolean };
const gsheetMeta = (rowMetadata: RowFlags[]): GoogleSheetMeta => ({
  sheets: [{ properties: { title: "S", index: 0 }, rowMetadata }],
});

describe("readGoogleSheetConservationSource — filtered rows", () => {
  const wordsOf = (rowMetadata: RowFlags[], rows: string[][]): string[] => {
    const source = readGoogleSheetConservationSource(gsheetMeta(rowMetadata), new Map([["S", rows]]));
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    return source.units.flatMap((u) => u.tokens.map((t) => t.text));
  };

  it("🔴 a row filtered out is not expected; the visible rows still are", () => {
    const words = wordsOf([{}, { hiddenByFilter: true }, {}], [["headword"], ["filteredword"], ["keptword"]]);
    // Failure value (before): "filteredword" is expected.
    expect(words).not.toContain("filteredword");
    expect(words).toContain("headword");
    expect(words).toContain("keptword");
  });

  it("🔴 a row hidden by the user AND by a filter is not expected", () => {
    const words = wordsOf([{ hiddenByUser: true, hiddenByFilter: true }, {}], [["bothword"], ["keptword"]]);
    expect(words).not.toContain("bothword");
    expect(words).toContain("keptword");
  });

  it("non-regression: hiddenByFilter false or absent is still expected", () => {
    const words = wordsOf([{ hiddenByFilter: false }, {}], [["firstword"], ["secondword"]]);
    expect(words).toEqual(expect.arrayContaining(["firstword", "secondword"]));
  });
});
