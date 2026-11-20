import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readSource } from "./source-reading-docx";
import { readXlsxSource } from "./source-reading-xlsx";

// 08/10/2026: a workbook or document with an unreadable styles part is PUBLISHED (default formatting), so its independent
// reading must not throw on it either, or the lost-words check would abstain for a published document.
const here = dirname(fileURLToPath(import.meta.url));
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("independent readings with an unreadable styles part", () => {
  it("xlsx: the reading still runs", async () => {
    const zip = await JSZip.loadAsync(
      new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx/09-Effectifs-Dates.xlsx"))),
    );
    zip.file("xl/styles.xml", '<?xml version="1.0"?><styleSheet <<< not xml');
    await expect(readXlsxSource(await zip.generateAsync({ type: "uint8array" }))).resolves.toBeDefined();
  });

  it("docx: the reading still runs and keeps the text", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    );
    zip.file(
      "_rels/.rels",
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    zip.file("word/styles.xml", '<?xml version="1.0"?><w:styles <<< not xml');
    zip.file(
      "word/document.xml",
      `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Kept words here</w:t></w:r></w:p></w:body></w:document>`,
    );
    const reading = await readSource(await zip.generateAsync({ type: "uint8array" }));
    expect(JSON.stringify(reading)).toContain("Kept");
  });
});
