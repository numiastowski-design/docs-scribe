import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readDocxConservationSource } from "../pipeline/source-reading-docx";
import { extractDocx } from "./docx";
import { extractGoogleDoc } from "./google-doc";

// Plan step 5: a docx is read under the limits it is given (`ReaderLimits`), the product's when it is given none, and the extraction and its
// independent second reading take the SAME limits. Run: npx vitest run --project=main src/extractor/reader-limits.docx.test.ts

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const cell = (text: string): string => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (...cells: string[]): string => `<w:tr>${cells.map(cell).join("")}</w:tr>`;
const grid = `<w:tbl>${row("alpha", "bravo", "charlie")}${row("delta", "echo", "foxtrot")}${row("golf", "hotel", "india")}</w:tbl>`;
const prose = "<w:p><w:r><w:t>kilo</w:t></w:r></w:p>";

const flat = (blocks: readonly { kind: string }[]): string => JSON.stringify(blocks);
const tokensOf = async (bytes: Uint8Array, limits?: Parameters<typeof extractDocx>[1]): Promise<string[]> => {
  const source = await readDocxConservationSource(bytes, limits === undefined ? undefined : { limits });
  if (!("units" in source)) throw new Error("the second reading failed");
  return source.units.flatMap((unit) => unit.tokens.map((token) => token.text));
};

describe("docx — the table caps are the limits it is given", () => {
  it("a 3 × 3 table is published under the product's caps", async () => {
    const result = await extractDocx(await docxOf(grid + prose));
    expect(flat(result.doc.blocks)).toContain("foxtrot");
    expect(result.notices ?? []).toEqual([]);
  });

  it("tableCols: 2 withholds it, says so, and keeps the prose (the extraction)", async () => {
    // FAILS WITH (limits ignored): expected '[{"kind":"table", … "foxtrot" …' not to contain 'foxtrot'.
    const result = await extractDocx(await docxOf(grid + prose), { tableCols: 2 });
    expect(flat(result.doc.blocks)).not.toContain("foxtrot");
    expect(flat(result.doc.blocks)).toContain("kilo");
    expect(result.notices).toContainEqual({ kind: "table-withheld", count: 1 });
  });

  it("tableRows: 2 withholds it too", async () => {
    const result = await extractDocx(await docxOf(grid + prose), { tableRows: 2 });
    expect(flat(result.doc.blocks)).not.toContain("foxtrot");
    expect(result.notices).toContainEqual({ kind: "table-withheld", count: 1 });
  });

  it("the Google Doc export takes the same limits", async () => {
    const result = await extractGoogleDoc(await docxOf(grid + prose), undefined, { tableCols: 2 });
    expect(flat(result.doc.blocks)).not.toContain("foxtrot");
  });

  it("the second reading expects none of the words of a table the page withholds under the SAME limits, and all of them without", async () => {
    const bytes = await docxOf(grid + prose);
    expect(await tokensOf(bytes)).toEqual(expect.arrayContaining(["alpha", "foxtrot", "kilo"]));
    // FAILS WITH (limits ignored): expected [ 'alpha', 'bravo', … ] not to include 'foxtrot' — the page withheld the table, the check still expects its words.
    const limited = await tokensOf(bytes, { tableCols: 2 });
    expect(limited).toContain("kilo");
    expect(limited).not.toContain("foxtrot");
  });
});
