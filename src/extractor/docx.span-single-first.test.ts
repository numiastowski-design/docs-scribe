import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "./docx";
import { cellText } from "./span-single-first";

// 27/10/2026: a table row whose ONLY non-empty cell is the first renders as ONE cell with a colspan over the table's full width,
// in Word (and Google Doc, same reader) as in the spreadsheets (`span-single-first.ts`). Synthetic WordprocessingML through the real reader.

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const cell = (text: string, tcPr = ""): string =>
  `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ""}<w:p>${text === "" ? "" : `<w:r><w:t>${text}</w:t></w:r>`}</w:p></w:tc>`;
const row = (cells: string[], header = false): string =>
  `<w:tr>${header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${cells.join("")}</w:tr>`;
const tbl = (...rows: string[]): string =>
  `<w:tbl><w:tblGrid><w:gridCol w="1"/><w:gridCol w="1"/><w:gridCol w="1"/></w:tblGrid>${rows.join("")}</w:tbl>`;

type Table = Extract<IRBlock, { kind: "table" }>;
const blocksOf = async (body: string): Promise<IRBlock[]> => (await extractDocx(await docxOf(body))).doc.blocks;
const tableOf = async (body: string): Promise<Table> => {
  const t = (await blocksOf(body)).find((b): b is Table => b.kind === "table");
  if (t === undefined) throw new Error("no table block");
  return t;
};
/** Per row: [cell count, colspan of the first cell]. */
const shape = (t: Table): [number, number | undefined][] => t.rows.map((r) => [r.cells.length, r.cells[0]!.colspan]);

describe("a Word row whose only text is in its first cell spans the table", () => {
  it("a leading title row spans the full width", async () => {
    const t = await tableOf(
      tbl(
        row([cell("Title"), cell(""), cell("")]),
        row([cell("a"), cell("b"), cell("c")]),
        row([cell("d"), cell("e"), cell("f")]),
      ),
    );
    expect(shape(t)).toEqual([
      [1, 3],
      [3, undefined],
      [3, undefined],
    ]);
    expect(cellText(t.rows[0]!.cells[0]!)).toBe("Title");
  });

  it("a trailing note row spans the full width", async () => {
    const t = await tableOf(
      tbl(
        row([cell("a"), cell("b"), cell("c")]),
        row([cell("d"), cell("e"), cell("f")]),
        row([cell("Note"), cell(""), cell("")]),
      ),
    );
    expect(shape(t)).toEqual([
      [3, undefined],
      [3, undefined],
      [1, 3],
    ]);
  });

  it("a declared header row (w:tblHeader) with text only in its first cell does not span", async () => {
    const t = await tableOf(tbl(row([cell("Head"), cell(""), cell("")], true), row([cell("a"), cell("b"), cell("c")])));
    expect(shape(t)).toEqual([
      [3, undefined],
      [3, undefined],
    ]);
    expect(t.rows[0]!.cells.every((c) => c.header === true)).toBe(true);
  });

  it("a vertically merged row (w:vMerge) does not span", async () => {
    const restart = '<w:vMerge w:val="restart"/>';
    const cont = "<w:vMerge/>";
    const t = await tableOf(
      tbl(
        row([cell("Side", restart), cell(""), cell("")]),
        row([cell("", cont), cell(""), cell("")]),
        row([cell("a"), cell("b"), cell("c")]),
      ),
    );
    expect(t.rows[0]!.cells[0]!.rowspan).toBe(2);
    expect(t.rows[0]!.cells).toHaveLength(3);
    expect(t.rows[0]!.cells.every((c) => c.colspan === undefined)).toBe(true);
  });

  it("a one-row layout table (three prose cells) stays columns, untouched", async () => {
    const prose = (n: number): string => Array.from({ length: 12 }, (_, i) => `word${n}x${i}`).join(" ");
    const blocks = await blocksOf(tbl(row([cell(prose(1)), cell(prose(2)), cell(prose(3))])));
    const columns = blocks.find((b) => b.kind === "columns");
    expect(columns).toBeDefined();
    expect(columns!.kind === "columns" && columns!.cells).toHaveLength(3);
  });
});
