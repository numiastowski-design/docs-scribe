import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractPptx } from "../pptx";
import { cellText } from "../span-single-first";

// 27/10/2026: a table row whose ONLY non-empty cell is the first renders as ONE cell with a colspan over the table's full width,
// on a slide as in the spreadsheets and Word (`span-single-first.ts`). Synthetic `a:tbl` through the real reader. Header rows: only the
// first row of a table with `a:tblPr firstRow` (`pptx-reader/table.ts`).

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const deckOf = async (tree: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/>`),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${tree}</p:spTree></p:cSld></p:sld>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const frame = (data: string): string =>
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="f2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${TABLE_URI}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;
const tc = (text: string, attrs = ""): string =>
  `<a:tc${attrs}><a:txBody><a:bodyPr/><a:p>${text === "" ? "" : `<a:r><a:t>${text}</a:t></a:r>`}</a:p></a:txBody><a:tcPr/></a:tc>`;
const tr = (...cells: string[]): string => `<a:tr h="100">${cells.join("")}</a:tr>`;
const tbl = (firstRow: boolean, ...rows: string[]): string =>
  `<a:tbl><a:tblPr${firstRow ? ' firstRow="1"' : ""}/><a:tblGrid>${'<a:gridCol w="100"/>'.repeat(3)}</a:tblGrid>${rows.join("")}</a:tbl>`;

type Table = Extract<IRBlock, { kind: "table" }>;
const tableOf = async (table: string): Promise<Table> => {
  const { doc } = await extractPptx(await deckOf(frame(table)));
  const blocks = doc.blocks.flatMap((b) => (b.kind === "tabs" ? b.tabs.flatMap((tab) => tab.blocks) : [b]));
  const t = blocks.find((b): b is Table => b.kind === "table");
  if (t === undefined) throw new Error("no table block");
  return t;
};
const shape = (t: Table): [number, number | undefined][] => t.rows.map((r) => [r.cells.length, r.cells[0]!.colspan]);

describe("a slide table row whose only text is in its first cell spans the table", () => {
  it("a leading title row spans the full width", async () => {
    const t = await tableOf(
      tbl(false, tr(tc("Title"), tc(""), tc("")), tr(tc("a"), tc("b"), tc("c")), tr(tc("d"), tc("e"), tc("f"))),
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
      tbl(false, tr(tc("a"), tc("b"), tc("c")), tr(tc("d"), tc("e"), tc("f")), tr(tc("Note"), tc(""), tc(""))),
    );
    expect(shape(t)).toEqual([
      [3, undefined],
      [3, undefined],
      [1, 3],
    ]);
  });

  it("a header row (a:tblPr firstRow) with text only in its first cell does not span", async () => {
    const t = await tableOf(tbl(true, tr(tc("Head"), tc(""), tc("")), tr(tc("a"), tc("b"), tc("c"))));
    expect(shape(t)).toEqual([
      [3, undefined],
      [3, undefined],
    ]);
    expect(t.rows[0]!.cells.every((c) => c.header === true)).toBe(true);
  });

  it("a vertically merged row (rowSpan / vMerge) does not span", async () => {
    const t = await tableOf(
      tbl(
        false,
        tr(tc("Side", ' rowSpan="2"'), tc(""), tc("")),
        tr(tc("", ' vMerge="1"'), tc(""), tc("")),
        tr(tc("a"), tc("b"), tc("c")),
      ),
    );
    expect(t.rows[0]!.cells[0]!.rowspan).toBe(2);
    expect(t.rows[0]!.cells).toHaveLength(3);
    expect(t.rows[0]!.cells.every((c) => c.colspan === undefined)).toBe(true);
  });

  it("a one-row table of prose in every cell is untouched (no cell is empty)", async () => {
    const t = await tableOf(tbl(false, tr(tc("one two three"), tc("four five six"), tc("seven eight nine"))));
    expect(shape(t)).toEqual([[3, undefined]]);
  });
});
