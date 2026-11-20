import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { readDocx } from "./index";

// The reader runs the layout-table pass on the BODY blocks only
// (the ones `readBlocks(body)` returns), never on a footnote, a text box or a table's cell. Here real .docx packages are built in memory and
// read by the real reader.
//   npx vitest run --project=main src/extractor/docx-reader/layout-tables.reader.test.ts
vi.setConfig({ testTimeout: 30_000 });

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:r="${REL}"`;

const build = async (body: string, footnotes?: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const footnoteRel =
    footnotes === undefined ? "" : `<Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/>`;
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${footnoteRel}</Relationships>`,
  );
  if (footnotes !== undefined)
    zip.file("word/footnotes.xml", `<?xml version="1.0"?><w:footnotes ${NS}>${footnotes}</w:footnotes>`);
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

/** `n` distinct words. */
const words = (n: number, tag: string): string =>
  Array.from({ length: n }, (_, i) => `${tag}${String.fromCharCode(97 + (i % 26))}${Math.floor(i / 26)}`).join(" ");
const cell = (text: string, tcPr = ""): string =>
  `<w:tc>${tcPr === "" ? "" : `<w:tcPr>${tcPr}</w:tcPr>`}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const tbl = (...rows: string[][]): string =>
  `<w:tbl><w:tblPr/><w:tblGrid/>${rows.map((r) => `<w:tr>${r.join("")}</w:tr>`).join("")}</w:tbl>`;
const para = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** Every block of the tree, with the kinds of its descendants. */
const everyBlock = (blocks: readonly IRBlock[]): IRBlock[] =>
  blocks.flatMap((b): IRBlock[] => {
    switch (b.kind) {
      case "table":
        return [b, ...everyBlock(b.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks)))];
      case "list":
        return [b, ...everyBlock(b.items.flat())];
      case "columns":
        return [b, ...everyBlock(b.cells.flat())];
      case "footnote":
      case "section":
        return [b, ...everyBlock(b.blocks)];
      case "tabs":
        return [b, ...everyBlock(b.tabs.flatMap((t) => t.blocks))];
      default:
        return [b];
    }
  });
const kindsOf = (blocks: readonly IRBlock[], kind: IRBlock["kind"]): number =>
  everyBlock(blocks).filter((b) => b.kind === kind).length;

const TWO_PROSE = (): string[] => [cell(words(12, "left")), cell(words(14, "right"))];

describe("the reader converts a layout table of the BODY", () => {
  it("a body-level one-row table of two prose cells is a `columns` block (two_equal), no table left", async () => {
    const { doc } = await readDocx(await build(para("Intro") + tbl(TWO_PROSE()) + para("Outro")));
    // Failure value: one `table`, no `columns`.
    expect(kindsOf(doc.blocks, "columns")).toBe(1);
    expect(kindsOf(doc.blocks, "table")).toBe(0);
    const columns = doc.blocks.find((b) => b.kind === "columns");
    expect(columns).toMatchObject({ kind: "columns", layout: "two_equal" });
    expect(JSON.stringify((columns as Extract<IRBlock, { kind: "columns" }>).cells[0])).toContain("lefta0");
    expect(JSON.stringify((columns as Extract<IRBlock, { kind: "columns" }>).cells[1])).toContain("righta0");
  });

  it("three prose cells: three_equal", async () => {
    const { doc } = await readDocx(
      await build(tbl([cell(words(10, "a")), cell(words(10, "b")), cell(words(10, "c"))])),
    );
    expect(doc.blocks.find((b) => b.kind === "columns")).toMatchObject({ layout: "three_equal" });
  });

  it("a form (label | value) stays a table", async () => {
    const { doc } = await readDocx(await build(tbl([cell("Nom"), cell("Dupont Jean")])));
    expect(kindsOf(doc.blocks, "table")).toBe(1);
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
  });

  it("9 words against 10 stays a table", async () => {
    const { doc } = await readDocx(await build(tbl([cell(words(10, "a")), cell(words(9, "b"))])));
    expect(kindsOf(doc.blocks, "table")).toBe(1);
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
  });

  it("a merged cell (gridSpan) stays a table", async () => {
    const { doc } = await readDocx(
      await build(tbl([cell(words(12, "a"), '<w:gridSpan w:val="2"/>'), cell(words(12, "b"))])),
    );
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
    expect(kindsOf(doc.blocks, "table")).toBe(1);
  });

  it("two rows stay a table", async () => {
    const { doc } = await readDocx(await build(tbl(TWO_PROSE(), TWO_PROSE())));
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
    expect(kindsOf(doc.blocks, "table")).toBe(1);
  });

  it("a layout-looking table nested in a cell of a data table stays a table (and the outer one is not converted either)", async () => {
    const inner = tbl(TWO_PROSE());
    const outer = `<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc>${inner}<w:p/></w:tc><w:tc>${para(words(12, "z"))}</w:tc></w:tr></w:tbl>`;
    const { doc } = await readDocx(await build(outer));
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
    expect(kindsOf(doc.blocks, "table")).toBe(2);
  });
});

describe("the reader never converts outside the body", () => {
  it("the same table inside a footnote stays a table", async () => {
    const note = `<w:footnote w:id="1">${tbl(TWO_PROSE())}<w:p><w:r><w:t>note end</w:t></w:r></w:p></w:footnote>`;
    const body = '<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const { doc } = await readDocx(await build(body, note));
    const footnote = doc.blocks.find((b) => b.kind === "footnote");
    expect(footnote, "the footnote is read").toBeDefined();
    expect(kindsOf(doc.blocks, "table")).toBe(1);
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
  });

  it("the same table inside a text box stays a table", async () => {
    const box =
      `<w:r><w:drawing><wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="1" name="b1"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></wps:spPr><wps:txbx><w:txbxContent>${tbl(TWO_PROSE())}<w:p/></w:txbxContent></wps:txbx>` +
      "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    const { doc } = await readDocx(await build(`<w:p><w:r><w:t>host</w:t></w:r>${box}</w:p>`));
    expect(JSON.stringify(doc.blocks)).toContain("lefta0");
    expect(kindsOf(doc.blocks, "table")).toBe(1);
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
  });

  it("the same table inside a VML text box (w:pict → v:textbox) stays a table", async () => {
    const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
    const box = `<w:r><w:pict ${V}><v:shape style="width:100pt"><v:textbox><w:txbxContent>${tbl(TWO_PROSE())}<w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
    const { doc } = await readDocx(await build(`<w:p><w:r><w:t>host</w:t></w:r>${box}</w:p>`));
    expect(JSON.stringify(doc.blocks)).toContain("lefta0");
    expect(kindsOf(doc.blocks, "table")).toBe(1);
    expect(kindsOf(doc.blocks, "columns")).toBe(0);
  });

  it("a table inside a block content control (w:sdt) of the body is a body table: it converts", async () => {
    const { doc } = await readDocx(await build(`<w:sdt><w:sdtContent>${tbl(TWO_PROSE())}</w:sdtContent></w:sdt>`));
    expect(kindsOf(doc.blocks, "columns")).toBe(1);
  });

  it("the body table converts while the footnote's identical table does not, in one document", async () => {
    const note = `<w:footnote w:id="1">${tbl(TWO_PROSE())}</w:footnote>`;
    const body = tbl(TWO_PROSE()) + '<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const { doc } = await readDocx(await build(body, note));
    expect(kindsOf(doc.blocks, "columns")).toBe(1);
    expect(kindsOf(doc.blocks, "table")).toBe(1);
  });
});
