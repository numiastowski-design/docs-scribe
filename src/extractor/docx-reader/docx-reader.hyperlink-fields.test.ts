import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock, Inline } from "doc-ir";
import { readDocx } from "./index";

// Word writes a hyperlink either as <w:hyperlink> or as a FIELD (HYPERLINK "url"): simple (w:fldSimple) or complex
// (fldChar begin / instrText / separate / result runs / end). The field's visible words must carry the href, once.
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/docx-reader/docx-reader.hyperlink-fields.test.ts
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}"`;
const URL = "https://example.org/page";

const build = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdL" Type="${REL}/hyperlink" Target="${URL}" TargetMode="External"/></Relationships>`,
  );
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${body}${a4}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const run = (t: string, bold = false): string =>
  `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const fld = (type: "begin" | "separate" | "end"): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (t: string): string => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
const p = (...parts: string[]): string => `<w:p>${parts.join("")}</w:p>`;
const complex = (instruction: string, ...results: string[]): string =>
  fld("begin") + instr(instruction) + fld("separate") + results.join("") + fld("end");

const firstParagraph = (blocks: readonly IRBlock[]): Inline[] => {
  const first = blocks.find((b) => b.kind === "paragraph");
  if (first?.kind !== "paragraph") throw new Error("no paragraph");
  return first.inlines;
};
const linked = (xs: readonly Inline[]): string =>
  xs
    .filter((i) => i.href !== undefined)
    .map((i) => i.text)
    .join("");
const unlinked = (xs: readonly Inline[]): string =>
  xs
    .filter((i) => i.href === undefined)
    .map((i) => i.text)
    .join("");
const full = (xs: readonly Inline[]): string => xs.map((i) => i.text).join("");
const hrefs = (xs: readonly Inline[]): string[] => [
  ...new Set(xs.flatMap((i) => (i.href === undefined ? [] : [i.href]))),
];

const read = async (body: string): Promise<Inline[]> => firstParagraph((await readDocx(await build(body))).doc.blocks);

describe("HYPERLINK as a field: the visible words carry the href", () => {
  it("F1 simple field (w:fldSimple): link words carry the href, text once, neighbours unlinked", async () => {
    const xs = await read(
      p(
        run("before "),
        `<w:fldSimple w:instr=" HYPERLINK &quot;${URL}&quot; ">${run("texte du lien")}</w:fldSimple>`,
        run(" after"),
      ),
    );
    expect(full(xs)).toBe("before texte du lien after");
    expect(linked(xs)).toBe("texte du lien");
    expect(hrefs(xs)).toEqual([URL]);
    expect(unlinked(xs)).toBe("before  after");
  });

  it("F2 complex field: link words carry the href, text once, neighbours unlinked", async () => {
    const xs = await read(p(run("before "), complex(` HYPERLINK "${URL}" `, run("texte du lien")), run(" after")));
    expect(full(xs)).toBe("before texte du lien after");
    expect(linked(xs)).toBe("texte du lien");
    expect(hrefs(xs)).toEqual([URL]);
    expect(unlinked(xs)).toBe("before  after");
  });

  it("F2b instruction split over two instrText", async () => {
    const xs = await read(
      p(
        run("before "),
        fld("begin") + instr(" HYPERLINK ") + instr(`"${URL}" `) + fld("separate") + run("texte du lien") + fld("end"),
        run(" after"),
      ),
    );
    expect(full(xs)).toBe("before texte du lien after");
    expect(linked(xs)).toBe("texte du lien");
    expect(hrefs(xs)).toEqual([URL]);
  });

  it("F2c instruction cut in the middle of the URL", async () => {
    const xs = await read(
      p(
        fld("begin") +
          instr(' HYPERLINK "https://example.org') +
          instr('/page" ') +
          fld("separate") +
          run("cut") +
          fld("end"),
      ),
    );
    expect(full(xs)).toBe("cut");
    expect(hrefs(xs)).toEqual([URL]);
  });

  it("F3 several result runs, one bold: the link covers all of them and the bold stays", async () => {
    const xs = await read(
      p(complex(` HYPERLINK "${URL}" `, run("plain "), run("heavy", true), run(" tail")), run(" after")),
    );
    expect(full(xs)).toBe("plain heavy tail after");
    expect(linked(xs)).toBe("plain heavy tail");
    expect(hrefs(xs)).toEqual([URL]);
    const heavy = xs.filter((i) => i.marks?.includes("strong"));
    expect(heavy.map((i) => i.text).join("")).toBe("heavy");
    expect(heavy.every((i) => i.href === URL)).toBe(true);
    expect(unlinked(xs)).toBe(" after");
  });

  it("F4 internal anchor switch \\l (complex): no external link, text published", async () => {
    const xs = await read(p(complex(' HYPERLINK \\l "bookmark1" ', run("jump here"))));
    expect(full(xs)).toBe("jump here");
    expect(xs.some((i) => /^https?:/.test(i.href ?? ""))).toBe(false);
  });

  it("F4b internal anchor switch \\l (simple): no external link, text published", async () => {
    const xs = await read(
      p(`<w:fldSimple w:instr=" HYPERLINK \\l &quot;sig&quot; ">${run("anchor only")}</w:fldSimple>`),
    );
    expect(full(xs)).toBe("anchor only");
    expect(xs.some((i) => /^https?:/.test(i.href ?? ""))).toBe(false);
  });

  it("F4c URL AND \\l fragment: the field is not turned into an external link to the bare URL by accident (text published)", async () => {
    const xs = await read(p(complex(` HYPERLINK "${URL}" \\l "frag" `, run("with fragment"))));
    expect(full(xs)).toBe("with fragment");
    // Either the fragment is kept (URL#frag) or no link at all; never the bare URL minus its fragment.
    expect(hrefs(xs).filter((h) => h === URL)).toEqual([]);
  });

  it("F5 forbidden scheme in the field (javascript:): text published, a kept href is the raw one", async () => {
    const xs = await read(p(run("x "), complex(' HYPERLINK "javascript:alert(1)" ', run("danger"))));
    expect(full(xs)).toBe("x danger");
    // Same contract as a w:hyperlink: the reader hands the raw target, the caller that renders the page drops it
    // (and holds that with its own test). Both outcomes pass here, but a kept href must be the raw one.
    for (const i of xs) if (i.href !== undefined) expect(i.href).toBe("javascript:alert(1)");
  });

  it("F5b an allowed https field carries its URL on its words, ready to be rendered as a link", async () => {
    const xs = await read(p(complex(` HYPERLINK "${URL}" `, run("go"))));
    expect(xs).toEqual([{ text: "go", href: URL }]);
  });

  it("F8 field inside a table cell", async () => {
    const cell = `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(complex(` HYPERLINK "${URL}" `, run("cell link")))}</w:tc>`;
    const other = `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(run("plain cell"))}</w:tc>`;
    const { doc } = await readDocx(
      await build(
        `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell}${other}</w:tr></w:tbl>`,
      ),
    );
    const table = doc.blocks.find((b) => b.kind === "table");
    if (table?.kind !== "table") throw new Error("no table");
    const c0 = table.rows[0]!.cells[0]!.blocks[0]!;
    if (c0.kind !== "paragraph") throw new Error("no paragraph in cell");
    expect(full(c0.inlines)).toBe("cell link");
    expect(hrefs(c0.inlines)).toEqual([URL]);
    const c1 = table.rows[0]!.cells[1]!.blocks[0]!;
    if (c1.kind !== "paragraph") throw new Error("no paragraph in cell");
    expect(hrefs(c1.inlines)).toEqual([]);
  });
});

describe("non-regression", () => {
  it("F6 PAGE / REF fields (complex and simple): result text published once, no link", async () => {
    const xs = await read(
      p(
        run("p. "),
        complex(" PAGE ", run("3")),
        run(" see "),
        complex(" REF bm1 \\h ", run("section two")),
        run(" and "),
        `<w:fldSimple w:instr=" PAGE ">${run("7")}</w:fldSimple>`,
      ),
    );
    expect(full(xs)).toBe("p. 3 see section two and 7");
    expect(hrefs(xs)).toEqual([]);
  });

  it("F7 an ordinary w:hyperlink keeps its href and text once", async () => {
    const xs = await read(
      p(run("before "), `<w:hyperlink r:id="rIdL">${run("plain link")}</w:hyperlink>`, run(" after")),
    );
    expect(full(xs)).toBe("before plain link after");
    expect(linked(xs)).toBe("plain link");
    expect(hrefs(xs)).toEqual([URL]);
    expect(unlinked(xs)).toBe("before  after");
  });

  it("F7b a w:hyperlink and a HYPERLINK field in one paragraph: field part is red, w:hyperlink part keeps its href", async () => {
    const xs = await read(
      p(
        `<w:hyperlink r:id="rIdL">${run("one")}</w:hyperlink>`,
        run(" "),
        complex(' HYPERLINK "https://example.org/two" ', run("two")),
      ),
    );
    expect(full(xs)).toBe("one two");
    expect(xs.filter((i) => i.text === "one").every((i) => i.href === URL)).toBe(true);
    expect(xs.filter((i) => i.text === "two").every((i) => i.href === "https://example.org/two")).toBe(true);
  });
});
