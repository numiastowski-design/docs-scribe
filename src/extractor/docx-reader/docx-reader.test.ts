import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { crc32, deflateRawSync } from "node:zlib";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock, Inline } from "doc-ir";
import { assertNoPendingRef, PENDING_REF, readDocx } from "./index";
import { createDamagedError, openPackage, PartDamagedError } from "../ooxml/package";
import { PAGE_SIZED_LIMITS } from "../page-caps.test.support";

// Witnesses: two Word files (`fixtures/docx/01`, `02`) and real Word files (`fixtures/docx-reel/`). Every expected
// value below was read in the fixture's raw XML before the test was written.
const here = dirname(fileURLToPath(import.meta.url));

const textOf = (inlines: readonly Inline[]): string => inlines.map((i) => i.text).join("");

/** Every text-bearing block's text, in document order, nested content included. */
const plainTexts = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((b): string[] => {
    switch (b.kind) {
      case "heading":
      case "paragraph":
        return [textOf(b.inlines)];
      case "list":
        return b.items.flatMap((item) => plainTexts(item));
      case "table":
        return b.rows.flatMap((row) => row.cells.flatMap((c) => plainTexts(c.blocks)));
      case "footnote":
      case "section":
        return plainTexts(b.blocks);
      case "columns":
        return b.cells.flatMap((c) => plainTexts(c));
      default:
        return [];
    }
  });

describe("step 0 — archive, relationships, ordered XML, paragraphs and text", () => {
  it("elements are known by their namespace, whatever prefix the producer bound to it", async () => {
    const zip = await JSZip.loadAsync(await synthetic({ body: "" }));
    zip.file(
      "word/document.xml",
      '<?xml version="1.0"?><doc:document xmlns:doc="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        "<doc:body><doc:p><doc:r><doc:t>Any prefix</doc:t></doc:r></doc:p></doc:body></doc:document>",
    );
    const { doc } = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    expect(plainTexts(doc.blocks)).toEqual(["Any prefix"]);
  });
});

// ---------------------------------------------------------------------------
// Synthetic packages, for the spec points no witness exercises (a character
// style, a toggle cancelled by the style hierarchy, a merged cell…). Minimal
// on purpose: only the parts OPC requires, plus the one under test.

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

interface SyntheticParts {
  body: string;
  styles?: string;
  numbering?: string;
  footnotes?: string;
  endnotes?: string;
  /** Extra relationships of the main part: [id, type segment, target, external?]. */
  rels?: [string, string, string, boolean?][];
  /** Extra archive entries (media). */
  files?: Record<string, Uint8Array>;
}

const synthetic = async (parts: SyntheticParts): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const rels: [string, string, string, boolean?][] = [...(parts.rels ?? [])];
  if (parts.styles) {
    zip.file("word/styles.xml", `<?xml version="1.0"?><w:styles ${W_NS}>${parts.styles}</w:styles>`);
    rels.push(["rIdS", "styles", "styles.xml"]);
  }
  if (parts.numbering) {
    zip.file("word/numbering.xml", `<?xml version="1.0"?><w:numbering ${W_NS}>${parts.numbering}</w:numbering>`);
    rels.push(["rIdN", "numbering", "numbering.xml"]);
  }
  if (parts.footnotes) {
    zip.file("word/footnotes.xml", `<?xml version="1.0"?><w:footnotes ${W_NS}>${parts.footnotes}</w:footnotes>`);
    rels.push(["rIdF", "footnotes", "footnotes.xml"]);
  }
  if (parts.endnotes) {
    zip.file("word/endnotes.xml", `<?xml version="1.0"?><w:endnotes ${W_NS}>${parts.endnotes}</w:endnotes>`);
    rels.push(["rIdE", "endnotes", "endnotes.xml"]);
  }
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels
        .map(
          ([id, type, target, external]) =>
            `<Relationship Id="${id}" Type="${REL_BASE}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`,
        )
        .join("") +
      "</Relationships>",
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${parts.body}</w:body></w:document>`);
  for (const [path, bytes] of Object.entries(parts.files ?? {})) zip.file(path, bytes);
  return zip.generateAsync({ type: "uint8array" });
};

const headingsOf = (blocks: readonly IRBlock[]): [number, string][] =>
  blocks.flatMap((b): [number, string][] => (b.kind === "heading" ? [[b.level, textOf(b.inlines)]] : []));

/** The inline carrying exactly this text, anywhere in the document. */
const inlineWithText = (blocks: readonly IRBlock[], text: string): Inline | undefined => {
  for (const b of blocks) {
    if (b.kind === "heading" || b.kind === "paragraph") {
      const found = b.inlines.find((i) => i.text === text);
      if (found) return found;
    }
    const nested =
      b.kind === "list"
        ? b.items.flat()
        : b.kind === "table"
          ? b.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks))
          : b.kind === "footnote" || b.kind === "section"
            ? b.blocks
            : [];
    const deeper = inlineWithText(nested, text);
    if (deeper) return deeper;
  }
  return undefined;
};

describe("step 1 — headings through the basedOn chain, run marks", () => {
  it("a heading is found through basedOn, by the style's NAME, whatever its id", async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Titre2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="MonTitre"><w:name w:val="Mon titre"/><w:basedOn w:val="Titre2"/></w:style>';
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Titre2"/></w:pPr><w:r><w:t>Localised id</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="MonTitre"/></w:pPr><w:r><w:t>Derived style</w:t></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body, styles }));
    expect(headingsOf(doc.blocks)).toEqual([
      [2, "Localised id"],
      [2, "Derived style"],
    ]);
  });

  it("marks from a character style (w:rStyle) and its basedOn chain; direct formatting wins; styles toggle", async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="Bold para"/><w:rPr><w:b/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="StrongEmphasis"><w:name w:val="Strong emphasis"/><w:basedOn w:val="Emphasis"/><w:rPr><w:b w:val="true"/><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Loud"><w:name w:val="Loud"/><w:rPr><w:b/></w:rPr></w:style>';
    const body =
      "<w:p>" +
      '<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>styled</w:t></w:r>' +
      '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="StrongEmphasis"/></w:rPr><w:t>inherited</w:t></w:r>' +
      '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:i w:val="false"/></w:rPr><w:t>overridden</w:t></w:r>' +
      "</w:p>" +
      '<w:p><w:pPr><w:pStyle w:val="BoldPara"/></w:pPr>' +
      "<w:r><w:t>from paragraph style</w:t></w:r>" +
      '<w:r><w:rPr><w:rStyle w:val="Loud"/></w:rPr><w:t>toggled off</w:t></w:r>' +
      '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>direct off</w:t></w:r>' +
      "</w:p>";
    const { doc } = await readDocx(await synthetic({ body, styles }));
    expect(inlineWithText(doc.blocks, "styled")?.marks).toEqual(["em"]);
    expect(inlineWithText(doc.blocks, "inherited")?.marks).toEqual(["strong", "em", "sup"]);
    expect(inlineWithText(doc.blocks, "overridden")?.marks).toBeUndefined();
    // ECMA-376 §17.7.3: bold set by the paragraph style AND by the character
    // style toggles off; direct formatting sets the value outright.
    expect(doc.blocks[1]).toEqual({
      kind: "paragraph",
      inlines: [{ text: "from paragraph style", marks: ["strong"] }, { text: "toggled offdirect off" }],
    });
  });
});

describe("step 2b — w:ruby (CT_Ruby, ECMA-376 Part 1 §17.3.3.25): a phonetic guide", () => {
  it("publishes the base text, then the guide text in parentheses, as the document's own words", async () => {
    const body =
      "<w:p><w:r><w:ruby>" +
      '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/></w:rubyPr>' +
      "<w:rt><w:r><w:t>Signature client</w:t></w:r></w:rt>" +
      "<w:rubyBase><w:r><w:t>Montant</w:t></w:r></w:rubyBase>" +
      "</w:ruby></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["Montant (Signature client)"]);
  });

  it("keeps the base text's own formatting; the parentheses and guide text carry the run's, not the base run's", async () => {
    const body =
      "<w:p><w:r><w:rPr><w:i/></w:rPr><w:ruby>" +
      '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/></w:rubyPr>' +
      "<w:rt><w:r><w:t>guide</w:t></w:r></w:rt>" +
      "<w:rubyBase><w:r><w:rPr><w:b/></w:rPr><w:t>base</w:t></w:r></w:rubyBase>" +
      "</w:ruby></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(inlineWithText(doc.blocks, "base")?.marks).toEqual(["strong"]);
    expect(inlineWithText(doc.blocks, " (")?.marks).toEqual(["em"]);
    expect(inlineWithText(doc.blocks, ")")?.marks).toEqual(["em"]);
  });

  it("a ruby with no guide text (w:rt empty or absent) publishes the base text alone, no empty parentheses", async () => {
    const body =
      "<w:p><w:r><w:ruby>" +
      '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/></w:rubyPr>' +
      '<w:rt><w:r><w:t xml:space="preserve">  </w:t></w:r></w:rt>' +
      "<w:rubyBase><w:r><w:t>Solo</w:t></w:r></w:rubyBase>" +
      "</w:ruby></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["Solo"]);
  });
});

const para = (text: string): IRBlock => ({ kind: "paragraph", inlines: [{ text }] });
const listsOf = (blocks: readonly IRBlock[]): IRBlock[] => blocks.filter((b) => b.kind === "list");

describe("step 3 — lists, numbering continuity, numbered headings", () => {
  it("startOverride, lvlRestart, numId 0, numbering declared by the paragraph style, letters and roman numerals", async () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="7">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1-%2)"/><w:lvlRestart w:val="0"/></w:lvl>' +
      "</w:abstractNum>" +
      '<w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>' +
      '<w:num w:numId="4"><w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>';
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="H2"><w:name w:val="heading 2"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr></w:style>';
    const p = (style: string, text: string, numPr = ""): string =>
      `<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${numPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const body =
      p("H1", "Alpha") +
      p("H2", "Beta") +
      p("H1", "Gamma") +
      p("H2", "Delta") + // lvlRestart 0: level 2 never restarts, so b, not a
      p("H1", "Not numbered", '<w:numPr><w:numId w:val="0"/></w:numPr>') +
      p("", "Five", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>');
    const { doc } = await readDocx(await synthetic({ body, styles, numbering }));
    expect(headingsOf(doc.blocks)).toEqual([
      [1, "I. Alpha"],
      [2, "I-a) Beta"],
      [1, "II. Gamma"],
      [2, "II-b) Delta"],
      [1, "Not numbered"],
    ]);
    expect(listsOf(doc.blocks)).toEqual([{ kind: "list", ordered: true, start: 5, items: [[para("Five")]] }]);
  });
});

const wordDocx = (name: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../../fixtures/docx", name)));

type TableBlock = Extract<IRBlock, { kind: "table" }>;
const tablesOf = (blocks: readonly IRBlock[]): TableBlock[] =>
  blocks.filter((b): b is TableBlock => b.kind === "table");
const cellTexts = (table: TableBlock): string[][] =>
  table.rows.map((r) => r.cells.map((c) => plainTexts(c.blocks).join("\n")));

describe("step 4 — tables, merged cells, header rows, nesting", () => {
  it("docx/01: two tables, bold first row stays a mark, not a header", async () => {
    const { doc } = await readDocx(wordDocx("01-Guide-Onboarding-Atria.docx"));
    const tables = tablesOf(doc.blocks);
    expect(tables.map(cellTexts)).toEqual([
      [
        ["Service", "Contact", "Email"],
        ["Ressources Humaines", "Camille Berthier", "rh@atria-solutions.example"],
        ["Support Informatique", "Yanis Delacroix", "support-it@atria-solutions.example"],
        ["Office Management", "Inès Laurent", "office@atria-solutions.example"],
      ],
      [
        ["Catégorie", "Détail"],
        ["Santé", "Mutuelle familiale prise en charge à 70 %"],
        ["Transport", "Forfait mobilités durables, 400 € par an"],
        ["Restauration", "Titres-restaurant, valeur faciale 9,50 €"],
        ["Formation", "Budget individuel de 1 200 € par an"],
      ],
    ]);
    expect(tables[0]?.rows[0]?.cells[0]).toEqual({
      blocks: [{ kind: "paragraph", inlines: [{ text: "Service", marks: ["strong"] }] }],
    });
    expect(tables[0]?.rows[1]?.cells[0]).toEqual({ blocks: [para("Ressources Humaines")] });
  });

  it("docx/02: no table", async () => {
    const { doc } = await readDocx(wordDocx("02-Procedure-Support-IT.docx"));
    expect(tablesOf(doc.blocks)).toEqual([]);
    expect(headingsOf(doc.blocks).map(([level]) => level)).toEqual([1, 2, 2]);
  });

  it("gridSpan → colspan, vMerge → rowspan (gridBefore counted), tblHeader by its value, nested table", async () => {
    const cell = (text: string, tcPr = ""): string =>
      `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ""}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const row = (cells: string, trPr = ""): string => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ""}${cells}</w:tr>`;
    const nested = `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const body =
      "<w:tbl>" +
      row(cell("H1") + cell("H2") + cell("H3"), "<w:tblHeader/>") +
      row(cell("also header", '<w:gridSpan w:val="3"/>'), '<w:tblHeader w:val="true"/>') +
      row(cell("tall", '<w:vMerge w:val="restart"/>') + cell("b") + cell("c"), '<w:tblHeader w:val="0"/>') +
      row(cell("", "<w:vMerge/>") + cell("wide", '<w:gridSpan w:val="2"/>'), '<w:tblHeader w:val="false"/>') +
      row(
        cell("", '<w:vMerge w:val="continue"/>') +
          cell("x") +
          `<w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>${nested}<w:p/></w:tc>`,
      ) +
      row(cell("after gap") + cell("d") + cell("e", '<w:vMerge w:val="restart"/>')) +
      row(cell("f") + cell("", "<w:vMerge/>"), '<w:gridBefore w:val="1"/>') +
      "</w:tbl>";
    const { doc } = await readDocx(await synthetic({ body }));
    const [table] = tablesOf(doc.blocks);
    if (!table) throw new Error("no table");
    const shape = table.rows.map((r) =>
      r.cells.map((c) => [plainTexts(c.blocks).join("|"), c.colspan ?? 1, c.rowspan ?? 1, c.header === true]),
    );
    expect(shape).toEqual([
      [
        ["H1", 1, 1, true],
        ["H2", 1, 1, true],
        ["H3", 1, 1, true],
      ],
      [["also header", 3, 1, true]],
      [
        ["tall", 1, 3, false],
        ["b", 1, 1, false],
        ["c", 1, 1, false],
      ],
      [["wide", 2, 1, false]],
      [
        ["x", 1, 1, false],
        ["outer|inner", 1, 1, false],
      ],
      [
        ["after gap", 1, 1, false],
        ["d", 1, 1, false],
        ["e", 1, 2, false],
      ],
      [["f", 1, 1, false]],
    ]);
    const outer = table.rows[4]?.cells[1]?.blocks;
    expect(outer?.map((b) => b.kind)).toEqual(["paragraph", "table"]);
  });
});

/** Top-level blocks as short labels, to read an order at a glance. */
const outline = (blocks: readonly IRBlock[]): string[] =>
  blocks.map((b) =>
    b.kind === "image"
      ? `image:${b.alt ?? ""}`
      : b.kind === "paragraph" || b.kind === "heading"
        ? `${b.kind}:${textOf(b.inlines)}`
        : b.kind,
  );

describe("step 5 — images, text boxes and drawings, equations", () => {
  it("text boxes of one group are ordered top to bottom, then left to right; an inline picture splits its paragraph", async () => {
    const box = (text: string, x: number, y: number): string =>
      `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="10" cy="10"/></a:xfrm></wps:spPr>` +
      `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>`;
    const NS =
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
    const group =
      `<w:r><mc:AlternateContent ${NS}><mc:Choice Requires="wpg"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name=""/>` +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp>' +
      box("bottom", 0, 900) +
      box("top right", 500, 100) +
      box("top left", 100, 100) +
      "</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
      "<mc:Fallback><w:t>FALLBACK TEXT</w:t></mc:Fallback></mc:AlternateContent></w:r>";
    const picture =
      `<w:r><w:drawing ${NS}><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="2" name="p" descr="a dot"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      '<pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const body =
      `<w:p><w:r><w:t>host</w:t></w:r>${group}</w:p>` +
      `<w:p><w:r><w:t>before</w:t></w:r>${picture}<w:r><w:t>after</w:t></w:r></w:p>`;
    const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const { doc, images } = await readDocx(
      await synthetic({ body, rels: [["rIdImg", "image", "media/dot.png"]], files: { "word/media/dot.png": png } }),
    );
    expect(outline(doc.blocks)).toEqual([
      "paragraph:host",
      "paragraph:top left",
      "paragraph:top right",
      "paragraph:bottom",
      "paragraph:before",
      "image:", // descr="a dot" is not published (09/10/2026)
      "paragraph:after",
    ]);
    expect(doc.blocks[5]).toMatchObject({ ref: { widthPt: 20, mimeType: "image/png" } });
    expect(images).toHaveLength(1);
    expect(images[0]?.bytes).toEqual(png);
  });
});

describe("step 6 — table of contents, other controls, fields, footnotes, revisions", () => {
  it("fields show their displayed result only; other content controls show their content", async () => {
    const r = (inner: string): string => `<w:r>${inner}</w:r>`;
    const body =
      "<w:p>" +
      r('<w:t xml:space="preserve">Page </w:t>') +
      r('<w:fldChar w:fldCharType="begin"/>') +
      r('<w:instrText xml:space="preserve"> PAGE </w:instrText>') +
      r('<w:fldChar w:fldCharType="separate"/>') +
      r("<w:t>7</w:t>") +
      r('<w:fldChar w:fldCharType="end"/>') +
      r('<w:t xml:space="preserve"> of </w:t>') +
      '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>' +
      r('<w:fldChar w:fldCharType="begin"/><w:instrText> DATE </w:instrText><w:fldChar w:fldCharType="end"/>') +
      r('<w:fldChar w:fldCharType="begin"/><w:instrText> IF </w:instrText>') +
      r(
        '<w:fldChar w:fldCharType="begin"/><w:instrText> MERGEFIELD x </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>hidden</w:t><w:fldChar w:fldCharType="end"/>',
      ) +
      r('<w:fldChar w:fldCharType="separate"/><w:t xml:space="preserve">.</w:t><w:fldChar w:fldCharType="end"/>') +
      "</w:p>" +
      '<w:sdt><w:sdtPr><w:alias w:val="Client"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside a control</w:t></w:r>' +
      '<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t xml:space="preserve"> and inline</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt>';
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["Page 7 of 9.", "Inside a control and inline"]);
  });

  it("a note in footnotes.xml resolves its links through its own part's relationships", async () => {
    const footnotes =
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="4"><w:p><w:r><w:footnoteRef/></w:r><w:hyperlink r:id="rIdX"><w:r><w:t>source</w:t></w:r></w:hyperlink></w:p></w:footnote>';
    const body = '<w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="4"/></w:r></w:p>';
    const bytes = await synthetic({ body, footnotes });
    // footnotes.xml.rels is its own part: added by hand next to footnotes.xml.
    const zip = await JSZip.loadAsync(bytes);
    zip.file(
      "word/_rels/footnotes.xml.rels",
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdX" Type="${REL_BASE}/hyperlink" Target="https://example.org/source" TargetMode="External"/></Relationships>`,
    );
    const { doc } = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    expect(doc.blocks).toEqual([
      { kind: "paragraph", inlines: [{ text: "Claim" }, { text: "[1]", marks: ["sup"] }] },
      {
        kind: "footnote",
        id: "1",
        blocks: [{ kind: "paragraph", inlines: [{ text: "source", href: "https://example.org/source" }] }],
      },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Checkpoint A, A4: words the first six steps left behind.

describe("A4 — no word lost", () => {
  it("endnotes are read like footnotes: one call sequence, all notes grouped at the end in call order", async () => {
    const note = (tag: string, id: string, text: string): string =>
      `<w:${tag} w:id="${id}"><w:p><w:r><w:${tag}Ref/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:${tag}>`;
    const body =
      '<w:p><w:r><w:t>One</w:t></w:r><w:r><w:endnoteReference w:id="3"/></w:r>' +
      '<w:r><w:t xml:space="preserve"> two</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const { doc } = await readDocx(
      await synthetic({
        body,
        footnotes: note("footnote", "1", "A footnote."),
        endnotes: '<w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote>' + note("endnote", "3", "An endnote."),
      }),
    );
    expect(doc.blocks).toEqual([
      {
        kind: "paragraph",
        inlines: [{ text: "One" }, { text: "[1]", marks: ["sup"] }, { text: " two" }, { text: "[2]", marks: ["sup"] }],
      },
      { kind: "footnote", id: "1", blocks: [para("An endnote.")] },
      { kind: "footnote", id: "2", blocks: [para("A footnote.")] },
    ]);
  });

  it("a VML text box (w:pict → v:textbox → w:txbxContent) gives its paragraphs after its host; a VML line gives nothing", async () => {
    const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
    const body =
      "<w:p><w:r><w:t>Host</w:t></w:r>" +
      `<w:r><w:pict ${V}><v:shape style="width:100pt"><v:textbox><w:txbxContent>` +
      "<w:p><w:r><w:t>Boxed words</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>" +
      `<w:r><w:pict ${V}><v:rect style="width:0;height:1.5pt"/></w:pict></w:r></w:p>`;
    const { doc } = await readDocx(await synthetic({ body }));
    expect(outline(doc.blocks)).toEqual(["paragraph:Host", "paragraph:Boxed words"]);
  });

  it("w:sym gives the character its w:char declares", async () => {
    const body =
      '<w:p><w:r><w:t xml:space="preserve">Check </w:t><w:sym w:font="Wingdings" w:char="2713"/><w:t xml:space="preserve"> done</w:t></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["Check ✓ done"]);
  });

  it("🔴 w:sym in the Symbol font gives the character the point draws (θ), another font keeps its point (26/10/2026)", async () => {
    // FAILS WITH: « \uF071 » published — Confluence draws no θ for it, and search cannot find it.
    const body =
      '<w:p><w:r><w:t xml:space="preserve">cos(</w:t><w:sym w:font="Symbol" w:char="F071"/><w:t xml:space="preserve">) </w:t>' +
      '<w:sym w:font="Wingdings" w:char="F0B7"/></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["cos(θ) \uF0B7"]);
  });

  it("🔴 a w:sym in the surrogate range (D800–DFFF) is skipped, never a lone surrogate on the page (counter-review of 26/10/2026)", async () => {
    // FAILS WITH: « a\uD800b » — a lone surrogate published, invalid in the page's XML.
    const body = '<w:p><w:r><w:t>a</w:t><w:sym w:font="Arial" w:char="D800"/><w:t>b</w:t></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["ab"]);
  });

  it("w:noBreakHyphen is a hyphen", async () => {
    const body = "<w:p><w:r><w:t>Saint</w:t><w:noBreakHyphen/><w:t>Étienne</w:t></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["Saint-Étienne"]);
  });

  it("an abstractNum that only links to a numbering style (numStyleLink) numbers through that style's definition", async () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="10"><w:styleLink w:val="LegalList"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="11"><w:numStyleLink w:val="LegalList"/></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="11"/></w:num>';
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/></w:style>' +
      '<w:style w:type="numbering" w:styleId="LegalList"><w:name w:val="Legal list"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
    const body =
      `<w:p><w:pPr><w:pStyle w:val="H1"/>${numPr}</w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>` +
      `<w:p><w:pPr>${numPr}</w:pPr><w:r><w:t>Item</w:t></w:r></w:p>`;
    const { doc } = await readDocx(await synthetic({ body, styles, numbering }));
    expect(doc.blocks).toEqual([
      { kind: "heading", level: 1, inlines: [{ text: "A. " }, { text: "Scope" }] },
      { kind: "list", ordered: true, start: 2, items: [[para("Item")]] },
    ]);
  });

  it("equations: delimiters of m:d (declared or default), the n-ary operator's default ∫, no mathematical layout", async () => {
    const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const r = (t: string): string => `<m:r><m:t>${t}</m:t></m:r>`;
    const body =
      `<w:p><m:oMath ${M}>` +
      `<m:d><m:e>${r("x")}</m:e><m:e>${r("y")}</m:e></m:d>` +
      `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/><m:sepChr m:val=";"/></m:dPr><m:e>${r("a")}</m:e><m:e>${r("b")}</m:e></m:d>` +
      `<m:nary><m:sub>${r("0")}</m:sub><m:sup>${r("1")}</m:sup><m:e>${r("f")}</m:e></m:nary>` +
      "</m:oMath></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["(x|y) [a;b] ∫ 0 1 f"]);
  });

  it("lit le texte d'une équation écrite en modification suivie, et laisse tomber ce qui est supprimé", async () => {
    // An equation's text can sit inside a <w:ins>, not as a direct child of its <m:r>: read as direct children only,
    // the sentence lost its subject.
    const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const body =
      "<w:p>" +
      '<w:r><w:t xml:space="preserve">Declare </w:t></w:r>' +
      `<m:oMath ${M}><m:sSub><m:e><m:r><w:ins><m:t>H</m:t></w:ins></m:r></m:e><m:sub><m:r><w:ins><m:t>submitter</m:t></w:ins></m:r></m:sub></m:sSub></m:oMath>` +
      '<w:r><w:t xml:space="preserve"> now.</w:t></w:r>' +
      "</w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks).join("")).toContain("submitter");
  });

  it("une équation ne SOUDE pas ses voisins : ce qu'on sait ne pas savoir rendre laisse une frontière de mot", async () => {
    // The reader recognises the equation and knows it will not render its layout: it must leave a word boundary on both
    // sides, whatever the file puts around it. Without one the letters are all there and the word is gone, so search no
    // longer finds it. The guard is on the boundary, which depends on us alone, not on the spaces the file holds.
    const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const body =
      "<w:p>" +
      '<w:r><w:t xml:space="preserve">Si les vecteurs</w:t></w:r>' +
      `<m:oMath ${M}><m:r><m:t xml:space="preserve"> </m:t></m:r><m:r><m:t>u</m:t></m:r></m:oMath>` +
      "<w:r><w:t>et</w:t></w:r>" +
      "</w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    const text = plainTexts(doc.blocks).join("");
    expect(text, "le mot « vecteurs » doit rester un mot").toContain("vecteurs ");
    expect(text, "« et » doit rester un mot").toContain(" et");
    expect(text).not.toContain("vecteursu");
    expect(text).not.toContain("uet");
  });
});

// ---------------------------------------------------------------------------
// Checkpoint A feedback (24/09/2026).

const docxReel = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../../fixtures/docx-reel", name)));

/** The text block whose text starts with `prefix`, anywhere (table cells,
 *  list items included). */
const blockStartingWith = (blocks: readonly IRBlock[], prefix: string): IRBlock | undefined => {
  for (const b of blocks) {
    if ((b.kind === "paragraph" || b.kind === "heading") && textOf(b.inlines).startsWith(prefix)) return b;
    const nested =
      b.kind === "list"
        ? b.items.flat()
        : b.kind === "table"
          ? b.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks))
          : [];
    const found = blockStartingWith(nested, prefix);
    if (found) return found;
  }
  return undefined;
};

describe("F2 — paragraph alignment declared by w:jc", () => {
  it("docx-reel/10: the « Our Reference / Date » block (in a table cell) is right-aligned; a justified paragraph carries nothing", async () => {
    const { doc } = await readDocx(docxReel("10-lpc-invitation-to-tender.docx"));
    expect(blockStartingWith(doc.blocks, "Our Reference: ")).toMatchObject({ kind: "paragraph", align: "right" });
    expect(blockStartingWith(doc.blocks, "Date: 13 July 2026")).toMatchObject({ kind: "paragraph", align: "right" });
    expect(blockStartingWith(doc.blocks, "Reference Number:")).not.toHaveProperty("align");
  });

  it("from the style chain too, on headings as on paragraphs; end is right; a direct left cancels the style's center", async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Centered"><w:name w:val="Centered"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/><w:basedOn w:val="Centered"/></w:style>';
    const p = (pPr: string, text: string): string => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const body =
      p('<w:pStyle w:val="H1"/>', "Centred title") +
      p('<w:pStyle w:val="Centered"/><w:jc w:val="left"/>', "Left after all") +
      p('<w:jc w:val="end"/>', "At the end") +
      p('<w:jc w:val="both"/>', "Justified");
    const { doc } = await readDocx(await synthetic({ body, styles }));
    expect(doc.blocks).toEqual([
      { kind: "heading", level: 1, inlines: [{ text: "Centred title" }], align: "center" },
      para("Left after all"),
      { kind: "paragraph", inlines: [{ text: "At the end" }], align: "right" },
      para("Justified"),
    ]);
  });
});

describe("F1 — links to bookmarks become links to Confluence anchors", () => {
  it("in-paragraph and body-level bookmarks, a bookmark before a table, an untargeted bookmark, a link with no bookmark", async () => {
    const link = (anchor: string, text: string): string =>
      `<w:hyperlink w:anchor="${anchor}"><w:r><w:t>${text}</w:t></w:r></w:hyperlink>`;
    const body =
      `<w:p>${link("inside", "to inside")}<w:r><w:t xml:space="preserve"> </w:t></w:r>${link("table", "to table")}<w:r><w:t xml:space="preserve"> </w:t></w:r>${link("nowhere", "to nowhere")}</w:p>` +
      '<w:p><w:r><w:t>Plain</w:t></w:r><w:bookmarkStart w:id="1" w:name="inside"/><w:r><w:t xml:space="preserve"> target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
      '<w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>' +
      "<w:p><w:r><w:t>Untargeted</w:t></w:r></w:p>" +
      '<w:bookmarkStart w:id="3" w:name="table"/><w:bookmarkEnd w:id="3"/>' +
      "<w:tbl><w:tr><w:tc><w:p/><w:p><w:r><w:t>First cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(doc.blocks).toEqual([
      {
        kind: "paragraph",
        inlines: [
          { text: "to inside", href: "#inside" },
          { text: " " },
          { text: "to table", href: "#table" },
          { text: " to nowhere" },
        ],
      },
      { kind: "paragraph", inlines: [{ text: "", anchor: "inside" }, { text: "Plain target" }] },
      para("Untargeted"),
      {
        kind: "table",
        rows: [
          {
            cells: [
              { blocks: [{ kind: "paragraph", inlines: [{ text: "", anchor: "table" }, { text: "First cell" }] }] },
            ],
          },
        ],
      },
    ]);
  });
});

describe("F4 — a table of contents declared as a bare TOC field", () => {
  it("docx-reel/08: the TOC field becomes ONE toc block; its 16 cached entries and page numbers are gone", async () => {
    const { doc } = await readDocx(docxReel("08-ontario-diu-annual-report.docx"));
    const labels = outline(doc.blocks);
    expect(labels.filter((l) => l === "toc")).toHaveLength(1);
    const at = labels.indexOf("toc");
    expect(labels[at - 1]).toBe("heading:Table of Contents");
    expect(labels[at + 1]).toBe("heading:1. Introduction");
    expect(plainTexts(doc.blocks)).not.toContain("4.4 Retention period 10");
  });

  it("with or without a cached result, as w:fldChar or w:fldSimple; the text around the field stays; other fields keep their result", async () => {
    const r = (inner: string): string => `<w:r>${inner}</w:r>`;
    const body =
      // TOC with a cached result spanning two paragraphs, text before and after.
      "<w:p>" +
      r('<w:t xml:space="preserve">Before </w:t>') +
      r('<w:fldChar w:fldCharType="begin"/>') +
      r('<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText>') +
      r('<w:fldChar w:fldCharType="separate"/>') +
      r("<w:t>Cached entry 1</w:t>") +
      "</w:p>" +
      "<w:p>" +
      r("<w:t>Cached entry 2</w:t>") +
      r('<w:fldChar w:fldCharType="end"/>') +
      r('<w:t xml:space="preserve"> after</w:t>') +
      "</w:p>" +
      // TOC never updated: no result at all.
      "<w:p>" +
      r('<w:fldChar w:fldCharType="begin"/><w:instrText>TOC \\o</w:instrText><w:fldChar w:fldCharType="end"/>') +
      "</w:p>" +
      // TOC as a simple field.
      '<w:p><w:fldSimple w:instr=" TOC \\o &quot;1-2&quot; "><w:r><w:t>Simple cached entry</w:t></w:r></w:fldSimple></w:p>' +
      // Another field keeps its displayed result.
      "<w:p>" +
      r(
        '<w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>7</w:t><w:fldChar w:fldCharType="end"/>',
      ) +
      "</w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(outline(doc.blocks)).toEqual(["paragraph:Before ", "toc", "paragraph: after", "toc", "toc", "paragraph:7"]);
  });
});

describe("F1 × F4 — the links of a table of contents that is not read target nothing", () => {
  it("docx-reel/08: the bare TOC field's cached entries link to _Toc bookmarks; none of those becomes an anchor", async () => {
    const { doc } = await readDocx(docxReel("08-ontario-diu-annual-report.docx"));
    expect(JSON.stringify(doc)).not.toContain('"anchor"');
  });
});

// ---------------------------------------------------------------------------
// G1 — the TOC content control is read like any other; the TOC FIELD inside
// it is the toc block (24/09/2026: the wrapper is an editing
// convenience, the table of contents is the field).

/** The label of the block right before the one toc block (fails unless there is exactly one). */
const beforeToc = (blocks: readonly IRBlock[]): string | undefined => {
  const labels = outline(blocks);
  expect(labels.filter((l) => l === "toc")).toHaveLength(1);
  return labels[labels.indexOf("toc") - 1]?.replace(/^(paragraph|heading):/, "").trim();
};

describe("G1 — a title typed inside the table of contents control is kept", () => {
  it("docx-reel/06 (TOCHeading style): « Contents » then the toc block", async () => {
    expect(beforeToc((await readDocx(docxReel("06-homes-england-ims-registration.docx"))).doc.blocks)).toBe("Contents");
  });

  it("docx-reel/07 (no style): « Contents » then the toc block", async () => {
    expect(beforeToc((await readDocx(docxReel("07-moj-data-sharing-guidance.docx"))).doc.blocks)).toBe("Contents");
  });

  it("a TOC control with no TOC field inside still gives exactly one toc block, and none of its content", async () => {
    const body =
      '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>' +
      "<w:sdtContent><w:p><w:r><w:t>Entry typed by hand 3</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
      "<w:p><w:r><w:t>After</w:t></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(outline(doc.blocks)).toEqual(["toc", "paragraph:After"]);
  });
});

describe("G3 — Word hidden text (w:vanish) is not published", () => {
  it("direct, from a character style, from a paragraph style; w:val=0 shows; specVanish and webHidden are ignored; a hidden list item does not break its list", async () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="HiddenPara"><w:name w:val="Hidden para"/><w:rPr><w:vanish/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="HiddenChar"><w:name w:val="Hidden char"/><w:rPr><w:vanish w:val="true"/></w:rPr></w:style>';
    const numbering =
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>';
    const r = (rPr: string, text: string): string =>
      `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
    const item = (text: string, pRpr = "", runRpr = ""): string =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${pRpr}</w:pPr>${r(runRpr, text)}</w:p>`;
    const body =
      "<w:p>" +
      r("", "Shown") +
      r("<w:vanish/>", " HIDDEN-DIRECT") +
      r('<w:rStyle w:val="HiddenChar"/>', " HIDDEN-STYLE") +
      r('<w:vanish w:val="0"/>', " shown-too") +
      r("<w:specVanish/>", " spec") +
      r("<w:webHidden/>", " web") +
      "</w:p>" +
      `<w:p><w:pPr><w:pStyle w:val="HiddenPara"/></w:pPr>${r("", "HIDDEN-PARAGRAPH")}</w:p>` +
      item("One") +
      item("HIDDEN-ITEM", "<w:rPr><w:vanish/></w:rPr>", "<w:vanish/>") +
      item("Two");
    const { doc } = await readDocx(await synthetic({ body, styles, numbering }));
    expect(doc.blocks).toEqual([
      para("Shown shown-too spec web"),
      { kind: "list", ordered: true, items: [[para("One")], [para("Two")]] },
    ]);
  });
});

// The pptx fuzz test found the gap in the package layer both readers share;
// `extractDocx` masked it for Word, `readDocx` itself did not.
describe("a part whose XML does not parse fails with our own message, never the parser's", () => {
  it("word/document.xml with an unclosed tag inside its XML declaration", async () => {
    const zip = await JSZip.loadAsync(wordDocx("01-Guide-Onboarding-Atria.docx"));
    const xml = await zip.file("word/document.xml")!.async("string");
    const offset = xml.indexOf("?>") + 1;
    zip.file("word/document.xml", xml.slice(0, offset) + "<w:corruptedUnclosed" + xml.slice(offset));
    await expect(readDocx(await zip.generateAsync({ type: "uint8array" }))).rejects.toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid Word document",
    });
  });
});

// ---------------------------------------------------------------------------
// Review of 27/09/2026 (night), docx path security and fidelity. Every hostile
// package below is built IN MEMORY: no hostile binary is committed.

/** A zip written by hand, entry by entry, through node's own zlib: a
 *  decompression bomb (tens of MB of one byte deflate to a few KB) costs
 *  milliseconds to build, where JSZip's JavaScript deflate costs seconds. An
 *  entry given `brokenStream` carries bytes that are no DEFLATE stream at all. */
type HandEntry = { name: string; bytes: Uint8Array; declaredSize?: number } | { name: string; brokenStream: true };
const handZip = (entries: readonly HandEntry[]): Uint8Array => {
  const enc = new TextEncoder();
  const pieces: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = enc.encode(entry.name);
    const data = "bytes" in entry ? new Uint8Array(deflateRawSync(entry.bytes)) : new Uint8Array(32).fill(0xff);
    const crc = "bytes" in entry ? crc32(entry.bytes) : 0;
    const size = "bytes" in entry ? (entry.declaredSize ?? entry.bytes.length) : 1000;
    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(8, 8, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    const dir = new Uint8Array(46 + name.length);
    const dv = new DataView(dir.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    dv.setUint16(10, 8, true);
    dv.setUint32(16, crc, true);
    dv.setUint32(20, data.length, true);
    dv.setUint32(24, size, true);
    dv.setUint16(28, name.length, true);
    dv.setUint32(42, offset, true);
    dir.set(name, 46);
    pieces.push(local, data);
    central.push(dir);
    offset += local.length + data.length;
  }
  const dirSize = central.reduce((n, d) => n + d.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, dirSize, true);
  ev.setUint32(16, offset, true);
  const all = [...pieces, ...central, end];
  const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of all) {
    out.set(p, at);
    at += p.length;
  }
  return out;
};

/** The entries of a synthetic package, to be re-written by `handZip` with
 *  some of them replaced or added. */
const entriesOf = async (docx: Uint8Array): Promise<{ name: string; bytes: Uint8Array }[]> => {
  const zip = await JSZip.loadAsync(docx);
  const out: { name: string; bytes: Uint8Array }[] = [];
  for (const file of Object.values(zip.files))
    if (!file.dir) out.push({ name: file.name, bytes: await file.async("uint8array") });
  return out;
};

const DAMAGED = { code: "EXTRACT_FAILED", message: "the file is damaged or is not a valid Word document" };
/** A part past an inflation ceiling is a bomb, not damage: its own code, which proposes nothing (lot « fichier abîmé », 25/10/2026). */
const UNSAFE = { code: "FILE_UNSAFE" };

const DRAWING_NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const pictureRun = (relId: string): string =>
  `<w:r><w:drawing ${DRAWING_NS}><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="1" name="p"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

describe("review 27/09 night, finding 2 — decompression is bounded", () => {
  it("a main part that inflates past the per-part ceiling is refused as unsafe, never parsed", async () => {
    const entries = await entriesOf(await synthetic({ body: "" }));
    const padded = new TextEncoder().encode(
      `<?xml version="1.0"?><w:document ${W_NS}><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p>${" ".repeat(65_000_000)}</w:body></w:document>`,
    );
    const bomb = handZip(entries.map((e) => (e.name === "word/document.xml" ? { name: e.name, bytes: padded } : e)));
    expect(bomb.length).toBeLessThan(1_000_000);
    expect(
      await readDocx(bomb).then(
        () => "read",
        (e: unknown) => e,
      ),
    ).toMatchObject(UNSAFE);
  });

  it("each read counts against the part ceiling, and every read against the package ceiling", async () => {
    const media = (n: number) => ({ name: `word/media/m${n}.png`, bytes: new Uint8Array(800).fill(n) });
    const base = await entriesOf(await synthetic({ body: "" }));
    const bytes = handZip([
      ...base,
      media(1),
      media(2),
      media(3),
      { name: "word/media/big.png", bytes: new Uint8Array(1001) },
    ]);
    const sizeOf = (name: string): number => base.find((e) => e.name === name)?.bytes.length ?? 0;
    // openPackage itself reads the content types and the package relationships.
    const opened = sizeOf("[Content_Types].xml") + sizeOf("_rels/.rels");
    const partOnly = await openPackage(bytes, createDamagedError, { partMaxBytes: 1000, packageMaxBytes: 1_000_000 });
    await expect(partOnly.readBytes("word/media/big.png")).rejects.toMatchObject(UNSAFE);
    expect((await partOnly.readBytes("word/media/m1.png"))?.length).toBe(800);
    // A caller's tighter budget for one read (the pictures' ceiling).
    // Its own allowance, not the bomb ceiling: callers may leave the picture out, so it keeps the read-failure code.
    await expect(partOnly.readBytes("word/media/m1.png", 799)).rejects.toMatchObject(DAMAGED);

    const pkg = await openPackage(bytes, createDamagedError, {
      partMaxBytes: 1000,
      packageMaxBytes: opened + 2 * 800 + 400,
    });
    expect((await pkg.readBytes("word/media/m1.png"))?.length).toBe(800);
    expect((await pkg.readBytes("word/media/m2.png"))?.length).toBe(800);
    await expect(pkg.readBytes("word/media/m3.png")).rejects.toMatchObject(UNSAFE);
  });
});

describe("arbitrage 28/09/2026 — the package layer names WHICH refusal: a damaged part may be degraded, a ceiling never", () => {
  it("a broken stream and an XML that does not parse are PartDamagedError; a part over its ceiling is a plain refusal", async () => {
    const entries = await entriesOf(await synthetic({ body: "" }));
    const bytes = handZip([
      ...entries,
      { name: "word/media/broken.png", brokenStream: true },
      { name: "word/media/big.png", bytes: new Uint8Array(2000).fill(7) },
      // The one XML shape the parser is known to throw on (see
      // corruptDeclaration above): an unclosed tag inside the declaration.
      { name: "word/notxml.xml", bytes: new TextEncoder().encode('<?xml version="1.0"?<broken><w:a/>') },
    ]);
    const pkg = await openPackage(bytes, createDamagedError, { partMaxBytes: 1000, packageMaxBytes: 1_000_000 });
    // Both degradable refusals carry the class AND the plain message: caught
    // nowhere, they refuse exactly as before.
    await expect(pkg.readBytes("word/media/broken.png")).rejects.toBeInstanceOf(PartDamagedError);
    await expect(pkg.readBytes("word/media/broken.png")).rejects.toMatchObject(DAMAGED);
    await expect(pkg.readXml("word/notxml.xml")).rejects.toBeInstanceOf(PartDamagedError);
    // The ceilings' refusal must never be degradable (a bomb is not a
    // document): its own code, deliberately NOT the class a caller may catch.
    const overCap: unknown = await pkg.readBytes("word/media/big.png").then(
      () => "read",
      (e: unknown) => e,
    );
    expect(overCap).toMatchObject(UNSAFE);
    expect(overCap).not.toBeInstanceOf(PartDamagedError);
  });
});

describe("review 27/09 night, finding 9 — pictures are loaded only when the walk uses one, under a ceiling", () => {
  it("a picture part no drawing uses is never inflated: its broken stream does not refuse the document", async () => {
    const entries = await entriesOf(
      await synthetic({
        body: "<w:p><w:r><w:t>Words</w:t></w:r></w:p>",
        rels: [["rIdUnused", "image", "media/unused.png"]],
      }),
    );
    const { doc, images } = await readDocx(
      handZip([...entries, { name: "word/media/unused.png", brokenStream: true }]),
    );
    expect(plainTexts(doc.blocks)).toEqual(["Words"]);
    expect(images).toEqual([]);
  });

  it("pictures past the image ceiling together are left out in document order, the text is published and the loss is counted (14/10/2026)", async () => {
    const body = `<w:p><w:r><w:t>Words</w:t></w:r></w:p><w:p>${pictureRun("rIdA")}${pictureRun("rIdB")}${pictureRun("rIdC")}</w:p>`;
    const entries = await entriesOf(
      await synthetic({
        body,
        rels: [
          ["rIdA", "image", "media/a.png"],
          ["rIdB", "image", "media/b.png"],
          ["rIdC", "image", "media/c.png"],
        ],
      }),
    );
    const half = 25_100_000;
    const bytes = handZip([
      ...entries,
      { name: "word/media/a.png", bytes: new Uint8Array(half) },
      { name: "word/media/b.png", bytes: new Uint8Array(half).fill(1) },
      // Small enough for what is left of the budget: taken, though it comes after a refused one.
      { name: "word/media/c.png", bytes: new Uint8Array(1000).fill(2) },
    ]);
    // Settled into a small value first: a failure must not ask the reporter to print 50 MB of pixels.
    const read = await readDocx(bytes).then(
      (r) => ({ texts: plainTexts(r.doc.blocks), images: r.images.length, notices: r.notices ?? [] }),
      (e: unknown) => ({ error: String(e) }),
    );
    expect(read).toEqual({ texts: ["Words"], images: 2, notices: [{ kind: "image-missing", count: 1 }] });
  });

  it("twenty skipped 30 MB pictures do not spend the package budget: the text is published, the skipped are counted (review 14/10/2026)", async () => {
    const ids = Array.from({ length: 20 }, (_, i) => `rId${i + 1}`);
    const entries = await entriesOf(
      await synthetic({
        body: `<w:p><w:r><w:t>Words</w:t></w:r></w:p><w:p>${ids.map((id) => pictureRun(id)).join("")}</w:p>`,
        rels: ids.map((id, i) => [id, "image", `media/p${i}.emf`] as [string, string, string]),
      }),
    );
    const heavy = new Uint8Array(30_000_000);
    const bytes = handZip([...entries, ...ids.map((_, i) => ({ name: `word/media/p${i}.emf`, bytes: heavy }))]);
    const read = await readDocx(bytes).then(
      (r) => ({ texts: plainTexts(r.doc.blocks), notices: r.notices ?? [] }),
      (e: unknown) => ({ error: String(e) }),
    );
    // The first fits (30 MB of the 50), the other nineteen are over what is left and are never inflated.
    expect(read).toEqual({ texts: ["Words"], notices: [{ kind: "image-missing", count: 19 }] });
  });

  it("a picture whose header declares less than it inflates to is still stopped at the allowance and left out", async () => {
    const entries = await entriesOf(
      await synthetic({
        body: `<w:p><w:r><w:t>Words</w:t></w:r></w:p><w:p>${pictureRun("rIdA")}</w:p>`,
        rels: [["rIdA", "image", "media/a.png"]],
      }),
    );
    const bytes = handZip([
      ...entries,
      { name: "word/media/a.png", bytes: new Uint8Array(60_000_000), declaredSize: 1000 },
    ]);
    const read = await readDocx(bytes).then(
      (r) => ({ texts: plainTexts(r.doc.blocks), images: r.images.length, notices: r.notices ?? [] }),
      (e: unknown) => ({ error: String(e) }),
    );
    expect(read).toEqual({ texts: ["Words"], images: 0, notices: [{ kind: "image-missing", count: 1 }] });
  });

  it("a single picture over the whole image ceiling is skipped alone", async () => {
    const entries = await entriesOf(
      await synthetic({
        body: `<w:p><w:r><w:t>Words</w:t></w:r></w:p><w:p>${pictureRun("rIdA")}</w:p>`,
        rels: [["rIdA", "image", "media/a.png"]],
      }),
    );
    const bytes = handZip([...entries, { name: "word/media/a.png", bytes: new Uint8Array(50_100_000) }]);
    const read = await readDocx(bytes).then(
      (r) => ({ texts: plainTexts(r.doc.blocks), images: r.images.length, notices: r.notices ?? [] }),
      (e: unknown) => ({ error: String(e) }),
    );
    expect(read).toEqual({ texts: ["Words"], images: 0, notices: [{ kind: "image-missing", count: 1 }] });
  });
});

describe("review 27/09 night, finding 1c — a picture's declared content type crosses only from our list", () => {
  it("a content type carrying CR LF becomes application/octet-stream; a listed one is kept", async () => {
    const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const docx = await synthetic({
      body: `<w:p>${pictureRun("rIdE")}</w:p><w:p>${pictureRun("rIdP")}</w:p>`,
      rels: [
        ["rIdE", "image", "media/evil.png"],
        ["rIdP", "image", "media/fine.png"],
      ],
      files: { "word/media/evil.png": png, "word/media/fine.png": new Uint8Array([...png, 1]) },
    });
    const zip = await JSZip.loadAsync(docx);
    const types = await zip.file("[Content_Types].xml")!.async("string");
    zip.file(
      "[Content_Types].xml",
      types.replace(
        "</Types>",
        '<Override PartName="/word/media/evil.png" ContentType="image/png&#13;&#10;X-Injected: 1"/></Types>',
      ),
    );
    const { doc, images } = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    expect(images.map((i) => i.mimeType)).toEqual(["application/octet-stream", "image/png"]);
    expect(doc.blocks.map((b) => (b.kind === "image" ? b.ref.mimeType : b.kind))).toEqual([
      "application/octet-stream",
      "image/png",
    ]);
  });
});

describe("review 27/09 night, findings 3 and 4 — numbering values are bounded", () => {
  const numbered = (numId: string, ilvl: string, text: string): string =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const decimalLevel = (ilvl: string, start: string): string =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${start}"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>`;

  it("w:start and w:startOverride are clamped to Word's own 0..32767", async () => {
    const numbering =
      `<w:abstractNum w:abstractNumId="0">${decimalLevel("0", "99999")}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1">${decimalLevel("0", "-5")}</w:abstractNum>` +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
      '<w:num w:numId="3"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="70000"/></w:lvlOverride></w:num>';
    const pause = "<w:p><w:r><w:t>break</w:t></w:r></w:p>";
    const body =
      numbered("1", "0", "Big") + pause + numbered("2", "0", "Negative") + pause + numbered("3", "0", "Override");
    const { doc } = await readDocx(await synthetic({ body, numbering }));
    expect(doc.blocks).toEqual([
      { kind: "list", ordered: true, start: 32767, items: [[para("Big")]] },
      para("break"),
      { kind: "list", ordered: true, start: 0, items: [[para("Negative")]] },
      para("break"),
      { kind: "list", ordered: true, start: 32767, items: [[para("Override")]] },
    ]);
  });

  // Not run red: before the clamp this start made roman() loop for ever
  // (1e21 − 1000 === 1e21), which would hang the runner, not fail it.
  it("a roman or letter numbering starting at 1e21 is spelled from 32767, in bounded time", async () => {
    const level = (ilvl: string, fmt: string): string =>
      `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1000000000000000000000"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${Number(ilvl) + 1}."/></w:lvl>`;
    const numbering =
      `<w:abstractNum w:abstractNumId="0">${level("0", "upperRoman")}${level("1", "lowerLetter")}</w:abstractNum>` +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const styles = '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/></w:style>';
    const heading = (ilvl: string): string =>
      `<w:p><w:pPr><w:pStyle w:val="H1"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>T</w:t></w:r></w:p>`;
    const { doc } = await readDocx(await synthetic({ body: heading("0") + heading("1"), styles, numbering }));
    const labels = doc.blocks.map((b) => (b.kind === "heading" ? b.inlines[0]?.text : undefined));
    expect(labels[0]).toBe(`${"M".repeat(32)}DCCLXVII. `);
    expect(labels[1]).toBe(`${"g".repeat(1261)}. `);
  });

  it("an ilvl past ECMA-376's nine levels (0..8) is read as the deepest level, never as an index", async () => {
    const numbering =
      `<w:abstractNum w:abstractNumId="0">${decimalLevel("0", "1")}${decimalLevel("100000", "1")}</w:abstractNum>` +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const body = numbered("1", "100000", "Deep") + numbered("1", "0", "Top");
    const { doc } = await readDocx(await synthetic({ body, numbering }));
    expect(doc.blocks).toEqual([para("Deep"), { kind: "list", ordered: true, items: [[para("Top")]] }]);
  });
});

describe("review 27/09 finding 5, centralized 28/09 — the reader hands hrefs RAW; the page's one exit door filters schemes", () => {
  it("the IR carries every external target verbatim", async () => {
    const targets: [string, string][] = [
      ["rIdJ", "javascript:alert(1)"],
      ["rIdD", "data:text/html,x"],
      ["rIdF", "file:///C:/secret.txt"],
      ["rIdR", "other.docx"],
      ["rIdH", "https://example.org/a"],
      ["rIdU", "HTTP://EXAMPLE.ORG/B"],
      ["rIdM", "mailto:someone@example.org"],
    ];
    const body = targets
      .map(([id]) => `<w:p><w:hyperlink r:id="${id}"><w:r><w:t>${id}</w:t></w:r></w:hyperlink></w:p>`)
      .join("");
    const rels = targets.map(([id, target]): [string, string, string, boolean] => [id, "hyperlink", target, true]);
    const { doc } = await readDocx(await synthetic({ body, rels }));
    // The reader no longer decides publication (arbitrage of 28/09/2026): the raw target travels, and the caller that
    // renders the page filters the schemes. That second half is held by the caller, on this very document.
    expect(doc.blocks).toEqual(
      targets.map(([id, target]) => ({ kind: "paragraph", inlines: [{ text: id, href: target }] })),
    );
  });

  it("an external target spelling one of OUR internal forms (drive:, #) is never trusted from a file: text, no link", async () => {
    // `#name` is minted from a bookmark, `drive:` by normalize-links from an
    // http Drive URL — a file DECLARING either would reach the renderer's
    // internal handling with a provenance it never had.
    const body =
      '<w:p><w:hyperlink r:id="rIdV"><w:r><w:t>fake-drive</w:t></w:r></w:hyperlink></w:p>' +
      '<w:p><w:hyperlink r:id="rIdA"><w:r><w:t>fake-anchor</w:t></w:r></w:hyperlink></w:p>';
    const rels: [string, string, string, boolean][] = [
      ["rIdV", "hyperlink", "drive:1234567890", true],
      ["rIdA", "hyperlink", "#note-1", true],
    ];
    const { doc } = await readDocx(await synthetic({ body, rels }));
    expect(doc.blocks).toEqual([para("fake-drive"), para("fake-anchor")]);
  });
});

describe("review 27/09 night, finding 7 — XML depth is bounded where the tree is built", () => {
  // document + body, two levels per control, then paragraph, run, text.
  const nested = (controls: number): string =>
    "<w:sdt><w:sdtContent>".repeat(controls) +
    "<w:p><w:r><w:t>deep</w:t></w:r></w:p>" +
    "</w:sdtContent></w:sdt>".repeat(controls);

  it("a document nested ~500 deep is read, the walk included", async () => {
    const { doc } = await readDocx(await synthetic({ body: nested(245) }));
    expect(plainTexts(doc.blocks)).toEqual(["deep"]);
  });

  it("past the ceiling, the document is refused as damaged — never a RangeError", async () => {
    await expect(readDocx(await synthetic({ body: nested(260) }))).rejects.toMatchObject(DAMAGED);
    await expect(readDocx(await synthetic({ body: nested(20_000) }))).rejects.toMatchObject(DAMAGED);
  });
});

describe("review 27/09 night, finding 8 — no DTD is read (OPC forbids them)", () => {
  it("a part declaring a DOCTYPE with its own entity is refused as damaged, the entity never expanded", async () => {
    const zip = await JSZip.loadAsync(await synthetic({ body: "" }));
    zip.file(
      "word/document.xml",
      `<?xml version="1.0"?><!DOCTYPE w:document [<!ENTITY who "Mallory">]><w:document ${W_NS}><w:body><w:p><w:r><w:t>&who;</w:t></w:r></w:p></w:body></w:document>`,
    );
    await expect(readDocx(await zip.generateAsync({ type: "uint8array" }))).rejects.toMatchObject(DAMAGED);
  });

  // No bound on entity references exists any more (header of `xml.ts`): the reader refuses every entity but the five
  // predefined, so none expands. A part holding a million `&lt;` is read whole.
  it("references expand to no more than they spell: a million `&lt;` are read, whole", async () => {
    const { doc } = await readDocx(
      await synthetic({ body: `<w:p><w:r><w:t>${"&lt;".repeat(1_000_000)}</w:t></w:r></w:p>` }),
    );
    expect(plainTexts(doc.blocks)).toEqual(["<".repeat(1_000_000)]);
  });
});

// 26/10/2026, switch of the XML reader: `package.ts` hands a part's BYTES to the reader, which finds the encoding (XML 1.0
// Appendix F, `xml/decode.ts`), instead of decoding every part as UTF-8 first (a valid UTF-16 part came out as gibberish,
// then was refused as damaged, no root found; a Windows-1252 part with an accent published U+FFFD).
describe("the encoding of a part is the part's own, read from its bytes", () => {
  const SENTENCE = "Voilà le café, naïve — Привет 日本語 €";
  const paragraph = (declaration: string): string =>
    `${declaration}<w:document ${W_NS}><w:body><w:p><w:r><w:t>${SENTENCE}</w:t></w:r></w:p></w:body></w:document>`;
  const utf16le = (text: string): Uint8Array => new Uint8Array(Buffer.from(`﻿${text}`, "utf16le"));
  const utf16be = (text: string): Uint8Array => utf16le(text).map((_, i, all) => all[i ^ 1] ?? 0);
  const withDocument = async (bytes: Uint8Array): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(await synthetic({ body: "" }));
    zip.file("word/document.xml", bytes);
    return zip.generateAsync({ type: "uint8array" });
  };
  const published = async (bytes: Uint8Array): Promise<string[]> =>
    plainTexts((await readDocx(await withDocument(bytes))).doc.blocks);

  it("UTF-16 with its byte-order mark, either order, is read: the sentence is published", async () => {
    const xml = paragraph('<?xml version="1.0" encoding="UTF-16"?>');
    expect(await published(utf16le(xml))).toEqual([SENTENCE]);
    expect(await published(utf16be(xml))).toEqual([SENTENCE]);
  });

  it("UTF-8 with a byte-order mark is read as it always was, the mark kept out of the text", async () => {
    const bytes = new Uint8Array(Buffer.from(`﻿${paragraph('<?xml version="1.0" encoding="UTF-8"?>')}`, "utf8"));
    expect(await published(bytes)).toEqual([SENTENCE]);
  });

  it("a part declaring Windows-1252 and written in it is read as declared", async () => {
    const xml = paragraph('<?xml version="1.0" encoding="windows-1252"?>').replace(
      SENTENCE,
      "Voilà le café, naïve — €",
    );
    const bytes = new Uint8Array(Buffer.from(xml.replace("—", "\u0097").replace("€", "\u0080"), "latin1"));
    expect(await published(bytes)).toEqual(["Voilà le café, naïve — €"]);
  });

  // Product owner, 26/10/2026, « préférer l'UTF-8 »: a part declaring Windows-1252 whose bytes are ALSO legal UTF-8 holding
  // accents is read as UTF-8 (the declaration lies); the declared name still reaches the tokenizer's agreement check.
  it("a part declaring Windows-1252 over UTF-8 accented bytes is read as UTF-8: the sentence is published", async () => {
    const bytes = new Uint8Array(Buffer.from(paragraph('<?xml version="1.0" encoding="windows-1252"?>'), "utf8"));
    expect(await published(bytes)).toEqual([SENTENCE]);
  });

  it("the same with a UTF-8 byte-order mark: the mark is stronger evidence of UTF-8, the sentence is published", async () => {
    const bytes = new Uint8Array(Buffer.from(`﻿${paragraph('<?xml version="1.0" encoding="windows-1252"?>')}`, "utf8"));
    expect(await published(bytes)).toEqual([SENTENCE]);
  });

  it("a UTF-8 mark over bytes that are NOT legal UTF-8, declaring Windows-1252, is still refused as damaged", async () => {
    const xml = paragraph('<?xml version="1.0" encoding="windows-1252"?>').replace(SENTENCE, "café");
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from(xml, "latin1")]);
    await expect(readDocx(await withDocument(bytes))).rejects.toMatchObject(DAMAGED);
  });
});

describe("review 27/09 night, finding 11 — an anchor opening a note is kept", () => {
  it("a bookmark at the head of a footnote keeps its anchor, and the note's leading space still goes", async () => {
    const footnotes =
      '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:bookmarkStart w:id="0" w:name="notebm"/>' +
      '<w:r><w:t xml:space="preserve"> Note text</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p></w:footnote>';
    const body =
      '<w:p><w:hyperlink w:anchor="notebm"><w:r><w:t>see note</w:t></w:r></w:hyperlink><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body, footnotes }));
    expect(doc.blocks).toEqual([
      {
        kind: "paragraph",
        inlines: [
          { text: "see note", href: "#notebm" },
          { text: "[1]", marks: ["sup"] },
        ],
      },
      {
        kind: "footnote",
        id: "1",
        blocks: [{ kind: "paragraph", inlines: [{ text: "", anchor: "notebm" }, { text: "Note text" }] }],
      },
    ]);
  });
});

describe("review 27/09 night, finding 12 — mc:AlternateContent takes the first Choice whose Requires the reader understands", () => {
  it("an unknown namespace sends the reader to the next Choice, then to the Fallback; a known one under any prefix is read", async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const CANVAS = 'xmlns:zzz="urn:example:unknown"';
    const SHAPE = 'xmlns:shape="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
    const alt = (inner: string): string =>
      `<w:p><w:r><mc:AlternateContent ${MC} ${CANVAS} ${SHAPE}>${inner}</mc:AlternateContent></w:r></w:p>`;
    const body =
      alt(
        '<mc:Choice Requires="zzz"><w:t>unread canvas</w:t></mc:Choice><mc:Fallback><w:t>fallback words</w:t></mc:Fallback>',
      ) +
      alt(
        '<mc:Choice Requires="zzz"><w:t>unread</w:t></mc:Choice><mc:Choice Requires="shape"><w:t>second choice</w:t></mc:Choice>' +
          "<mc:Fallback><w:t>no</w:t></mc:Fallback>",
      ) +
      alt(
        '<mc:Choice Requires="shape zzz"><w:t>half known</w:t></mc:Choice><mc:Fallback><w:t>fallback again</w:t></mc:Fallback>',
      ) +
      alt('<mc:Choice Requires="unbound"><w:t>unbound prefix</w:t></mc:Choice>');
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["fallback words", "second choice", "fallback again"]);
  });

  it("a Choice with NO Requires at all is never taken: ECMA-376 Part 3 makes the attribute required (arbitrage 28/09)", async () => {
    // The written limit, decided on the norm rather than a corpus probe
    // (28/09/2026): a Requires-less Choice is a non-conforming
    // file, and the Fallback — the only variant whose meaning is certain —
    // is what such a file gets. Failure value: "no requires" published.
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const body =
      `<w:p><w:r><mc:AlternateContent ${MC}>` +
      "<mc:Choice><w:t>no requires</w:t></mc:Choice><mc:Fallback><w:t>fallback wins</w:t></mc:Fallback>" +
      "</mc:AlternateContent></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["fallback wins"]);
  });
});

// M16 docx (02/10/2026): ECMA-376 Part 3 (Markup Compatibility), AlternateContent
// — the first Choice whose Requires namespaces the consumer ALL understands,
// else the Fallback. Word understands more than the ten namespaces this reader
// walks (w14, wp14, wpc…): a Choice requiring them is what Word DISPLAYS, so
// its Fallback is not published (never both branches).
describe("M16 docx (02/10) — a Choice Word understands is the branch published, never its Fallback", () => {
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const WP14 = 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"';
  const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const alt = (inner: string): string =>
    `<w:p><w:r><mc:AlternateContent ${MC} ${WP14} ${W14}>${inner}</mc:AlternateContent></w:r></w:p>`;

  it("Requires wp14, w14 and wp14 w14 are Word's: their Choice is read, the Fallback is not", async () => {
    const body =
      alt(
        '<mc:Choice Requires="wp14"><w:t>wp14 choice</w:t></mc:Choice><mc:Fallback><w:t>fallback one</w:t></mc:Fallback>',
      ) +
      alt(
        '<mc:Choice Requires="w14"><w:t>w14 choice</w:t></mc:Choice><mc:Fallback><w:t>fallback two</w:t></mc:Fallback>',
      ) +
      alt(
        '<mc:Choice Requires="wp14 w14"><w:t>both choice</w:t></mc:Choice><mc:Fallback><w:t>fallback three</w:t></mc:Fallback>',
      );
    const { doc } = await readDocx(await synthetic({ body }));
    // Failure value: the three fallbacks (the 10-prefix list sent them to the Fallback).
    expect(plainTexts(doc.blocks)).toEqual(["wp14 choice", "w14 choice", "both choice"]);
  });

  it("a Choice mixing a Word namespace with a foreign one is not understood: the Fallback, alone", async () => {
    const body = alt(
      '<mc:Choice Requires="wp14 zzz" xmlns:zzz="urn:example:unknown"><w:t>half known</w:t></mc:Choice><mc:Fallback><w:t>fallback words</w:t></mc:Fallback>',
    );
    const { doc } = await readDocx(await synthetic({ body }));
    expect(plainTexts(doc.blocks)).toEqual(["fallback words"]);
  });
});

// M14 (02/10/2026): a floating drawing anchored RELATIVE TO THE PAGE whose
// box lies entirely outside the page is not displayed by Word — it is not
// published. The page is the largest `w:pgSz` of the document (a box off the
// largest page is off every page); any other anchor frame, a partly visible
// box or a rotated one is kept (bias: only what is PROVABLY off the page is
// dropped). Crafted package (`synthetic`), the rule is under test.
describe("M14 (02/10) — a drawing anchored entirely off the page is not published", () => {
  const A4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  const NS =
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  // A4 in EMU: 7 560 310 × 10 692 130 (11906 × 16838 twips × 635).
  const box = (
    text: string,
    h: { from: string; offset: number },
    v: { from: string; offset: number },
    cx = 1_000_000,
    cy = 500_000,
    rot = "",
  ): string =>
    `<w:p><w:r><w:t>host ${text}</w:t></w:r><w:r><w:drawing ${NS}><wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="${h.from}"><wp:posOffset>${h.offset}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${v.from}"><wp:posOffset>${v.offset}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="b"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
    `<wps:spPr><a:xfrm${rot}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
  const page = (x: number, y: number): [{ from: string; offset: number }, { from: string; offset: number }] => [
    { from: "page", offset: x },
    { from: "page", offset: y },
  ];
  const published = async (body: string): Promise<string[]> => {
    const { doc } = await readDocx(await synthetic({ body: body + A4 }));
    return plainTexts(doc.blocks).filter((t) => !t.startsWith("host"));
  };

  it("fully past the right edge, the bottom edge, the left edge or above the top: not published", async () => {
    const body =
      box("right off", ...page(8_000_000, 1_000_000)) +
      box("below off", ...page(1_000_000, 11_000_000)) +
      box("left off", ...page(-1_500_000, 1_000_000)) +
      box("above off", ...page(1_000_000, -600_000));
    // Failure value: the four texts (w:pgSz is never read).
    expect(await published(body)).toEqual([]);
  });

  it("on the page, or partly on it: published", async () => {
    const body =
      box("inside", ...page(1_000_000, 1_000_000)) +
      box("straddles right", ...page(7_000_000, 1_000_000)) +
      box("straddles left", ...page(-500_000, 1_000_000));
    expect(await published(body)).toEqual(["inside", "straddles right", "straddles left"]);
  });

  it("an anchor relative to anything but the page cannot be resolved: published, however far its offset", async () => {
    const body =
      box("column far", { from: "column", offset: 90_000_000 }, { from: "paragraph", offset: 90_000_000 }) +
      box("margin far", { from: "margin", offset: 90_000_000 }, { from: "page", offset: 1_000_000 });
    expect(await published(body)).toEqual(["column far", "margin far"]);
  });

  it("a rotated box is kept: its turned outline is not the declared one", async () => {
    expect(await published(box("turned", ...page(8_000_000, 1_000_000), 1_000_000, 500_000, ' rot="5400000"'))).toEqual(
      ["turned"],
    );
  });

  it("the page is the LARGEST of the document's sections: a box off a small page but on a bigger one is kept", async () => {
    const small =
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="8000" w:h="11000"/></w:sectPr></w:pPr><w:r><w:t>host s1</w:t></w:r></w:p>';
    const big = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>';
    // x = 9 000 000 EMU: past 8 000 twips (5 080 000) yet inside 16 838 twips (10 692 130).
    const { doc } = await readDocx(
      await synthetic({ body: small + box("on the big page", ...page(9_000_000, 1_000_000)) + big }),
    );
    expect(plainTexts(doc.blocks).filter((t) => !t.startsWith("host"))).toEqual(["on the big page"]);
  });
});

// The recensement of 28/09/2026 asked for this guard: `resolve`'s `fix`
// patches an INVENTORY of block kinds, and an image carried anywhere else
// (an Inline.image, a future block) would cross with its sentinel as
// sha256 — an attachment named from `pending:`, in silence. The guard
// walks the WHOLE tree, and it is OUR defect when it bites.
describe("arbitrage 28/09 (8e) — no pending: ref survives resolve, wherever it hides", () => {
  it("bites on a sentinel in an inline image (a spot fix does not patch); passes a resolved doc", () => {
    const hidden: IRBlock = {
      kind: "paragraph",
      inlines: [{ text: "4", image: { sha256: `${PENDING_REF}0`, mimeType: "image/png" } }],
    };
    expect(() => assertNoPendingRef({ blocks: [hidden] })).toThrow("docx-reader: a pending image ref survived resolve");
    expect(() => assertNoPendingRef({ blocks: [para("sain")] })).not.toThrow();
  });

  it("bites on a sentinel in a chart image nested in a columns cell — carriers no spot fix patches", () => {
    // ONE tree exercising BOTH carriers: reaching the sentinel requires the
    // walk to enter a `columns` block's cells AND a `chart` block's image.
    const hidden: IRBlock = {
      kind: "columns",
      cells: [
        [para("left column, clean")],
        [{ kind: "chart", image: { sha256: `${PENDING_REF}1`, mimeType: "image/png" }, data: { rows: [["1"]] } }],
      ],
    };
    expect(() => assertNoPendingRef({ blocks: [hidden] })).toThrow("docx-reader: a pending image ref survived resolve");
  });
});

describe("review 27/09 night, finding 13 — a cell's gridSpan is clamped to Word's 63 columns", () => {
  it("gridSpan 1 000 000 is written as colspan 63", async () => {
    const body =
      '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="1000000"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc } = await readDocx(await synthetic({ body }));
    expect(doc.blocks).toEqual([{ kind: "table", rows: [{ cells: [{ blocks: [para("wide")], colspan: 63 }] }] }]);
  });
});

// Review of 01/10/2026 (security): the connector order of a drawing cluster
// ran on any number of elements — the pptx cap (MAX_SLIDE_ELEMENTS) was not
// applied per cluster, and separate drawings of one cluster bypassed the
// per-drawing bound. Past the cap the cluster keeps today's order.
describe("review 01/10 — a drawing cluster past MAX_SLIDE_ELEMENTS keeps document order", () => {
  const NS =
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
  const anchored = (x: number, cx: number, cy: number, id: number, shape: string): string =>
    `<w:r><w:drawing><wp:anchor ${NS}><wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="s${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${shape}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const textBox = (k: number): string =>
    anchored(
      k * 2_000_000,
      1_000_000,
      400_000,
      k + 1,
      '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
        `<wps:txbx><w:txbxContent><w:p><w:r><w:t>t${k}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>`,
    );
  // From box k+1's left edge to box k's right edge, an arrow head at the end.
  const connector = (k: number, n: number): string =>
    anchored(
      k * 2_000_000 + 1_000_000,
      1_000_000,
      0,
      n + k + 1,
      '<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm flipH="1"><a:off x="0" y="200000"/><a:ext cx="1000000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
        '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></wps:spPr></wps:wsp>',
    );
  const orderOf = async (n: number): Promise<string[]> => {
    const boxes = Array.from({ length: n }, (_, k) => textBox(k)).join("");
    const lines = Array.from({ length: n - 1 }, (_, k) => connector(k, n)).join("");
    // The element cap is the PRODUCT's, asked for (09/11/2026): the reader holds none of its own, so without this the
    // 519-element cluster below would be drawn and its words would come back reversed.
    const { doc } = await readDocx(
      await synthetic({ body: `<w:p>${boxes}${lines}</w:p>` }),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    return plainTexts(doc.blocks).filter((t) => /^t\d+$/.test(t));
  };
  it("control: 100 boxes and 99 connectors are put in the order of the connectors (the arrows run right to left)", async () => {
    const got = await orderOf(100);
    // Failure value of the cap test below: this reversed order.
    expect(got).toEqual(Array.from({ length: 100 }, (_, k) => `t${99 - k}`));
  });
  it("the cluster cap is a parameter (ReaderLimits.slideElements): 10 boxes and 9 connectors keep document order once it is lowered to 18", async () => {
    const boxes = Array.from({ length: 10 }, (_, k) => textBox(k)).join("");
    const lines = Array.from({ length: 9 }, (_, k) => connector(k, 10)).join("");
    const bytes = await synthetic({ body: `<w:p>${boxes}${lines}</w:p>` });
    const order = async (limits?: { slideElements: number }): Promise<string[]> =>
      plainTexts((await readDocx(bytes, undefined, undefined, limits)).doc.blocks).filter((t) => /^t\d+$/.test(t));
    // FAILS WITH (limits ignored): expected [ 't9', 't8', … ] to deeply equal [ 't0', 't1', … ].
    expect(await order({ slideElements: 18 })).toEqual(Array.from({ length: 10 }, (_, k) => `t${k}`));
    expect(await order({ slideElements: 19 })).toEqual(Array.from({ length: 10 }, (_, k) => `t${9 - k}`));
    expect(await order()).toEqual(Array.from({ length: 10 }, (_, k) => `t${9 - k}`));
  });
  it("260 boxes and 259 connectors (519 elements) stay in document order, all words published", async () => {
    const got = await orderOf(260);
    expect(got).toEqual(Array.from({ length: 260 }, (_, k) => `t${k}`));
  }, 60_000);
});

// Audit of 07/10/2026: one damaged image must not refuse the whole document — the text is published, the damaged image
// is not invented (a presentation already degrades a damaged part).
describe("readDocx — a damaged image refuses nothing else", () => {
  it("🔴 an image whose stream is broken leaves the document published, its text included", async () => {
    const real = new Uint8Array(readFileSync(join(here, "../../../fixtures/docx/02-Procedure-Support-IT.docx")));
    const entries = await entriesOf(real);
    const media = entries.find((e) => e.name.startsWith("word/media/") && e.bytes.length > 0)!.name;
    const bytes = handZip(entries.map((e) => (e.name === media ? { name: media, brokenStream: true as const } : e)));
    const { doc } = await readDocx(bytes);
    expect(doc.blocks.length).toBeGreaterThan(0);
    expect(JSON.stringify(doc.blocks)).not.toContain(PENDING_REF);
  });

  // 07/10/2026: the loss is said, with the notice the app already uses for an image that did not come.
  it("🔴 the damaged image is said: one image-missing notice", async () => {
    const real = new Uint8Array(readFileSync(join(here, "../../../fixtures/docx/02-Procedure-Support-IT.docx")));
    const entries = await entriesOf(real);
    const media = entries.find((e) => e.name.startsWith("word/media/") && e.bytes.length > 0)!.name;
    const { notices } = await readDocx(
      handZip(entries.map((e) => (e.name === media ? { name: media, brokenStream: true as const } : e))),
    );
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });
});

// Exhaustive audit of 07/10/2026: the 1×1-pixel probe of a Fallback picture (index.ts ≈ 253) read its part without the
// damaged-part rule — one broken Fallback image still refused the whole document.
describe("readDocx — a damaged Fallback picture refuses nothing else", () => {
  it("🔴 the document is published, its text included", async () => {
    const body =
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice>' +
      '<mc:Fallback><w:drawing><a:blip r:embed="rIdFb"/></w:drawing></mc:Fallback></mc:AlternateContent></w:r>' +
      "<w:r><w:t>Bonjour</w:t></w:r></w:p>";
    const base = await synthetic({
      body,
      rels: [["rIdFb", "image", "media/fb.png"]],
      files: { "word/media/fb.png": new Uint8Array([1, 2, 3]) },
    });
    const entries = await entriesOf(base);
    const bytes = handZip(
      entries.map((e) => (e.name === "word/media/fb.png" ? { name: e.name, brokenStream: true as const } : e)),
    );
    const { doc } = await readDocx(bytes);
    expect(JSON.stringify(doc.blocks)).toContain("Bonjour");
  });

  it("🔴 big Fallback pictures are never inflated by the 1×1 probe: they spend none of the package budget (15/10/2026)", async () => {
    const fbIds = ["rIdF1", "rIdF2", "rIdF3", "rIdF4", "rIdF5"];
    const fallbacks = fbIds
      .map((id) => `<mc:Fallback><w:drawing><a:blip r:embed="${id}"/></w:drawing></mc:Fallback>`)
      .join("");
    const body =
      `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice>${fallbacks}</mc:AlternateContent></w:r>` +
      `<w:r><w:t>Bonjour</w:t></w:r></w:p><w:p>${pictureRun("rIdA")}</w:p>`;
    const base = await synthetic({
      body,
      rels: [...fbIds, "rIdA"].map((id, i) => [id, "image", `media/p${i}.png`] as [string, string, string]),
    });
    const entries = (await entriesOf(base)).filter((e) => !e.name.startsWith("word/media/"));
    const heavy = new Uint8Array(49_900_000);
    const small = new Uint8Array(10_000_000).fill(2);
    // 5 × 49.9 MB probed + 10 MB read for the body picture passes the 256 MB package budget (the probe is what spends it).
    const bytes = handZip([
      ...entries,
      ...fbIds.map((_, i) => ({ name: `word/media/p${i}.png`, bytes: heavy })),
      { name: "word/media/p5.png", bytes: small },
    ]);
    const read = await readDocx(bytes).then(
      (r) => ({ images: r.images.length, notices: r.notices ?? [] }),
      (e: unknown) => ({ error: String(e).slice(0, 80) }),
    );
    expect(read).toEqual({ images: 1, notices: [] });
  });

  it("🔴 an oversized Fallback picture (past the image budget) refuses nothing either (review 14/10/2026)", async () => {
    const body =
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice>' +
      '<mc:Fallback><w:drawing><a:blip r:embed="rIdFb"/></w:drawing></mc:Fallback></mc:AlternateContent></w:r>' +
      "<w:r><w:t>Bonjour</w:t></w:r></w:p>";
    const base = await synthetic({
      body,
      rels: [["rIdFb", "image", "media/fb.png"]],
      files: { "word/media/fb.png": new Uint8Array([1, 2, 3]) },
    });
    const entries = await entriesOf(base);
    const bytes = handZip(
      entries.map((e) => (e.name === "word/media/fb.png" ? { name: e.name, bytes: new Uint8Array(50_100_000) } : e)),
    );
    const { doc } = await readDocx(bytes);
    expect(JSON.stringify(doc.blocks)).toContain("Bonjour");
  });
});

// ---------------------------------------------------------------------------
// Lot 10C1 (15/10/2026): the steps `readDocx` strings together, pinned before it was cut into one function per step. Every
// expectation below was seen green on the uncut function.
describe("readDocx — the steps it strings together", () => {
  it("a numbering that only links to a numbering style takes that style's numId", async () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="LinkedList"/></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>';
    const styles =
      '<w:style w:type="numbering" w:styleId="LinkedList"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const body =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body, styles, numbering }));
    expect(doc.blocks).toEqual([
      { kind: "list", ordered: true, items: [[{ kind: "paragraph", inlines: [{ text: "first" }] }]] },
    ]);
  });

  it("a numbering style link reaches a numId declared FURTHER DOWN: `w:num` is read in two passes", async () => {
    // The linked instance (numId 2) comes after the instance that links to it (numId 1): one pass over `w:num` would not
    // know numId 2 yet, and the paragraph would lose its number. Failure value: a plain paragraph.
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="LinkedList"/></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';
    const styles =
      '<w:style w:type="numbering" w:styleId="LinkedList"><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>';
    const body =
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>';
    const { doc } = await readDocx(await synthetic({ body, styles, numbering }));
    expect(doc.blocks).toEqual([
      { kind: "list", ordered: true, items: [[{ kind: "paragraph", inlines: [{ text: "first" }] }]] },
    ]);
  });

  it("two calls to one note give one footnote and the same number twice", async () => {
    const call = '<w:r><w:footnoteReference w:id="1"/></w:r>';
    const footnotes = '<w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>';
    const { doc } = await readDocx(
      await synthetic({ body: `<w:p><w:r><w:t>a</w:t></w:r>${call}<w:r><w:t>b</w:t></w:r>${call}</w:p>`, footnotes }),
    );
    expect(doc.blocks).toEqual([
      {
        kind: "paragraph",
        inlines: [{ text: "a" }, { text: "[1]", marks: ["sup"] }, { text: "b" }, { text: "[1]", marks: ["sup"] }],
      },
      { kind: "footnote", id: "1", blocks: [{ kind: "paragraph", inlines: [{ text: "note" }] }] },
    ]);
  });

  it("the SmartArt words of a footnote and of an endnote are gathered, the footnote's first", async () => {
    const cache = await (
      await JSZip.loadAsync(readFileSync(join(here, "../../../fixtures/diagrammes-reel/06-fr-la-veille.pptx")))
    )
      .file("ppt/diagrams/drawing1.xml")!
      .async("string");
    const smart =
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const body = '<w:p><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>';
    const base = await synthetic({
      body,
      footnotes: `<w:footnote w:id="1">${smart}</w:footnote>`,
      endnotes: `<w:endnote w:id="1">${smart}</w:endnote>`,
    });
    const noteRels = (target: string): string =>
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="${target}"/></Relationships>`;
    const bytes = handZip([
      ...(await entriesOf(base)),
      { name: "word/_rels/footnotes.xml.rels", bytes: new TextEncoder().encode(noteRels("diagrams/footnote.xml")) },
      { name: "word/_rels/endnotes.xml.rels", bytes: new TextEncoder().encode(noteRels("diagrams/endnote.xml")) },
      { name: "word/diagrams/footnote.xml", bytes: new TextEncoder().encode(cache) },
      {
        name: "word/diagrams/endnote.xml",
        bytes: new TextEncoder().encode(cache.replace(/<a:t>Recherche/g, "<a:t>Zzz")),
      },
    ]);
    const { doc } = await readDocx(bytes);
    const diagram = doc.blocks.find((b) => b.kind === "diagramText");
    if (diagram?.kind !== "diagramText") throw new Error("the SmartArt words are missing");
    expect(diagram.sections.map((s) => s.inlines[0]!.text.slice(0, 8))).toEqual(["Recherch", "Zzz d’in"]);
  });

  it("a Fallback picture whose relationship does not resolve is no pixel and refuses nothing", async () => {
    const body =
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice>' +
      '<mc:Fallback><w:drawing><a:blip r:embed="rIdGone"/></w:drawing></mc:Fallback></mc:AlternateContent></w:r>' +
      "<w:r><w:t>Bonjour</w:t></w:r></w:p>";
    const { doc } = await readDocx(await synthetic({ body }));
    expect(JSON.stringify(doc.blocks)).toContain("Bonjour");
  });

  it("a Fallback picture whose header declares less than it inflates to is not read past the probe's ceiling, and refuses nothing", async () => {
    const body =
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>choice</w:t></mc:Choice>' +
      '<mc:Fallback><w:drawing><a:blip r:embed="rIdFb"/></w:drawing></mc:Fallback></mc:AlternateContent></w:r>' +
      "<w:r><w:t>Bonjour</w:t></w:r></w:p>";
    const base = await synthetic({
      body,
      rels: [["rIdFb", "image", "media/fb.png"]],
      files: { "word/media/fb.png": new Uint8Array([1, 2, 3]) },
    });
    const entries = await entriesOf(base);
    const bytes = handZip(
      entries.map((e) =>
        e.name === "word/media/fb.png" ? { name: e.name, bytes: new Uint8Array(100_000), declaredSize: 100 } : e,
      ),
    );
    const { doc } = await readDocx(bytes);
    expect(JSON.stringify(doc.blocks)).toContain("Bonjour");
  });
});
