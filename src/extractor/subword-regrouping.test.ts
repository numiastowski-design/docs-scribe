/**
 * Sub-word regrouping: a hidden prefix glued to a visible word must never make another word.
 *
 * The attack (« When File Ingestion Misleads Language Models: Measuring and Mitigating Divergences in Office
 * Documents », arXiv:2611.05806): a hidden `un` is put before a visible `authorized`; an extractor that reads the
 * hidden run and joins the fragments publishes « unauthorized », which reverses the sentence. The result is an
 * ordinary dictionary word, so no text filter can spot it afterwards.
 *
 * Why the reader is immune: a `w:vanish` run is never published, so the prefix is gone before the join.
 *
 * Both directions are tested. The first test proves the hidden prefix does not pass. The second proves that without
 * the mask the runs do join into one word: without it the first could pass simply because nothing is ever joined.
 */

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "./docx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}"`;
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

/** The smallest .docx that carries a body. */
const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  zip.file("word/_rels/document.xml.rels", RELS(""));
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}${a4}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const run = (text: string, rPr = ""): string => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const VANISH = "<w:rPr><w:vanish/></w:rPr>";

/** The text of the page, inlines joined: in a `JSON.stringify` of the blocks two joined pieces would stay two
 *  neighbouring strings, which is what this test must tell apart. */
const plainText = (blocks: readonly IRBlock[]): string => {
  const out: string[] = [];
  const visit = (list: readonly IRBlock[]): void => {
    for (const block of list) {
      if (block.kind === "paragraph" || block.kind === "heading")
        out.push(block.inlines.map((inline) => inline.text).join(""));
      else if (block.kind === "table") block.rows.forEach((row) => row.cells.forEach((cell) => visit(cell.blocks)));
    }
  };
  visit(blocks);
  return out.join("\n");
};

/** « Access is <prefix>authorized by default. », the prefix carrying the mask or not. */
const sentence = (prefix: string, rPr: string): string =>
  para(run("Access is "), run(prefix, rPr), run("authorized"), run(" by default."));

const publishedText = async (body: string): Promise<string> =>
  plainText((await extractDocx(await docxOf(body))).doc.blocks);

describe("regroupement en sous-mots (arXiv:2611.05806) — un préfixe caché ne fabrique pas un autre mot", () => {
  // The five prefixes the paper cites.
  for (const prefix of ["un", "non", "dis", "in", "im"]) {
    it(`🔴 « ${prefix} » caché devant « authorized » : la page dit authorized, jamais ${prefix}authorized`, async () => {
      const text = await publishedText(sentence(prefix, VANISH));
      expect(text).not.toContain(`${prefix}authorized`);
      expect(text).toContain("Access is authorized by default.");
    });
  }

  it("🔴 SANS le masque, les runs se recollent bien en un seul mot — sans quoi le test ci-dessus ne garderait rien", async () => {
    // Failure value: « Access is un authorized », which would make the immunity above an accident of the assembly.
    const text = await publishedText(sentence("un", ""));
    expect(text).toContain("Access is unauthorized by default.");
  });

  it("un préfixe caché À L'INTÉRIEUR d'un mot ne coupe pas le mot visible en deux", async () => {
    // The other edge of the same mechanism: the mask falls between two visible fragments, which must join.
    const text = await publishedText(para(run("auth"), run("XX", VANISH), run("orized")));
    expect(text).toContain("authorized");
    expect(text).not.toContain("XX");
  });
});
