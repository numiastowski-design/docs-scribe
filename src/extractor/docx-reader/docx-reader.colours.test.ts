/** What the Word reader keeps of colour. The rule these tests hold: the reader hands on what the file states, never a
 *  colour it worked out. A theme colour (`w:themeColor`) names a slot of the document's palette, not a colour: as long
 *  as the theme part is not read, nothing is handed on for it. That is a limit of scope, not a principle. */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { PAGE_SIZED_LIMITS } from "../page-caps.test.support";
import { extractDocx } from "../docx";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}"`;

const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  zip.file("word/_rels/document.xml.rels", RELS(""));
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

/** The first inline of the first block, as the reader hands it on. */
const firstInline = async (runProperties: string, text = "règle"): Promise<Record<string, unknown>> => {
  const body = `<w:p><w:r><w:rPr>${runProperties}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const { doc } = await extractDocx(await docxOf(body), PAGE_SIZED_LIMITS);
  const block = doc.blocks[0] as { inlines?: Record<string, unknown>[] } | undefined;
  return block?.inlines?.[0] ?? {};
};

describe("Word, la couleur que le fichier énonce", () => {
  it("🔴 une couleur de texte en hex est gardée", async () => {
    expect(await firstInline('<w:color w:val="FF0000"/>')).toMatchObject({ text: "règle", color: "#ff0000" });
  });

  it("🔴 un surlignage nommé est gardé sous son nom, en minuscules", async () => {
    expect(await firstInline('<w:highlight w:val="yellow"/>')).toMatchObject({ background: "yellow" });
    expect(await firstInline('<w:highlight w:val="darkBlue"/>')).toMatchObject({ background: "darkblue" });
  });

  it("🔴 un fond de trame en hex est gardé quand il n'y a pas de surlignage", async () => {
    expect(await firstInline('<w:shd w:val="clear" w:fill="FFFF00"/>')).toMatchObject({ background: "#ffff00" });
  });

  it("le surlignage l'emporte sur la trame : c'est la marque qu'un lecteur du document voit", async () => {
    expect(await firstInline('<w:shd w:val="clear" w:fill="FFFF00"/><w:highlight w:val="green"/>')).toMatchObject({
      background: "green",
    });
  });

  it("🔴 une couleur de THÈME ne publie rien : elle nomme une case, pas une couleur", async () => {
    // Failure value: `color: "#4472c4"`.
    expect(await firstInline('<w:color w:val="4472C4" w:themeColor="accent1"/>')).not.toHaveProperty("color");
    expect(await firstInline('<w:shd w:val="clear" w:fill="D9E2F3" w:themeFill="accent1"/>')).not.toHaveProperty(
      "background",
    );
  });

  it("🔴 « auto » et « none » énoncent l'absence de couleur, et ne publient rien", async () => {
    expect(await firstInline('<w:color w:val="auto"/>')).not.toHaveProperty("color");
    expect(await firstInline('<w:highlight w:val="none"/>')).not.toHaveProperty("background");
  });

  it("un run sans couleur n'en porte aucune, et son texte reste intact", async () => {
    const inline = await firstInline("", "corps");
    expect(inline).toMatchObject({ text: "corps" });
    expect(inline).not.toHaveProperty("color");
    expect(inline).not.toHaveProperty("background");
  });

  it("la couleur et le gras voyagent ensemble, sans que l'une chasse l'autre", async () => {
    expect(await firstInline('<w:b/><w:color w:val="FF0000"/>')).toMatchObject({ marks: ["strong"], color: "#ff0000" });
  });
});
