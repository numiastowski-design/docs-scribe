import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { readDocx } from "./index";

// Decision of 08/10/2026: damaged STYLES are not a reason to refuse the document. It is published with default
// formatting AND the loss is said in one { kind: "styles-defaulted" } notice (no other field, one entry per document).
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/docx-reader/docx-reader.damaged-styles.test.ts
vi.setConfig({ testTimeout: 30_000 });

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const STYLES =
  `<?xml version="1.0"?><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>` +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>';
const NUMBERING =
  `<?xml version="1.0"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>` +
  '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const BROKEN_STYLES = '<?xml version="1.0"?><w:styles <<< not xml';
const BROKEN_NUMBERING = '<?xml version="1.0"?><w:numbering <<< not xml';

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title kept</w:t></w:r></w:p>' +
  "<w:p><w:r><w:t>Plain kept</w:t></w:r></w:p>" +
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item one kept</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item two kept</w:t></w:r></w:p>';
const TEXTS = ["Title kept", "Plain kept", "Item one kept", "Item two kept"];

const build = async (parts: { styles?: string; numbering?: string }): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const rels: string[] = [];
  if (parts.styles !== undefined) {
    zip.file("word/styles.xml", parts.styles);
    rels.push(`<Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>`);
  }
  if (parts.numbering !== undefined) {
    zip.file("word/numbering.xml", parts.numbering);
    rels.push(`<Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/>`);
  }
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>${BODY}</w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

/** Breaks the DEFLATE stream of one entry behind its intact local header (same technique as docx-reader.test.ts). */
const corruptStream = (bytes: Uint8Array, entry: string): Uint8Array => {
  const out = new Uint8Array(bytes);
  const name = new TextEncoder().encode(entry);
  const view = new DataView(out.buffer, out.byteOffset);
  let nameAt = -1;
  for (let i = 30; i < out.length - name.length; i++) {
    if (!name.every((b, j) => out[i + j] === b)) continue;
    if (view.getUint32(i - 30, true) === 0x04034b50 && view.getUint16(i - 30 + 26, true) === name.length) {
      nameAt = i;
      break;
    }
  }
  expect(nameAt).toBeGreaterThan(0);
  const dataStart = nameAt + name.length + view.getUint16(nameAt - 30 + 28, true);
  for (let i = 0; i < 8; i++) out[dataStart + i] = (out[dataStart + i] ?? 0) ^ 0xff;
  return out;
};

const kinds = (blocks: readonly IRBlock[]): string[] => blocks.map((b) => b.kind);
const defaulted = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "styles-defaulted");
const expectAllTexts = (blocks: readonly IRBlock[]): void => {
  const json = JSON.stringify(blocks);
  for (const t of TEXTS) expect(json).toContain(t);
};

describe("S1 docx — styles.xml unreadable", () => {
  it("non-regression: readable styles and numbering give a heading, a list, and no styles-defaulted notice", async () => {
    const { doc, notices } = await readDocx(await build({ styles: STYLES, numbering: NUMBERING }));
    // Proves the fixture exercises styles: the damaged cases below must differ from this one (no heading).
    expect(kinds(doc.blocks)).toContain("heading");
    expect(kinds(doc.blocks)).toContain("list");
    expectAllTexts(doc.blocks);
    expect(defaulted(notices)).toEqual([]);
  });

  it("🔴 XML that does not parse: published, texts intact, no heading, one styles-defaulted notice with no other field", async () => {
    // Failure value (before): readDocx rejects with EXTRACT_FAILED.
    const { doc, notices } = await readDocx(await build({ styles: BROKEN_STYLES, numbering: NUMBERING }));
    expectAllTexts(doc.blocks);
    expect(kinds(doc.blocks)).not.toContain("heading");
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });

  it("🔴 corrupt DEFLATE stream: published, texts intact, no heading, one styles-defaulted notice", async () => {
    const bytes = corruptStream(await build({ styles: STYLES, numbering: NUMBERING }), "word/styles.xml");
    const { doc, notices } = await readDocx(bytes);
    expectAllTexts(doc.blocks);
    expect(kinds(doc.blocks)).not.toContain("heading");
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });
});

describe("S1 docx — numbering.xml unreadable", () => {
  it("🔴 XML that does not parse: published, texts intact, one styles-defaulted notice; readable styles keep the heading", async () => {
    const { doc, notices } = await readDocx(await build({ styles: STYLES, numbering: BROKEN_NUMBERING }));
    expectAllTexts(doc.blocks);
    expect(kinds(doc.blocks)).toContain("heading");
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });

  it("🔴 corrupt DEFLATE stream: published, texts intact, one styles-defaulted notice", async () => {
    const bytes = corruptStream(await build({ styles: STYLES, numbering: NUMBERING }), "word/numbering.xml");
    const { doc, notices } = await readDocx(bytes);
    expectAllTexts(doc.blocks);
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });
});

describe("S1 docx — both parts unreadable", () => {
  it("🔴 ONE styles-defaulted entry, not two", async () => {
    const { doc, notices } = await readDocx(await build({ styles: BROKEN_STYLES, numbering: BROKEN_NUMBERING }));
    expectAllTexts(doc.blocks);
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });
});

describe("docx — what still refuses", () => {
  it("non-regression: a corrupt word/document.xml still fails with EXTRACT_FAILED", async () => {
    const bytes = corruptStream(await build({ styles: STYLES, numbering: NUMBERING }), "word/document.xml");
    await expect(readDocx(bytes)).rejects.toMatchObject({ code: "EXTRACT_FAILED" });
  });
});
