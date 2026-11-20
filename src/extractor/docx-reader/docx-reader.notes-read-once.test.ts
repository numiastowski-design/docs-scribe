import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { readDocx } from "./index";

// Review of 01/10/2026 (performance): footnotes.xml and endnotes.xml were
// parsed by the drawing-object pre-read AND again by `noteBlocks`, and every
// inflate counts against PACKAGE_MAX_BYTES (« a part read twice counts
// twice »). Each notes part is now read once, and not at all when the body
// calls no note.
const reads: string[] = [];
vi.mock("../ooxml/package", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../ooxml/package")>();
  return {
    ...actual,
    openPackage: async (...args: Parameters<typeof actual.openPackage>) => {
      const pkg = await actual.openPackage(...args);
      return {
        ...pkg,
        readXml: async (path: string) => {
          reads.push(path);
          return pkg.readXml(path);
        },
      };
    },
  };
});

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const docxWithNotes = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdE" Type="${REL}/endnotes" Target="endnotes.xml"/></Relationships>`,
  );
  zip.file(
    "word/footnotes.xml",
    `<?xml version="1.0"?><w:footnotes ${W_NS}><w:footnote w:id="1"><w:p><w:r><w:t>foot text</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  );
  zip.file(
    "word/endnotes.xml",
    `<?xml version="1.0"?><w:endnotes ${W_NS}><w:endnote w:id="1"><w:p><w:r><w:t>end text</w:t></w:r></w:p></w:endnote></w:endnotes>`,
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const texts = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((b): string[] =>
    b.kind === "paragraph" ? [b.inlines.map((i) => i.text).join("")] : b.kind === "footnote" ? texts(b.blocks) : [],
  );
const readsOf = (path: string): number => reads.filter((p) => p === path).length;

describe("review 01/10 — each notes part is parsed once", () => {
  it("a body that calls a footnote and an endnote reads footnotes.xml and endnotes.xml once each, the notes unchanged", async () => {
    reads.length = 0;
    const body =
      '<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>';
    const { doc } = await readDocx(await docxWithNotes(body));
    // Failure value: 2 reads each (the pre-read, then `noteBlocks`).
    expect([readsOf("word/footnotes.xml"), readsOf("word/endnotes.xml")]).toEqual([1, 1]);
    expect(texts(doc.blocks)).toEqual(["body[1][2]", "foot text", "end text"]);
  });

  it("a body that calls only an endnote (inside a table cell) reads the endnotes part and publishes the note", async () => {
    reads.length = 0;
    const body =
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc } = await readDocx(await docxWithNotes(body));
    expect(readsOf("word/endnotes.xml")).toBeGreaterThan(0);
    expect(JSON.stringify(doc.blocks)).toContain("end text");
  });

  it("a body that calls no note does not read the notes parts at all", async () => {
    reads.length = 0;
    await readDocx(await docxWithNotes("<w:p><w:r><w:t>plain</w:t></w:r></w:p>"));
    expect([readsOf("word/footnotes.xml"), readsOf("word/endnotes.xml")]).toEqual([0, 0]);
  });
});
