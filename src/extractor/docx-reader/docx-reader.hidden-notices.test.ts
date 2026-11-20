import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { encodePng } from "doc-ir/internal";
import { readDocx } from "./index";

// Decision of 08/10/2026: hidden content dropped by the reader is SAID in a DocumentNotice
// { kind: "hidden-dropped", what: "shape" | "tracked-deletion", count } (kind and count only, never the content);
// one entry per document and per `what`, absent when the count is 0. Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/docx-reader/docx-reader.hidden-notices.test.ts
vi.setConfig({ testTimeout: 30_000 });

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:r="${REL}"`;

const dot = (rgb: [number, number, number]): Uint8Array => encodePng(1, 1, "rgb-24", rgb)!;

const picture = (relId: string, id: number, hidden: string): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="${id}" name="p${id}"${hidden}/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const textBox = (words: string, id: number, hidden: string): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="${id}" name="b${id}"${hidden}/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${words}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const para = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const run = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;
const del = (t: string, id: number): string =>
  `<w:del w:id="${id}" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
const emptyDel = (id: number): string =>
  `<w:del w:id="${id}" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r></w:r></w:del>`;
const ins = (t: string, id: number): string =>
  `<w:ins w:id="${id}" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>${t}</w:t></w:r></w:ins>`;

const build = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdRed" Type="${REL}/image" Target="media/red.png"/><Relationship Id="rIdBlue" Type="${REL}/image" Target="media/blue.png"/></Relationships>`,
  );
  zip.file("word/media/red.png", dot([255, 0, 0]));
  zip.file("word/media/blue.png", dot([0, 0, 255]));
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${body}${a4}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const hiddenNotices = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "hidden-dropped");
const HID = ' hidden="1"';

describe("N1 — hidden drawings are counted in a hidden-dropped/shape notice", () => {
  it("🔴 one hidden picture: { hidden-dropped, shape, 1 }", async () => {
    const { notices } = await readDocx(
      await build(para(run("host"), picture("rIdRed", 1, HID), picture("rIdBlue", 2, ""))),
    );
    // Failure value (before): no hidden-dropped notice at all.
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "shape", count: 1 }]);
  });

  it("🔴 three hidden drawings (picture, picture, text box) give ONE entry with count 3; the text box counts 1", async () => {
    const body = para(
      run("host"),
      picture("rIdRed", 1, HID),
      picture("rIdBlue", 2, ' hidden="true"'),
      textBox("secretwords", 3, HID),
    );
    const { notices } = await readDocx(await build(body));
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "shape", count: 3 }]);
  });

  it("🔴 a hidden text box alone counts 1 and the notice never carries its words", async () => {
    const { notices } = await readDocx(await build(para(run("host"), textBox("secretwordsinbox", 1, HID))));
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "shape", count: 1 }]);
    expect(JSON.stringify(notices)).not.toContain("secretwordsinbox");
  });

  it.each([' hidden="0"', ' hidden="false"', ""])(
    "non-regression: docPr%s (visible) gives no hidden-dropped notice",
    async (attr) => {
      const { notices } = await readDocx(
        await build(para(run("host"), picture("rIdRed", 1, attr), textBox("visiblewords", 2, attr))),
      );
      expect(hiddenNotices(notices)).toEqual([]);
    },
  );

  it("non-regression: a plain document has no hidden-dropped notice", async () => {
    const { notices } = await readDocx(await build(para(run("plain"))));
    expect(hiddenNotices(notices)).toEqual([]);
  });
});

describe("N2 — tracked deletions are counted in a hidden-dropped/tracked-deletion notice", () => {
  it("🔴 one w:del with text: { hidden-dropped, tracked-deletion, 1 }, its words stay unpublished and out of the notice", async () => {
    const { doc, notices } = await readDocx(await build(para(run("kept "), del("removedwords", 1))));
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "tracked-deletion", count: 1 }]);
    expect(JSON.stringify(doc.blocks)).not.toContain("removedwords");
    expect(JSON.stringify(notices)).not.toContain("removedwords");
  });

  it("🔴 four w:del with text over two paragraphs: ONE entry with count 4", async () => {
    const body = para(run("a"), del("x1", 1), del("x2", 2)) + para(run("b"), del("x3", 3), del("x4", 4));
    const { notices } = await readDocx(await build(body));
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "tracked-deletion", count: 4 }]);
  });

  it("🔴 an empty w:del is not counted: 2 with text + 1 empty gives count 2", async () => {
    const { notices } = await readDocx(await build(para(run("a"), del("x1", 1), emptyDel(2), del("x2", 3))));
    expect(hiddenNotices(notices)).toEqual([{ kind: "hidden-dropped", what: "tracked-deletion", count: 2 }]);
  });

  it("non-regression: only an empty w:del gives no notice", async () => {
    const { notices } = await readDocx(await build(para(run("a"), emptyDel(1))));
    expect(hiddenNotices(notices)).toEqual([]);
  });

  it("non-regression: a w:ins is published and gives no notice", async () => {
    const { doc, notices } = await readDocx(await build(para(run("a "), ins("insertedwords", 1))));
    expect(JSON.stringify(doc.blocks)).toContain("insertedwords");
    expect(hiddenNotices(notices)).toEqual([]);
  });

  it("non-regression: no tracked change gives no notice", async () => {
    const { notices } = await readDocx(await build(para(run("plain"))));
    expect(hiddenNotices(notices)).toEqual([]);
  });
});

describe("N1 + N2 together", () => {
  it("🔴 a hidden drawing and a deletion give two separate entries, one per `what`", async () => {
    const { notices } = await readDocx(await build(para(run("a"), picture("rIdRed", 1, HID), del("gone", 2))));
    const found = hiddenNotices(notices);
    expect(found).toHaveLength(2);
    expect(found).toContainEqual({ kind: "hidden-dropped", what: "shape", count: 1 });
    expect(found).toContainEqual({ kind: "hidden-dropped", what: "tracked-deletion", count: 1 });
  });
});
