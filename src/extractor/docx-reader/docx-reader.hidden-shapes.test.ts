import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { encodePng } from "doc-ir/internal";
import { readDocx } from "./index";

// Decision of 08/10/2026: hidden content is never published. A drawing whose wp:docPr carries hidden="1"
// (or "true") is not published: neither its picture, nor its text-box words. The host paragraph and the other
// drawings stay. Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/docx-reader/docx-reader.hidden-shapes.test.ts
vi.setConfig({ testTimeout: 30_000 });

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:r="${REL}"`;

const redDot = (): Uint8Array => encodePng(1, 1, "rgb-24", [255, 0, 0])!;
const blueDot = (): Uint8Array => encodePng(1, 1, "rgb-24", [0, 0, 255])!;

/** `hidden` is the raw attribute text put on wp:docPr (e.g. ` hidden="1"`), empty for none. */
const picture = (relId: string, id: number, hidden: string, container: "inline" | "anchor" = "inline"): string => {
  const head =
    container === "inline"
      ? "<wp:inline>"
      : '<wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1000000</wp:posOffset></wp:positionV>';
  return (
    `<w:r><w:drawing>${head}<wp:extent cx="254000" cy="254000"/><wp:docPr id="${id}" name="p${id}"${hidden}/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    `<pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:${container}></w:drawing></w:r>`
  );
};

const textBox = (words: string, id: number, hidden: string): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="${id}" name="b${id}"${hidden}/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${words}</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const para = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const run = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;

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
  zip.file("word/media/red.png", redDot());
  zip.file("word/media/blue.png", blueDot());
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${body}${a4}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

const allText = (blocks: unknown): string => JSON.stringify(blocks);

describe("H1 — a hidden picture (wp:docPr hidden)", () => {
  it.each(["1", "true"])(
    '🔴 hidden="%s": the hidden picture is not published, the visible one beside it is, the host text stays',
    async (value) => {
      const body = para(run("host words"), picture("rIdRed", 1, ` hidden="${value}"`), picture("rIdBlue", 2, ""));
      const { doc, images } = await readDocx(await build(body));
      // Failure value (before): 2 images.
      expect(images).toHaveLength(1);
      expect(doc.blocks.filter((b) => b.kind === "image")).toHaveLength(1);
      expect(allText(doc.blocks)).toContain("host words");
    },
  );

  it("🔴 an anchored hidden picture alone: nothing published as image, the paragraph text stays", async () => {
    const { doc, images } = await readDocx(
      await build(para(run("anchor host"), picture("rIdRed", 1, ' hidden="1"', "anchor"))),
    );
    expect(images).toHaveLength(0);
    expect(doc.blocks.filter((b) => b.kind === "image")).toHaveLength(0);
    expect(allText(doc.blocks)).toContain("anchor host");
  });

  it.each([' hidden="0"', ' hidden="false"', ""])("non-regression: docPr%s publishes the picture", async (attr) => {
    const { doc, images } = await readDocx(
      await build(para(run("host"), picture("rIdRed", 1, attr), picture("rIdBlue", 2, attr))),
    );
    expect(images).toHaveLength(2);
    expect(doc.blocks.filter((b) => b.kind === "image")).toHaveLength(2);
  });
});

describe("H1 — a hidden text box (wp:docPr hidden)", () => {
  it.each(["1", "true"])(
    '🔴 hidden="%s": its words are not published; the visible box and the host paragraph are',
    async (value) => {
      const body = para(
        run("host paragraph"),
        textBox("secretwordsinbox", 1, ` hidden="${value}"`),
        textBox("visiblewordsinbox", 2, ""),
      );
      const { doc } = await readDocx(await build(body));
      const text = allText(doc.blocks);
      // Failure value (before): the secret words are published.
      expect(text).not.toContain("secretwordsinbox");
      expect(text).toContain("visiblewordsinbox");
      expect(text).toContain("host paragraph");
    },
  );

  it.each([' hidden="0"', ""])("non-regression: docPr%s publishes the box words", async (attr) => {
    const { doc } = await readDocx(await build(para(run("host"), textBox("boxwordspublished", 1, attr))));
    expect(allText(doc.blocks)).toContain("boxwordspublished");
  });
});
