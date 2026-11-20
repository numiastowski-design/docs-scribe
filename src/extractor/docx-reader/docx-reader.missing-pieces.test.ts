import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { encodePng } from "doc-ir/internal";
import { readDocx } from "./index";

// Decision of 08/10/2026: when a piece is missing or damaged, the rest is published AND the loss is said in a
// DocumentNotice. Docx half: a picture whose part is absent, a linked (external) picture, a damaged SmartArt model.
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/docx-reader/docx-reader.missing-pieces.test.ts
vi.setConfig({ testTimeout: 30_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = (path: string): Buffer => readFileSync(join(here, "../../../fixtures", path));
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${REL}"`;

interface Rel {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

const pictureRun = (blipAttr: string, id: number): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="254000" cy="254000"/><wp:docPr id="${id}" name="p${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip ${blipAttr}/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const embed = (relId: string, id: number): string => `<w:p>${pictureRun(`r:embed="${relId}"`, id)}</w:p>`;
const link = (relId: string, id: number): string => `<w:p>${pictureRun(`r:link="${relId}"`, id)}</w:p>`;
const text = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

/** A minimal docx; `model` adds the real SmartArt drawing cache and a data model in the given state. */
const build = async (opts: {
  body: string;
  rels?: Rel[];
  files?: Record<string, Uint8Array | string>;
  model?: "text" | "damaged" | "absent";
}): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const rels: Rel[] = [...(opts.rels ?? [])];
  if (opts.model) {
    const pptx = await JSZip.loadAsync(fixtures("diagrammes-reel/06-fr-la-veille.pptx"));
    rels.push({
      id: "rId6",
      type: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
      target: "diagrams/drawing1.xml",
    });
    zip.file("word/diagrams/drawing1.xml", await pptx.file("ppt/diagrams/drawing1.xml")!.async("string"));
    if (opts.model !== "absent") rels.push({ id: "rId2", type: `${REL}/diagramData`, target: "diagrams/data1.xml" });
    if (opts.model === "text")
      zip.file("word/diagrams/data1.xml", await pptx.file("ppt/diagrams/data1.xml")!.async("string"));
    if (opts.model === "damaged") zip.file("word/diagrams/data1.xml", "<dgm:dataModel <<< not xml");
  }
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels
        .map(
          (r) =>
            `<Relationship Id="${r.id}" Type="${r.type.startsWith("http") ? r.type : `${REL}/${r.type}`}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ""}/>`,
        )
        .join("") +
      "</Relationships>",
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${opts.body}</w:body></w:document>`);
  for (const [path, bytes] of Object.entries(opts.files ?? {})) zip.file(path, bytes);
  return zip.generateAsync({ type: "uint8array" });
};

// Two distinct pictures (distinct pixels, so they do not merge into one by hash).
const redDot = (): Uint8Array => encodePng(1, 1, "rgb-24", [255, 0, 0])!;
const blueDot = (): Uint8Array => encodePng(1, 1, "rgb-24", [0, 0, 255])!;

const imageMissing = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "image-missing");

describe("A1 — a picture whose relationship exists but whose part is absent from the archive", () => {
  const body = `${text("before")}${embed("rIdOk", 1)}${embed("rIdGone", 2)}${text("after")}`;
  const rels: Rel[] = [
    { id: "rIdOk", type: "image", target: "media/ok.png" },
    { id: "rIdGone", type: "image", target: "media/gone.png" },
  ];

  it("🔴 the document is published with its other picture and text, and one image-missing notice counts the lost picture", async () => {
    const { doc, images, notices } = await readDocx(
      await build({ body, rels, files: { "word/media/ok.png": redDot() } }),
    );
    expect(JSON.stringify(doc.blocks)).toContain("before");
    expect(JSON.stringify(doc.blocks)).toContain("after");
    expect(images).toHaveLength(1);
    expect(doc.blocks.filter((b: IRBlock) => b.kind === "image")).toHaveLength(1);
    // Failure value (before): no notice at all.
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });

  it("🔴 two pictures with absent parts are counted as 2 in one notice", async () => {
    const { notices } = await readDocx(
      await build({
        body: `${embed("rIdGone", 1)}${embed("rIdGone2", 2)}${text("kept")}`,
        rels: [
          { id: "rIdGone", type: "image", target: "media/gone.png" },
          { id: "rIdGone2", type: "image", target: "media/gone2.png" },
        ],
      }),
    );
    expect(imageMissing(notices)).toEqual([{ kind: "image-missing", count: 2 }]);
  });

  it("non-regression: every picture present, both published, no image-missing notice", async () => {
    const { images, notices } = await readDocx(
      await build({
        body: `${embed("rIdOk", 1)}${embed("rIdOk2", 2)}`,
        rels: [
          { id: "rIdOk", type: "image", target: "media/ok.png" },
          { id: "rIdOk2", type: "image", target: "media/ok2.png" },
        ],
        files: { "word/media/ok.png": redDot(), "word/media/ok2.png": blueDot() },
      }),
    );
    expect(images).toHaveLength(2);
    expect(imageMissing(notices)).toEqual([]);
  });
});

describe("A2 — a picture linked externally (TargetMode=External, r:link)", () => {
  const rels: Rel[] = [
    { id: "rIdOk", type: "image", target: "media/ok.png" },
    { id: "rIdLink", type: "image", target: "https://example.com/remote.png", external: true },
  ];

  it("🔴 published without it, and one image-missing notice (count 1)", async () => {
    const body = `${text("before")}${embed("rIdOk", 1)}${link("rIdLink", 2)}${text("after")}`;
    const { doc, images, notices } = await readDocx(
      await build({ body, rels, files: { "word/media/ok.png": redDot() } }),
    );
    expect(JSON.stringify(doc.blocks)).toContain("before");
    expect(JSON.stringify(doc.blocks)).toContain("after");
    expect(images).toHaveLength(1);
    // Failure value (before): no notice at all.
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });

  it("non-regression: an embedded picture next to an unused external image relationship is published, no notice", async () => {
    const { images, notices } = await readDocx(
      await build({ body: embed("rIdOk", 1), rels, files: { "word/media/ok.png": redDot() } }),
    );
    expect(images).toHaveLength(1);
    expect(imageMissing(notices)).toEqual([]);
  });
});

describe("A5 — a SmartArt whose data model is damaged but whose drawing cache is intact", () => {
  const smartGraphic =
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic>';
  const body = `${text("around")}<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/>${smartGraphic}</wp:inline></w:drawing></w:r></w:p>`;
  const WORDS = ["Mise sous surveillance des sources et mots-clés", "Repères de signaux significatifs"];
  const diagramWords = (value: unknown, out: string[] = []): string[] => {
    if (Array.isArray(value)) for (const v of value) diagramWords(v, out);
    else if (typeof value === "object" && value !== null) {
      const b = value as { kind?: string; sections?: { inlines: { text: string }[] }[] };
      if (b.kind === "diagramText") for (const s of b.sections ?? []) for (const i of s.inlines) out.push(i.text);
      for (const v of Object.values(value)) diagramWords(v, out);
    }
    return out;
  };

  it("🔴 the words of the diagram are published from the cache, the text around them too", async () => {
    const { doc } = await readDocx(await build({ body, model: "damaged" }));
    const words = diagramWords(doc.blocks).join("");
    // Failure value (before): the diagram's words are lost.
    for (const w of WORDS) expect(words).toContain(w);
    expect(JSON.stringify(doc.blocks)).toContain("around");
  });

  it("non-regression: the model absent gives the same words from the cache", async () => {
    const { doc } = await readDocx(await build({ body, model: "absent" }));
    const words = diagramWords(doc.blocks).join("");
    for (const w of WORDS) expect(words).toContain(w);
  });

  it("non-regression: the model intact gives the words once", async () => {
    const { doc } = await readDocx(await build({ body, model: "text" }));
    expect(diagramWords(doc.blocks).join("").split(WORDS[1]!).length - 1).toBe(1);
  });
});
