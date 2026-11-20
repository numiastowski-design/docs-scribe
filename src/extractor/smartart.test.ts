import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readDocx } from "./docx-reader";
import { parseXml } from "./ooxml/xml";
import { readDataModelPieces, readDrawingPieces } from "./smartart";

// ECMA-376 Part 1 §21.4: `dgm:pt` text, `dgm:cxn` (no `type` = parOf) with `srcOrd`.
const here = dirname(fileURLToPath(import.meta.url));
const DGM =
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const pt = (id: string, text: string, type?: string): string =>
  `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ""}><dgm:t><a:bodyPr/><a:p>${text === "" ? "" : `<a:r><a:t>${text}</a:t></a:r>`}</a:p></dgm:t></dgm:pt>`;
const model = (points: string[], cxns: string): string =>
  `<dgm:dataModel ${DGM}><dgm:ptLst>${points.join("")}</dgm:ptLst><dgm:cxnLst>${cxns}</dgm:cxnLst></dgm:dataModel>`;
const cxn = (src: string, dest: string, srcOrd: number, type?: string): string =>
  `<dgm:cxn modelId="c${src}${dest}"${type ? ` type="${type}"` : ""} srcId="${src}" destId="${dest}" srcOrd="${srcOrd}" destOrd="0"/>`;
const piecesOf = (xml: string): string[] => readDataModelPieces(parseXml(xml)!).map((p) => p.text);
const levelsOf = (xml: string): number[] => readDataModelPieces(parseXml(xml)!).map((p) => p.level);

describe("readDataModelPieces", () => {
  it("tree order: the parOf tree by srcOrd, not the order of the points", () => {
    const xml = model(
      [pt("R", "", "doc"), pt("B", "second"), pt("A", "first"), pt("C", "child of first"), pt("T", "", "parTrans")],
      cxn("R", "B", 1) + cxn("R", "A", 0) + cxn("A", "C", 0),
    );
    // Failure value: document order, « second · first · child of first ».
    expect(piecesOf(xml)).toEqual(["first", "child of first", "second"]);
  });

  // The `doc` point carries no text: its children are the first level, so levels count from the smallest level
  // published, never from the technical root.
  it("tree level: the children of the doc point are level 0, their children level 1", () => {
    const xml = model(
      [
        pt("R", "", "doc"),
        pt("A", "a sustainable agency"),
        pt("C", "energy"),
        pt("D", "waste"),
        pt("B", "second root"),
      ],
      cxn("R", "A", 0) + cxn("A", "C", 0) + cxn("A", "D", 1) + cxn("R", "B", 1),
    );
    expect(piecesOf(xml)).toEqual(["a sustainable agency", "energy", "waste", "second root"]);
    // Failure value: [0, 0, 0, 0] — the flat list we publish today.
    expect(levelsOf(xml)).toEqual([0, 1, 1, 0]);
  });

  it("not a tree: every level is 0, so the list stays flat and invents no structure", () => {
    const xml = model(
      [pt("R", "root", "doc"), pt("A", "a"), pt("B", "b"), pt("C", "c")],
      cxn("R", "A", 0) + cxn("R", "B", 1) + cxn("A", "C", 0) + cxn("B", "C", 0),
    );
    expect(levelsOf(xml)).toEqual([0, 0, 0, 0]);
  });

  it("a point the tree does not reach keeps level 0", () => {
    const xml = model(
      [pt("R", "", "doc"), pt("S", "between", "sibTrans"), pt("A", "a"), pt("C", "child")],
      cxn("R", "A", 0) + cxn("A", "C", 0),
    );
    expect(piecesOf(xml)).toEqual(["a", "child", "between"]);
    expect(levelsOf(xml)).toEqual([0, 1, 0]);
  });

  it("no relations: document order of the points", () => {
    expect(piecesOf(model([pt("A", "one"), pt("B", "two")], ""))).toEqual(["one", "two"]);
  });

  it("not a tree (a point with two parents): document order, nothing dropped", () => {
    const xml = model(
      [pt("R", "root", "doc"), pt("A", "a"), pt("B", "b"), pt("C", "c")],
      cxn("R", "A", 0) + cxn("R", "B", 1) + cxn("A", "C", 0) + cxn("B", "C", 0),
    );
    expect(piecesOf(xml)).toEqual(["root", "a", "b", "c"]);
  });

  it("a cycle: nothing dropped", () => {
    const xml = model([pt("A", "a"), pt("B", "b")], cxn("A", "B", 0) + cxn("B", "A", 0));
    expect(piecesOf(xml).sort()).toEqual(["a", "b"]);
  });

  it("a cycle the walk enters (the doc point is on it): the document order, every level 0", () => {
    // R is the `doc` point and A's child as well as its parent: each point has one parent, and the walk from R comes back to R.
    const xml = model([pt("A", "a"), pt("R", "root", "doc")], cxn("R", "A", 0) + cxn("A", "R", 0));
    // Failure value: the walk's own order and depths, « root, a » at levels 0 and 1.
    expect(piecesOf(xml)).toEqual(["a", "root"]);
    expect(levelsOf(xml)).toEqual([0, 0]);
  });

  it("a connection whose end is no point: not a tree, the document order, every level 0", () => {
    const xml = model([pt("B", "b"), pt("A", "a")], cxn("A", "B", 0) + cxn("A", "missing", 1));
    // Failure value: the tree of the one valid connection, « a, b » at levels 0 and 1.
    expect(piecesOf(xml)).toEqual(["b", "a"]);
    expect(levelsOf(xml)).toEqual([0, 0]);
  });

  it("presentation relations are not the tree; a text on a sibling transition is kept, last", () => {
    const xml = model(
      [pt("R", "", "doc"), pt("S", "between", "sibTrans"), pt("B", "b"), pt("A", "a")],
      cxn("R", "B", 1) + cxn("R", "A", 0) + cxn("B", "A", 0, "presOf"),
    );
    expect(piecesOf(xml)).toEqual(["a", "b", "between"]);
  });

  it("paragraphs of one point are one piece, a line break a space", () => {
    const xml = `<dgm:dataModel ${DGM}><dgm:ptLst><dgm:pt modelId="A"><dgm:t><a:p><a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r></a:p><a:p><a:r><a:t>three</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>`;
    expect(piecesOf(xml)).toEqual(["one two three"]);
  });

  it("the drawing cache: one piece per shape with text", () => {
    const xml = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ${DGM}><dsp:spTree><dsp:sp><dsp:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></dsp:txBody></dsp:sp><dsp:sp><dsp:txBody><a:p/></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`;
    expect(readDrawingPieces(parseXml(xml)!)).toEqual(["x"]);
  });
});

// No Word SmartArt exists among the fixtures: a real SmartArt (PowerPoint's,
// `fixtures/diagrammes-reel/06-fr-la-veille.pptx`) is set into a minimal
// docx, the way Word writes the frame (`w:drawing`, `wp:inline`, `dgm:relIds`).
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const docxWithSmartArt = async (withModelText: boolean): Promise<Uint8Array> => {
  const src = await JSZip.loadAsync(readFileSync(join(here, "../../fixtures/diagrammes-reel/06-fr-la-veille.pptx")));
  const data = await src.file("ppt/diagrams/data1.xml")!.async("string");
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
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${REL}/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>`,
  );
  // Without the model's text, the cache (Office's own rendering) is the only source.
  zip.file("word/diagrams/data1.xml", withModelText ? data : data.replace(/<a:t>[^<]*<\/a:t>/g, "<a:t></a:t>"));
  zip.file("word/diagrams/drawing1.xml", await src.file("ppt/diagrams/drawing1.xml")!.async("string"));
  return zip.generateAsync({ type: "uint8array" });
};

describe("SmartArt in a docx", () => {
  it("its words come out in the one diagram-text block, in tree order, none added", async () => {
    const { doc } = await readDocx(await docxWithSmartArt(true));
    const block = doc.blocks.find((b) => b.kind === "diagramText");
    // Failure value (before): no diagramText block, the four texts lost.
    expect(block).toBeDefined();
    if (block?.kind !== "diagramText") throw new Error("unreachable");
    expect(block.sections).toHaveLength(1);
    // One inline per piece of the model, no separator: each one is an item of the published list (29/10/2026).
    expect(block.sections[0]!.inlines.map((i) => i.text)).toEqual([
      "Recherche d’information : sources, mots-clés, etc.",
      "Mise sous surveillance des sources et mots-clés",
      "Repères de signaux significatifs",
      "Recherche d’information dans le processus d’analyse : interrogation nouveaux outils",
    ]);
    expect(block.sections[0]!.breaks).toEqual([1, 2, 3]);
    expect(block.sections[0]!.levels).toHaveLength(4);
  });

  // The file says where the list goes: the `dgm:relIds` frame sits at one place of the body, where the picture
  // already comes out.
  it("the list follows its own picture, not the end of the document", async () => {
    const { doc } = await readDocx(await docxWithSmartArt(true));
    const at = doc.blocks.findIndex((b) => b.kind === "diagramText");
    // Failure value: the last block of the document, far from its drawing.
    expect(doc.blocks[at - 1]?.kind).toBe("image");
    // The text that follows the frame in the body still follows the list: nothing was pushed to the end.
    expect(doc.blocks.slice(at + 1).map((b) => b.kind)).toEqual(["paragraph"]);
  });

  it("a model with no text falls back to the drawing cache's shapes", async () => {
    const { doc } = await readDocx(await docxWithSmartArt(false));
    const block = doc.blocks.find((b) => b.kind === "diagramText");
    if (block?.kind !== "diagramText") throw new Error("expected a diagramText block");
    const text = block.sections[0]!.inlines.map((i) => i.text).join("");
    expect(text).toContain("Mise sous surveillance des sources et mots-clés");
  });
});

// Review of 01/10/2026 (security): the tree walk was recursive (a chain of
// 20 000 points overflowed the stack, which is no PartDamagedError, so the
// whole document failed). Past the point/edge bound the document order of
// the points stands: every word is still published. The case that bounds a
// DURATION (a star of 40 000 points) is in `smartart.cost.test.ts`.
describe("readDataModelPieces: bounded on hostile models", () => {
  it("a chain of 20 000 points does not throw, and no word is lost", () => {
    const n = 20_000;
    const points = Array.from({ length: n }, (_, i) => pt(`p${i}`, `w${i}`));
    const cxns = Array.from({ length: n - 1 }, (_, i) => cxn(`p${i}`, `p${i + 1}`, 0)).join("");
    // Failure value: RangeError, Maximum call stack size exceeded.
    expect(piecesOf(model(points, cxns))).toEqual(Array.from({ length: n }, (_, i) => `w${i}`));
  }, 60_000);

  it("a chain under the bound, its points listed backwards, is read in chain order (iterative walk)", () => {
    const n = 4_000;
    const points = Array.from({ length: n }, (_, i) => pt(`p${n - 1 - i}`, `w${n - 1 - i}`));
    const cxns = Array.from({ length: n - 1 }, (_, i) => cxn(`p${i}`, `p${i + 1}`, 0)).join("");
    expect(piecesOf(model(points, cxns))).toEqual(Array.from({ length: n }, (_, i) => `w${i}`));
  }, 60_000);
});
