/**
 * The REAL SmartArt files of the repository, published as drawings:
 * `fixtures/diagrammes-reel/06-fr-la-veille.pptx` (slides 7 and 14) and `07-fr-trouver-une-image.pptx` (slides 3, 4, 5, 7). The expected text is read off each drawing's XML by the test, with
 * its own code (never the extractor's), then compared with the `<text>` of the published SVG, whitespace removed.
 * Run: npx vitest run src/extractor/smartart-drawing.fixtures.test.ts
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractPptx } from "./pptx";
import { drawingCompact, imageRefs, svgCompact } from "./smartart-drawing.test.support";

vi.setConfig({ testTimeout: 30_000 });

const here = dirname(fileURLToPath(import.meta.url));
const bytesOf = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));
const VEILLE = "diagrammes-reel/06-fr-la-veille.pptx";
const IMAGE = "diagrammes-reel/07-fr-trouver-une-image.pptx";

/** Relationship targets of a part, by id, resolved to archive paths. */
const relsOf = async (zip: JSZip, holder: string): Promise<Map<string, string>> => {
  const dir = holder.slice(0, holder.lastIndexOf("/"));
  const xml = await zip.file(`${dir}/_rels/${holder.slice(holder.lastIndexOf("/") + 1)}.rels`)!.async("string");
  const map = new Map<string, string>();
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])![1]!;
    const target = /Target="([^"]+)"/.exec(m[0])![1]!;
    const parts = `${dir}/${target}`.split("/");
    const out: string[] = [];
    for (const p of parts) {
      if (p === "..") out.pop();
      else out.push(p);
    }
    map.set(id, out.join("/"));
  }
  return map;
};

/** The drawing parts of the SmartArt frames a part holds, in frame order: the file's own pairing (frame, `r:dm`, the model's `dsp:dataModelExt relId`). */
const drawingsOf = async (zip: JSZip, holder: string): Promise<string[]> => {
  const rels = await relsOf(zip, holder);
  const body = await zip.file(holder)!.async("string");
  const out: string[] = [];
  for (const m of body.matchAll(/<dgm:relIds\b[^>]*r:dm="([^"]+)"/g)) {
    const data = await zip.file(rels.get(m[1]!)!)!.async("string");
    out.push(await zip.file(rels.get(/<dsp:dataModelExt\b[^>]*relId="([^"]+)"/.exec(data)![1]!)!)!.async("string"));
  }
  return out;
};

const tabsOf = (blocks: readonly IRBlock[]): { slideNumber?: number; blocks: IRBlock[] }[] => {
  const tabs = blocks.find((b) => b.kind === "tabs");
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  return tabs.tabs;
};

const svgOfBlocks = (
  result: { images: readonly { sha256: string; bytes: Uint8Array; mimeType: string }[] },
  blocks: readonly IRBlock[],
): string[] => {
  const bySha = new Map(result.images.map((i) => [i.sha256, i]));
  return imageRefs(blocks)
    .filter((r) => r.mimeType === "image/svg+xml")
    .map((r) => new TextDecoder().decode(bySha.get(r.sha256)!.bytes));
};

describe.each([
  [VEILLE, [7, 14]],
  [IMAGE, [3, 4, 5, 7]],
] as const)("%s: every SmartArt slide carries its drawing", (file, slides) => {
  it(`slides ${slides.join(", ")} publish a picture holding the words of their SmartArt, in the drawing's order`, async () => {
    const zip = await JSZip.loadAsync(bytesOf(file));
    const result = await extractPptx(bytesOf(file));
    const tabs = tabsOf(result.doc.blocks);
    for (const number of slides) {
      const expected = (await drawingsOf(zip, `ppt/slides/slide${number}.xml`)).map(drawingCompact);
      expect(
        expected.every((e) => e.length > 0),
        `slide ${number}: the fixture holds text`,
      ).toBe(true);
      const tab = tabs.find((t) => t.slideNumber === number);
      expect(tab, `tab of slide ${number}`).toBeDefined();
      const drawn = svgOfBlocks(result, tab!.blocks).map(svgCompact).join("|");
      // Failure value: « » (no picture on a table-form or flow slide).
      for (const text of expected) expect(drawn, `slide ${number}`).toContain(text);
    }
  });
});

describe("colours of the real drawings ([C5])", () => {
  it("06-fr-la-veille slide 7: boxes whose fill is accent1 with zero offsets come out filled (not black on white)", async () => {
    const result = await extractPptx(bytesOf(VEILLE));
    const tab = tabsOf(result.doc.blocks).find((t) => t.slideNumber === 7)!;
    const svg = svgOfBlocks(result, tab.blocks).join("");
    // Failure value: no picture. After the change the 4 text boxes carry an exact fill (the theme's accent1).
    expect(svg).toMatch(/<rect\b[^>]*fill="#[0-9A-F]{6}"/);
  });

  it("06-fr-la-veille slide 14 (a gradient fill on its text boxes): outline and words, no invented colour", async () => {
    const zip = await JSZip.loadAsync(bytesOf(VEILLE));
    const [cache] = await drawingsOf(zip, "ppt/slides/slide14.xml");
    expect(cache).toContain("<a:gradFill");
    const result = await extractPptx(bytesOf(VEILLE));
    const tab = tabsOf(result.doc.blocks).find((t) => t.slideNumber === 14)!;
    const svg = svgOfBlocks(result, tab.blocks).join("");
    expect(svgCompact(svg)).toContain(drawingCompact(cache!));
    expect(svg).not.toMatch(/<rect\b[^>]*fill="url\(/);
  });
});
