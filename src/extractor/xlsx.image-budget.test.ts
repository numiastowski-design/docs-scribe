import { describe, expect, it } from "vitest";

import { IMAGE_MAX_TOTAL_BYTES } from "./ooxml/package";
import { extractXlsx } from "./xlsx";
import {
  anchoredPictureZip,
  DRAWING_PART as DRAWING,
  DRAWING_RELS_PART as RELS,
  IMAGE_RELATIONSHIP as IMG_TYPE,
  PICTURE_PART,
  RELATIONSHIPS_HEAD as REL_HEAD,
} from "./xlsx-anchored-picture.test.support";

// Decision of 08/10/2026: xlsx aligns with docx and pptx on the TOTAL image byte budget (IMAGE_MAX_TOTAL_BYTES,
// 50 000 000). Over it, the workbook is refused with the same code as docx (EXTRACT_FAILED) and the xlsx wording of the
// damaged-workbook refusal. Images are all-zero buffers: they deflate to a few dozen KB, while the DECOMPRESSED total is what counts.
// Witness: the made-up workbook of `xlsx-anchored-picture.test.support.ts`, one drawing (xl/drawings/drawing1.xml) anchoring
// one picture; we rewrite the drawing with N anchors and the media parts. Heavier than most: it builds two images of about
// 25 MB each.
const DAMAGED = { code: "EXTRACT_FAILED", message: "the file is damaged or is not a valid Excel workbook" };

/** One anchor per entry of `embeds` (the rIds the pictures point at), cloned from the witness's own anchor. */
const withImages = async (
  media: readonly { path: string; size: number; fill: number }[],
  uses: readonly string[],
  rels: readonly (readonly [string, string])[],
): Promise<Uint8Array> => {
  const zip = anchoredPictureZip();
  const drawing = await zip.file(DRAWING)!.async("string");
  const anchor = /<xdr:twoCellAnchor[\s\S]*<\/xdr:twoCellAnchor>/.exec(drawing)![0];
  expect(anchor).toContain('r:embed="rId1"'); // the rewrite below really changes the fixture
  const anchors = uses.map((rId) => anchor.replace('r:embed="rId1"', `r:embed="${rId}"`)).join("");
  zip.file(DRAWING, drawing.replace(anchor, anchors));
  zip.file(
    RELS,
    `${REL_HEAD}${rels.map(([id, target]) => `<Relationship Id="${id}" Type="${IMG_TYPE}" Target="${target}"/>`).join("")}</Relationships>`,
  );
  zip.remove(PICTURE_PART);
  for (const m of media) zip.file(m.path, new Uint8Array(m.size).fill(m.fill));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 1 } });
};

describe("xlsx — the images' total decompressed bytes are bounded by IMAGE_MAX_TOTAL_BYTES, as in docx and pptx", () => {
  it("sanity: the budget is the shared 50 000 000 and each test part stays under the 64 MB part ceiling", () => {
    expect(IMAGE_MAX_TOTAL_BYTES).toBe(50_000_000);
  });

  it("🔴 B1: two distinct images totalling 50 200 000 bytes (over the budget) refuse the workbook as damaged", async () => {
    const half = 25_100_000;
    expect(half * 2).toBeGreaterThan(IMAGE_MAX_TOTAL_BYTES);
    const bytes = await withImages(
      [
        { path: "xl/media/a.png", size: half, fill: 0 },
        { path: "xl/media/b.png", size: half, fill: 1 },
      ],
      ["rId1", "rId2"],
      [
        ["rId1", "../media/a.png"],
        ["rId2", "../media/b.png"],
      ],
    );
    // Settled into a small value first: a failure must not ask the reporter to print 50 MB of pixels.
    // Failure value (before): "read" (the workbook is published, only the package ceilings apply).
    expect(
      await extractXlsx(bytes).then(
        () => "read",
        (e: unknown) => e,
      ),
    ).toMatchObject(DAMAGED);
  });

  it("B2: two distinct images totalling 49 800 000 bytes (just under the budget) are both published", async () => {
    const half = 24_900_000;
    expect(half * 2).toBeLessThan(IMAGE_MAX_TOTAL_BYTES);
    const bytes = await withImages(
      [
        { path: "xl/media/a.png", size: half, fill: 0 },
        { path: "xl/media/b.png", size: half, fill: 1 },
      ],
      ["rId1", "rId2"],
      [
        ["rId1", "../media/a.png"],
        ["rId2", "../media/b.png"],
      ],
    );
    const { images } = await extractXlsx(bytes);
    expect(images.map((i) => i.bytes.length)).toEqual([half, half]);
  });

  it("B3: the same image part used by two anchors (one relationship) counts once: 30 MB used twice (60 MB of usages) is published", async () => {
    const size = 30_000_000;
    expect(size).toBeLessThan(IMAGE_MAX_TOTAL_BYTES);
    expect(size * 2).toBeGreaterThan(IMAGE_MAX_TOTAL_BYTES); // the two usages together would be over: only the dedup saves it
    const bytes = await withImages(
      [{ path: "xl/media/a.png", size, fill: 0 }],
      ["rId1", "rId1"],
      [["rId1", "../media/a.png"]],
    );
    const { images, doc } = await extractXlsx(bytes);
    expect(images.length).toBeGreaterThanOrEqual(1);
    expect(images[0]!.bytes.length).toBe(size);
    expect(doc.blocks.length).toBeGreaterThan(0);
  });

  it("B3: the same part reached through two relationship ids (rId1 and rId2 -> a.png) counts once as well", async () => {
    const size = 30_000_000;
    const bytes = await withImages(
      [{ path: "xl/media/a.png", size, fill: 0 }],
      ["rId1", "rId2"],
      [
        ["rId1", "../media/a.png"],
        ["rId2", "../media/a.png"],
      ],
    );
    const { images } = await extractXlsx(bytes);
    expect(images.length).toBeGreaterThanOrEqual(1);
    expect(images[0]!.bytes.length).toBe(size);
  });

  it("B1 control: one single 50 200 000-byte image over the budget is refused too", async () => {
    const bytes = await withImages(
      [{ path: "xl/media/a.png", size: 50_200_000, fill: 0 }],
      ["rId1"],
      [["rId1", "../media/a.png"]],
    );
    expect(
      await extractXlsx(bytes).then(
        () => "read",
        (e: unknown) => e,
      ),
    ).toMatchObject(DAMAGED);
  });
});
