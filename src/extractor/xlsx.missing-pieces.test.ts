import type JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractXlsx } from "./xlsx";
import {
  anchoredPictureWorkbook,
  anchoredPictureZip,
  DRAWING_PART as DRAWING,
  DRAWING_RELS_PART as RELS,
  IMAGE_RELATIONSHIP as IMG_TYPE,
  PICTURE_PART,
  PNG_CONTENT_TYPE,
  RELATIONSHIPS_HEAD as REL_HEAD,
} from "./xlsx-anchored-picture.test.support";

// Decision of 08/10/2026: a missing or damaged piece is said, the rest is published. Xlsx half: a drawing image
// linked externally or whose relationship does not resolve (said as image-missing), and an image whose content type is
// not declared (published as application/octet-stream, like docx and pptx, instead of being dropped as "missing").
// Witness: the made-up workbook of `xlsx-anchored-picture.test.support.ts`, one picture (rId1 -> xl/media/image1.png)
// anchored in sheet 1.

const edited = async (edit: (zip: JSZip) => Promise<void>): Promise<Uint8Array> => {
  const zip = anchoredPictureZip();
  await edit(zip);
  return zip.generateAsync({ type: "uint8array" });
};

const imageMissing = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "image-missing");

describe("xlsx — a drawing image that cannot be read is said, the workbook is published", () => {
  it("non-regression: the untouched workbook publishes its picture and says nothing", async () => {
    const { images, notices, doc } = await extractXlsx(await anchoredPictureWorkbook());
    expect(images).toHaveLength(1);
    expect(imageMissing(notices)).toEqual([]);
    expect(doc.blocks.length).toBeGreaterThan(0);
  });

  it("🔴 A3: an image linked externally (r:link, TargetMode=External) is left out and said: one image-missing notice", async () => {
    const bytes = await edited(async (zip) => {
      zip.file(
        RELS,
        `${REL_HEAD}<Relationship Id="rId1" Type="${IMG_TYPE}" Target="https://example.com/remote.png" TargetMode="External"/></Relationships>`,
      );
      const drawing = await zip.file(DRAWING)!.async("string");
      expect(drawing).toContain('r:embed="rId1"'); // the removal below really changes the fixture
      zip.file(DRAWING, drawing.replace('r:embed="rId1"', 'r:link="rId1"'));
    });
    const { images, notices, doc } = await extractXlsx(bytes);
    expect(images).toHaveLength(0);
    expect(doc.blocks.length).toBeGreaterThan(0);
    // Failure value (before): no notice at all.
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });

  it("🔴 A3: an image relationship that does not resolve (unknown Id) is left out and said: one image-missing notice", async () => {
    const bytes = await edited(async (zip) => {
      zip.file(
        RELS,
        `${REL_HEAD}<Relationship Id="rId9" Type="${IMG_TYPE}" Target="../media/image1.png"/></Relationships>`,
      );
    });
    const { images, notices, doc } = await extractXlsx(bytes);
    expect(images).toHaveLength(0);
    expect(doc.blocks.length).toBeGreaterThan(0);
    // Failure value (before): no notice at all.
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });

  it("non-regression: a picture whose part is absent from the archive is already said (one image-missing notice)", async () => {
    const bytes = await edited(async (zip) => {
      zip.remove(PICTURE_PART);
    });
    const { images, notices } = await extractXlsx(bytes);
    expect(images).toHaveLength(0);
    expect(notices ?? []).toContainEqual({ kind: "image-missing", count: 1 });
  });
});

describe("xlsx — an image whose content type is not declared is published, as in docx and pptx", () => {
  const withoutPngType = (): Promise<Uint8Array> =>
    edited(async (zip) => {
      const types = await zip.file("[Content_Types].xml")!.async("string");
      const stripped = types.replace(PNG_CONTENT_TYPE, "");
      expect(stripped).not.toBe(types); // the fixture really declared it: the test exercises the removal
      zip.file("[Content_Types].xml", stripped);
    });

  it("🔴 A4: the image is in the result as application/octet-stream and no image-missing is emitted", async () => {
    const { images, notices, doc } = await extractXlsx(await withoutPngType());
    // Failure value (before): images is empty and image-missing is emitted.
    expect(images.map((i) => i.mimeType)).toEqual(["application/octet-stream"]);
    expect(imageMissing(notices)).toEqual([]);
    const tab = (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs[0]!;
    expect(tab.blocks.some((b) => b.kind === "image")).toBe(true);
  });

  it("non-regression: with the type declared, the same image is published as image/png", async () => {
    const { images } = await extractXlsx(await anchoredPictureWorkbook());
    expect(images.map((i) => i.mimeType)).toEqual(["image/png"]);
  });
});
