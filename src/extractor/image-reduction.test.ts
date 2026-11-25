import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { decode as decodeJpeg, encode as encodeJpeg } from "jpeg-js";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { imagePixelSize, MAX_LONG_SIDE_PX, reduceLargeImage, resampleBilinear } from "./image-reduction";
import { encodePngDeflate } from "doc-ir/internal";
import { extractPptx } from "./pptx";

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));
const ROTATED = "pptx/04-Images-tournees.pptx";

/** Deterministic noise (an LCG): incompressible pixels, so a JPEG of them
 *  is lighter than a PNG of them, as for a photograph. */
const noise = (length: number, seed = 1): Uint8Array => {
  const out = new Uint8Array(length);
  let s = seed;
  for (let i = 0; i < length; i++) {
    s = (Math.imul(s, 1_103_515_245) + 12_345) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
};

const rgbaNoise = (w: number, h: number): Uint8Array => {
  const data = noise(w * h * 4);
  for (let i = 3; i < data.length; i += 4) data[i] = 255;
  return data;
};

const jpegOf = (w: number, h: number): Uint8Array =>
  new Uint8Array(encodeJpeg({ width: w, height: h, data: rgbaNoise(w, h) }, 90).data);

const decoded = (bytes: Uint8Array): { width: number; height: number; data: Uint8Array } =>
  decodeJpeg(bytes, { useTArray: true, formatAsRGBA: true });

/** An APP1 EXIF segment declaring `orientation`, big-endian TIFF, one IFD0 entry. */
const exifSegment = (orientation: number): Uint8Array => {
  const tiff = [
    0x4d,
    0x4d,
    0x00,
    0x2a,
    0,
    0,
    0,
    8,
    0,
    1,
    0x01,
    0x12,
    0,
    3,
    0,
    0,
    0,
    1,
    0,
    orientation,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
  ];
  const payload = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff];
  const length = payload.length + 2;
  return new Uint8Array([0xff, 0xe1, length >> 8, length & 0xff, ...payload]);
};

const withExif = (jpeg: Uint8Array, orientation: number): Uint8Array => {
  const seg = exifSegment(orientation);
  const out = new Uint8Array(jpeg.length + seg.length);
  out.set(jpeg.subarray(0, 2), 0);
  out.set(seg, 2);
  out.set(jpeg.subarray(2), 2 + seg.length);
  return out;
};

describe("resampleBilinear", () => {
  it("averages two pixels into one, halfway between their centres", () => {
    const src = new Uint8Array([0, 0, 0, 255, 200, 100, 50, 255]);
    expect([...resampleBilinear(src, 2, 1, 1, 1)]).toEqual([100, 50, 25, 255]);
  });

  it("keeps a uniform image uniform", () => {
    const src = new Uint8Array(40 * 30 * 4).map((_, i) => (i % 4 === 3 ? 255 : 77));
    const out = resampleBilinear(src, 40, 30, 16, 12);
    expect(out.length).toBe(16 * 12 * 4);
    expect(out.every((v, i) => v === (i % 4 === 3 ? 255 : 77))).toBe(true);
  });
});

describe("imagePixelSize", () => {
  it("reads both sides of a PNG and of a JPEG from their headers", () => {
    expect(imagePixelSize(encodePngDeflate(7, 3, 3, new Uint8Array(7 * 3 * 3))!)).toEqual({ width: 7, height: 3 });
    expect(imagePixelSize(jpegOf(9, 5))).toEqual({ width: 9, height: 5 });
    expect(imagePixelSize(new TextEncoder().encode("<svg/>"))).toBeUndefined();
  });
});

describe("reduceLargeImage", () => {
  it("leaves an image at or under the cap untouched (no result: same bytes, same sha)", () => {
    expect(reduceLargeImage(jpegOf(MAX_LONG_SIDE_PX, 8))).toBeUndefined();
    expect(reduceLargeImage(jpegOf(8, MAX_LONG_SIDE_PX))).toBeUndefined();
    expect(
      reduceLargeImage(encodePngDeflate(MAX_LONG_SIDE_PX, 4, 3, noise(MAX_LONG_SIDE_PX * 4 * 3))!),
    ).toBeUndefined();
  });

  it("brings a wide JPEG to 1 600 px on its long side, JPEG, width reported", () => {
    const reduced = reduceLargeImage(jpegOf(1700, 170))!;
    expect(reduced.mimeType).toBe("image/jpeg");
    expect(reduced.pixelWidth).toBe(1600);
    const out = decoded(reduced.bytes);
    expect([out.width, out.height]).toEqual([1600, 160]);
  });

  it("brings a tall JPEG to 1 600 px on its HEIGHT", () => {
    const reduced = reduceLargeImage(jpegOf(120, 1800))!;
    expect([decoded(reduced.bytes).width, decoded(reduced.bytes).height]).toEqual([107, 1600]);
    expect(reduced.pixelWidth).toBe(107);
  });

  it("turns a large opaque PNG into a JPEG", () => {
    const png = encodePngDeflate(2000, 100, 3, noise(2000 * 100 * 3))!;
    const reduced = reduceLargeImage(png)!;
    expect(reduced.mimeType).toBe("image/jpeg");
    expect([decoded(reduced.bytes).width, decoded(reduced.bytes).height]).toEqual([1600, 80]);
    expect(reduced.bytes.length).toBeLessThan(png.length);
  });

  it("lays a transparent PNG on WHITE (chosen bias), never on black", () => {
    const pixels = noise(2000 * 100 * 4);
    for (let i = 3; i < pixels.length; i += 4) pixels[i] = 0; // every pixel fully transparent
    const reduced = reduceLargeImage(encodePngDeflate(2000, 100, 4, pixels)!)!;
    const out = decoded(reduced.bytes).data;
    const darkest = out.reduce((m, v) => Math.min(m, v), 255);
    expect(darkest).toBeGreaterThan(245);
  });

  it("reads a grey PNG as grey", () => {
    const grey = new Uint8Array(3000 * 50).fill(128);
    // Noise in one row so the JPEG is not larger than the PNG.
    grey.set(noise(3000 * 10), 0);
    const reduced = reduceLargeImage(encodePngDeflate(3000, 50, 1, grey)!)!;
    const out = decoded(reduced.bytes);
    const at = (out.height - 1) * out.width * 4;
    expect(Math.abs(out.data[at]! - 128)).toBeLessThan(4);
    expect(Math.abs(out.data[at]! - out.data[at + 2]!)).toBeLessThan(4);
  });

  it("lays a grey + alpha PNG on white: a transparent grey pixel comes out white, an opaque one stays grey", () => {
    const pairs = new Uint8Array(3000 * 50 * 2);
    // Opaque mid-grey, then the lower half fully transparent black; noise on the top rows so the JPEG is the lighter file.
    for (let i = 0; i < 3000 * 50; i++) pairs.set(i < 3000 * 25 ? [128, 255] : [0, 0], i * 2);
    pairs.set(noise(3000 * 2 * 10), 0);
    for (let i = 1; i < 3000 * 10 * 2; i += 2) pairs[i] = 255;
    const out = decoded(reduceLargeImage(encodePngDeflate(3000, 50, 2, pairs)!)!.bytes);
    const row = (y: number): number => out.data[(y * out.width + 800) * 4]!;
    expect(Math.abs(row(Math.floor(out.height * 0.4)) - 128)).toBeLessThan(6);
    expect(row(out.height - 2)).toBeGreaterThan(245);
  });

  it("is deterministic: same input, same bytes", () => {
    const jpeg = jpegOf(2400, 60);
    expect(reduceLargeImage(jpeg)!.bytes).toEqual(reduceLargeImage(jpeg)!.bytes);
  });

  it("keeps the original when the reduction would not be lighter (a flat PNG)", () => {
    expect(reduceLargeImage(encodePngDeflate(4000, 10, 3, new Uint8Array(4000 * 10 * 3).fill(200))!)).toBeUndefined();
  });

  it("leaves a JPEG turned by its EXIF orientation untouched, and reduces one that declares no turn", () => {
    const jpeg = jpegOf(3200, 40);
    expect(reduceLargeImage(withExif(jpeg, 6))).toBeUndefined();
    expect(reduceLargeImage(withExif(jpeg, 1))).toBeDefined();
  });

  it("leaves a damaged JPEG untouched (no grey blocks are re-encoded)", () => {
    const whole = jpegOf(1700, 64);
    const damaged = whole.subarray(0, Math.floor(whole.length / 2));
    expect(imagePixelSize(damaged)?.width).toBe(1700);
    expect(reduceLargeImage(damaged)).toBeUndefined();
    expect(reduceLargeImage(whole)).toBeDefined(); // control: the intact file is reduced
  });

  it("leaves another format untouched", () => {
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x40, 0x1f, 0x40, 0x1f, 0, 0, 0]);
    expect(reduceLargeImage(gif)).toBeUndefined();
  });
});

describe("a presentation's images through extractPptx", () => {
  it("rotated-images deck: our drawing is untouched and still carries the small picture as it is", async () => {
    const pptx = fixture(ROTATED);
    const media = await (await JSZip.loadAsync(pptx)).file("ppt/media/image1.png")!.async("uint8array");
    const { images } = await extractPptx(pptx);
    const svgs = images.filter((i) => i.mimeType === "image/svg+xml");
    expect(images.length).toBe(svgs.length);
    const base64 = Buffer.from(media).toString("base64");
    expect(svgs.some((i) => new TextDecoder().decode(i.bytes).includes(base64))).toBe(true);
  });
});
