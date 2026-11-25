import { decode as decodeJpegJs } from "jpeg-js";
import { describe, expect, it, vi } from "vitest";

import { decodeJpegRgba } from "./decode";
import { buildHuffmanTable } from "./huffman";
import { encodeJpegRgba, optimalHuffmanSpec, type RgbaImage } from "./encode";
import { fixtureJpegs, rgbaDistance } from "./jpeg-fixtures.test.support";

// Every fixture JPEG is encoded (and some decoded twice) in these tests.
vi.setConfig({ testTimeout: 300_000 });

/** The guards both product callers give the decoder. */
const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };

/** Starting values to calibrate, measured 06/10/2026 (mean absolute error per channel, 0..255 scale).
 *  MAE_GRADIENT: a ramp in each channel, quality 90; measured 0 (1×1), 3.79 (15×17), 4.07 (16×16), 6.23 (17×9), 0.80 (257×131):
 *  chroma averaged over 2×2 pixels loses most where the ramp steps fastest (a 16-pixel ramp steps 16 levels per pixel). Bound 6.5.
 *  MAE_FIXTURE: every sixth fixture JPEG, decoded then re-encoded at 80; measured 0.00 to 5.90 over the 12 pictures. Bound 6.5. */
const MAE_GRADIENT = 6.5;
const MAE_FIXTURE = 6.5;
/** Total bytes of the fixture JPEGs, decoded then re-encoded at quality 80, measured 06/10/2026: on the first 72, old 4:4:4
 *  encoder (jpeg-js's bytes) 5 799 000, this 4:2:0 encoder 4 904 491 (15.4 % fewer; libjpeg measured 14.7 %). The 73rd (the
 *  progressive 4:4:4 Adobe picture of fixtures/docx-reel/11-dgemc-programme-terminale.docx, added 06/10) measured alone with
 *  jpeg-js: 7 590 bytes. */
const OLD_444_TOTAL = 5_799_000 + 7_590;
/** Same fixtures, decoded then re-encoded at 80, measured 06/10/2026 (the decoder's fancy chroma upsampling already in; the 4 904 491 above
 *  predates it): on the first 72, example tables 4 865 428 and per-image tables 4 646 581 (4.5 % fewer; libjpeg's optimize_coding gains
 *  3 to 6 % on photographs); the 73rd adds 5 706 and 5 235. The 74th (the picture of
 *  fixtures/docx-reel/12-normene-maths-tle-pro.docx, added 29/10/2026 with the document of issue #60) adds 47 323 and
 *  43 635. These totals COUNT the corpus, they judge nothing: a fixture dropped in the tree moves them. */
const FIXED_TABLES_TOTAL = 4_865_428 + 5_706 + 47_323;
const OPTIMIZED_TOTAL = 4_646_581 + 5_235 + 43_635;

/** Deterministic pseudo-random bytes (an LCG), like the repository's tests. */
const noise = (length: number, seed = 1): Uint8Array => {
  const out = new Uint8Array(length);
  let s = seed >>> 0;
  for (let i = 0; i < length; i++) {
    s = (Math.imul(s, 1_103_515_245) + 12_345) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
};

const rgbaNoise = (width: number, height: number, seed = 1): RgbaImage => {
  const data = noise(width * height * 4, seed);
  for (let i = 3; i < data.length; i += 4) data[i] = 255;
  return { width, height, data };
};

const rgbaUniform = (width: number, height: number, [r, g, b]: [number, number, number]): RgbaImage => {
  const data = new Uint8Array(width * height * 4);
  for (let i = 0; i < data.length; i += 4) data.set([r, g, b, 255], i);
  return { width, height, data };
};

/** A smooth picture: what a photograph's blocks look like, unlike noise. */
const rgbaGradient = (width: number, height: number): RgbaImage => {
  const data = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      data.set(
        [
          (x * 255) / Math.max(1, width - 1),
          (y * 255) / Math.max(1, height - 1),
          ((x + y) * 255) / Math.max(1, width + height - 2),
          255,
        ],
        (y * width + x) * 4,
      );
    }
  }
  return { width, height, data };
};

const refusal = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(Error);
    return (e as Error).message;
  }
  throw new Error("expected a refusal");
};

/** Mean absolute error per channel (R, G, B; alpha ignored) between two pictures of the same size. */
const meanAbsoluteError = (a: RgbaImage, b: RgbaImage): number => {
  expect([a.width, a.height]).toEqual([b.width, b.height]);
  let sum = 0;
  for (let i = 0; i < a.data.length; i += 4) {
    sum +=
      Math.abs(a.data[i]! - b.data[i]!) +
      Math.abs(a.data[i + 1]! - b.data[i + 1]!) +
      Math.abs(a.data[i + 2]! - b.data[i + 2]!);
  }
  return sum / (a.width * a.height * 3);
};

/** Our encoder writes 4:2:0, which our decoder upsamples with libjpeg's fancy filter and jpeg-js's replicates (`color.ts`):
 *  same size, bounded distance. Measured 06/10/2026 (worst level, mean level, at most): 90 and 14.1
 *  on the gradient and noise images (chroma changing at every pixel: the worst case; bounds 100 and 15), 70 and 3.1 on the
 *  re-encoded fixtures (bounds 80 and 3.5). Valeurs de départ à calibrer.
 *  🔴 THE MEAN BOUND OF THE FIXTURES GOES FROM 3.5 TO 5.5 ON 29/10/2026, AND IT MEASURES NEITHER A REGRESSION NOR OUR
 *  QUALITY. Adding one fixture shifted « every sixth », which brought into the sample a picture nobody had measured:
 *  `ppt/media/image3.jfif` of a real third-party deck of the measurement corpus, a saturated ORANGE logo, where the two upsamplers
 *  stand 5.00 apart on average — and that distance is exactly what this number is, libjpeg's fancy filter against jpeg-js's
 *  replication on the SAME bytes. A saturated chroma edge is where they part. Two populations: 3.1 over the twelve pictures
 *  of 06/10/2026, 5.00 on this one. `MAE_FIXTURE` (the input against our own round trip, bound 6.5) passes on all of them. */
const expectNear = (a: RgbaImage, b: RgbaImage, label: string, worstBound: number, meanBound: number): void => {
  expect([a.width, a.height], label).toEqual([b.width, b.height]);
  const { worst, mean } = rgbaDistance(a.data, b.data);
  expect(worst, `${label}: worst ${worst}`).toBeLessThanOrEqual(worstBound);
  expect(mean, `${label}: mean ${mean}`).toBeLessThanOrEqual(meanBound);
};

const theirDecode = (jpeg: Uint8Array): RgbaImage => {
  const out = decodeJpegJs(jpeg, { useTArray: true, formatAsRGBA: true, tolerantDecoding: false, ...GUARD });
  return { width: out.width, height: out.height, data: new Uint8Array(out.data) };
};

/** Offset of the first marker `marker` (second byte) among the header segments, before the entropy-coded data. */
const segmentAt = (jpeg: Uint8Array, marker: number): number => {
  for (let i = 2; i < jpeg.length - 1;) {
    expect(jpeg[i]).toBe(0xff);
    if (jpeg[i + 1] === marker) return i;
    i += 2 + ((jpeg[i + 2]! << 8) | jpeg[i + 3]!);
  }
  throw new Error(`marker ${marker.toString(16)} not found`);
};

describe("encodeJpegRgba — 4:2:0 frame (T.81 A.1.1, A.2.3)", () => {
  it("declares Y 2×2, Cb and Cr 1×1 in SOF0, and the three components in SOS", () => {
    const jpeg = encodeJpegRgba(rgbaGradient(20, 10), 80);
    const sof = segmentAt(jpeg, 0xc0);
    // Marker(2) length(2) precision(1) height(2) width(2) count(1), then (id, H<<4|V, Tq) per component.
    expect(Array.from(jpeg.subarray(sof + 10, sof + 19))).toEqual([1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]);
    const sos = segmentAt(jpeg, 0xda);
    // Marker(2) length(2) count(1), then (id, table) per component.
    expect(Array.from(jpeg.subarray(sos + 4, sos + 11))).toEqual([3, 1, 0x00, 2, 0x11, 3, 0x11]);
  });

  it("keeps luminance at full resolution and averages the chroma of each 2×2 group: red and blue pixels alternating read back with their own brightness but one shared hue", () => {
    const width = 16;
    const height = 16;
    const data = new Uint8Array(width * height * 4);
    for (let i = 0; i < width * height; i++) data.set(i % 2 === 0 ? [255, 0, 0, 255] : [0, 0, 255, 255], i * 4);
    const back = decodeJpegRgba(encodeJpegRgba({ width, height, data }, 100), GUARD);
    const luma = (o: number): number =>
      (19595 * back.data[o]! + 38470 * back.data[o + 1]! + 7471 * back.data[o + 2]!) / 65536;
    // Red (Y 76) and blue (Y 29) keep their difference...
    expect(luma(0) - luma(4)).toBeGreaterThan(30);
    // ...while the chroma pair is one average: the two pixels have the same Cb, so the red pixel is not red any more.
    expect(back.data[0]! - back.data[1]!).toBeLessThan(200);
    expect(back.data[2]!).toBeGreaterThan(40);
  });
});

describe("encodeJpegRgba — read back by two independent decoders", () => {
  it("is read to bounded-near pixels by jpeg-js's decode and by our decoder (not equal: the chroma filters differ), at sizes that exercise the chroma edge padding", () => {
    for (const [width, height] of [
      [1, 1],
      [15, 17],
      [16, 16],
      [17, 9],
      [257, 131],
    ] as const) {
      for (const image of [rgbaGradient(width, height), rgbaNoise(width, height, width + height)]) {
        const ours = encodeJpegRgba(image, 90);
        expectNear(decodeJpegRgba(ours, GUARD), theirDecode(ours), `${width}×${height}`, 100, 15);
      }
    }
  });

  it("stays within MAE_GRADIENT (mean absolute error per channel) of a smooth source, in both decoders, at quality 90", () => {
    for (const [width, height] of [
      [1, 1],
      [15, 17],
      [16, 16],
      [17, 9],
      [257, 131],
    ] as const) {
      const image = rgbaGradient(width, height);
      const ours = encodeJpegRgba(image, 90);
      expect(meanAbsoluteError(image, decodeJpegRgba(ours, GUARD)), `${width}×${height}`).toBeLessThanOrEqual(
        MAE_GRADIENT,
      );
      expect(meanAbsoluteError(image, theirDecode(ours)), `${width}×${height}`).toBeLessThanOrEqual(MAE_GRADIENT);
    }
  });

  it("reads back every sixth fixture JPEG, decoded then re-encoded at 80, to near pixels in both decoders and within MAE_FIXTURE of the input", async () => {
    const jpegs = await fixtureJpegs();
    let checked = 0;
    for (const [index, { id, bytes }] of jpegs.entries()) {
      if (index % 6 !== 0) continue;
      const image = decodeJpegRgba(bytes, GUARD);
      const ours = encodeJpegRgba(image, 80);
      const back = decodeJpegRgba(ours, GUARD);
      expectNear(back, theirDecode(ours), id, 80, 5.5);
      expect(meanAbsoluteError(image, back), id).toBeLessThanOrEqual(MAE_FIXTURE);
      checked++;
    }
    expect(checked).toBeGreaterThanOrEqual(12);
  });

  it("stays within 7 levels of a smooth source at quality 100 (stated tolerance: chroma averaging and rounding loss)", () => {
    const image = rgbaGradient(64, 48);
    const back = decodeJpegRgba(encodeJpegRgba(image, 100), GUARD);
    let worst = 0;
    for (let i = 0; i < image.data.length; i++) worst = Math.max(worst, Math.abs(image.data[i]! - back.data[i]!));
    // Measured 5; a chroma group shifted by one row or column measures 23 to 25 (review of 06/10/2026).
    expect(worst).toBeLessThanOrEqual(7);
  });
});

/** The output with the example tables of T.81 annex K, the reference the optimized output must read back identically to. */
const fixedTables = (image: RgbaImage, quality: number): Uint8Array =>
  encodeJpegRgba(image, quality, { fixedTables: true });

/** The four DHT tables of `jpeg` (class and id byte, counts of each length, values), in file order. */
const dhtTables = (jpeg: Uint8Array): { id: number; counts: Uint8Array; values: Uint8Array }[] => {
  const at = segmentAt(jpeg, 0xc4);
  const end = at + 2 + ((jpeg[at + 2]! << 8) | jpeg[at + 3]!);
  const out: { id: number; counts: Uint8Array; values: Uint8Array }[] = [];
  for (let p = at + 4; p < end;) {
    const counts = jpeg.slice(p + 1, p + 17);
    const total = counts.reduce((a, b) => a + b, 0);
    out.push({ id: jpeg[p]!, counts, values: jpeg.slice(p + 17, p + 17 + total) });
    p += 17 + total;
  }
  return out;
};

/** Huffman coding is lossless: same coefficients, so the same pixels, in each decoder, tolerance 0. */
const expectSamePixels = (image: RgbaImage, quality: number, label: string): void => {
  const optimized = encodeJpegRgba(image, quality);
  const fixed = fixedTables(image, quality);
  expect(
    Buffer.compare(decodeJpegRgba(optimized, GUARD).data, decodeJpegRgba(fixed, GUARD).data),
    `${label}: our decoder`,
  ).toBe(0);
  expect(Buffer.compare(theirDecode(optimized).data, theirDecode(fixed).data), `${label}: jpeg-js`).toBe(0);
};

describe("encodeJpegRgba — per-image Huffman tables (T.81 annex K.2, libjpeg optimize_coding)", () => {
  it("writes DHT tables that differ from K.3 to K.6 on a real picture, none over 16 bits, none with the all-1 code", async () => {
    const [first] = await fixtureJpegs();
    const image = decodeJpegRgba(first!.bytes, GUARD);
    const optimized = dhtTables(encodeJpegRgba(image, 80));
    const fixed = dhtTables(fixedTables(image, 80));
    expect(optimized.map((t) => t.id)).toEqual([0x00, 0x10, 0x01, 0x11]);
    for (const [i, table] of optimized.entries()) {
      // buildHuffmanTable refuses an all-1 code and a table whose counts disagree with its values.
      expect(() => buildHuffmanTable(table), `table ${table.id}`).not.toThrow();
      expect(table.counts.length).toBe(16);
      expect(table.values.length, `table ${table.id}`).toBeGreaterThan(0);
      expect(Array.from(table.counts), `table ${table.id} is K.${3 + i}`).not.toEqual(Array.from(fixed[i]!.counts));
    }
    // Only the used symbols are declared: fewer than the 162 of the example AC tables.
    expect(optimized[1]!.values.length).toBeLessThan(162);
  });

  it("limits codes to 16 bits and keeps the all-1 code free, whatever the frequencies (T.81 K.2 figure K.3 adjustment)", () => {
    // Fibonacci frequencies give the deepest tree: 40 symbols plus the dummy need 21-bit codes unadjusted (measured, review of 06/10/2026), above the 16-bit cap.
    const freq = new Array<number>(256).fill(0);
    let a = 1;
    let b = 1;
    for (let s = 0; s < 40; s++) {
      freq[s] = a;
      [a, b] = [b, a + b];
    }
    const spec = optimalHuffmanSpec(freq);
    expect(spec.counts.length).toBe(16);
    expect(spec.values.length).toBe(40);
    expect(new Set(spec.values)).toEqual(new Set(Array.from({ length: 40 }, (_, i) => i)));
    expect(() => buildHuffmanTable(spec)).not.toThrow();
    // Value k of the list has the code size of the first length whose cumulative count exceeds k; the most frequent symbol is never longer than the rarest.
    const sizeOf = (symbol: number): number => {
      let k = Array.from(spec.values).indexOf(symbol);
      for (let size = 1; size <= 16; size++) {
        if (k < spec.counts[size - 1]!) return size;
        k -= spec.counts[size - 1]!;
      }
      throw new Error("symbol not in table");
    };
    expect(sizeOf(39)).toBeLessThan(sizeOf(0));
    expect(Math.max(...Array.from({ length: 40 }, (_, i) => sizeOf(i)))).toBe(16);
    expect(Array.from(spec.values.slice(0, 2))).toEqual([38, 39]);
  });

  it("builds a valid table from one used symbol, and from none", () => {
    const one = new Array<number>(256).fill(0);
    one[7] = 1000;
    const single = optimalHuffmanSpec(one);
    expect(Array.from(single.values)).toEqual([7]);
    expect(Array.from(single.counts)).toEqual([1, ...new Array<number>(15).fill(0)]);
    expect(() => buildHuffmanTable(single)).not.toThrow();
    expect(() => buildHuffmanTable(optimalHuffmanSpec(new Array<number>(256).fill(0)))).not.toThrow();
  });

  it("reads back to the same pixels as the fixed-table output (tolerance 0), in our decoder and in jpeg-js, on tiny, flat and noisy pictures", () => {
    for (const [width, height] of [
      [1, 1],
      [16, 16],
      [257, 131],
    ] as const) {
      expectSamePixels(rgbaUniform(width, height, [200, 30, 90]), 80, `flat ${width}×${height}`);
      expectSamePixels(rgbaNoise(width, height, width), 80, `noise ${width}×${height}`);
      expectSamePixels(rgbaGradient(width, height), 90, `gradient ${width}×${height}`);
    }
  });

  it("declares a one-symbol AC table for a flat picture, and still reads it back", () => {
    const jpeg = encodeJpegRgba(rgbaUniform(16, 16, [10, 200, 10]), 80);
    const ac = dhtTables(jpeg)[1]!;
    expect(Array.from(ac.values)).toEqual([0x00]);
    expect(decodeJpegRgba(jpeg, GUARD).width).toBe(16);
  });

  it("reads back every sixth fixture JPEG, re-encoded at 80, to the same pixels as the fixed-table output", async () => {
    const jpegs = await fixtureJpegs();
    for (const [index, { id, bytes }] of jpegs.entries()) {
      if (index % 6 !== 0) continue;
      expectSamePixels(decodeJpegRgba(bytes, GUARD), 80, id);
    }
  });

  it("writes fewer bytes in total than the example tables on the 73 fixture JPEGs, decoded then re-encoded at quality 80", async () => {
    const jpegs = await fixtureJpegs();
    let optimized = 0;
    let fixed = 0;
    for (const { bytes } of jpegs) {
      const image = decodeJpegRgba(bytes, GUARD);
      optimized += encodeJpegRgba(image, 80).length;
      fixed += fixedTables(image, 80).length;
    }
    expect(fixed).toBe(FIXED_TABLES_TOTAL);
    expect(optimized).toBeLessThan(fixed);
    expect(optimized).toBe(OPTIMIZED_TOTAL);
  });
});

describe("encodeJpegRgba — the end of the entropy-coded segment (T.81 F.1.2.3)", () => {
  it("pads a partial byte with 1-bits and adds no whole byte of 1-bits when the data already ends on a byte boundary", () => {
    // A flat picture ends with the Cr block's EOB, the 2-bit chroma AC code 00 (table K.6). The data ends aligned
    // when the last data byte's low 2 bits are 00, and padded when it ends with 00 then 1 to 6 one-bits. A final
    // 0xFF 0x00 is never legitimate here: 0xFF holds no 00 pair, so it could only be the old whole byte of 1-bits.
    let aligned = 0;
    let padded = 0;
    for (let v = 0; v < 256; v++) {
      for (const colour of [
        [v, v, v],
        [v, 255 - v, 128],
        [255, v, 0],
      ] as const) {
        const jpeg = encodeJpegRgba(rgbaUniform(16, 16, [colour[0], colour[1], colour[2]]), 80);
        expect(Array.from(jpeg.subarray(jpeg.length - 2))).toEqual([0xff, 0xd9]);
        const last = jpeg[jpeg.length - 3]!;
        expect(last === 0x00 && jpeg[jpeg.length - 4] === 0xff, colour.join(",")).toBe(false);
        if ((last & 3) === 0) aligned++;
        else if ([1, 2, 3, 4, 5, 6].some((k) => ((last >> k) & 3) === 0 && (last & ((1 << k) - 1)) === (1 << k) - 1))
          padded++;
      }
    }
    // Both situations were met, so the test is not vacuous.
    expect(aligned).toBeGreaterThan(0);
    expect(padded).toBeGreaterThan(0);
  });
});

describe("encodeJpegRgba — the 73 fixture JPEGs, decoded then re-encoded at quality 80", () => {
  it("writes at least 10 % fewer bytes than the old 4:4:4 encoder (OLD_444_TOTAL in total)", async () => {
    const jpegs = await fixtureJpegs();
    expect(jpegs.length).toBe(74);
    let total = 0;
    for (const { bytes } of jpegs) total += encodeJpegRgba(decodeJpegRgba(bytes, GUARD), 80).length;
    expect(total).toBeLessThanOrEqual(OLD_444_TOTAL * 0.9);
  });
});

describe("encodeJpegRgba — inputs refused (a caller's bug, never image content in the message)", () => {
  const ok = rgbaNoise(4, 4);

  it("refuses a width or a height outside 1..65 535 (T.81 B.2.2: 16 bits, 0 means a DNL marker we never write)", () => {
    expect(refusal(() => encodeJpegRgba({ ...ok, width: 0, data: new Uint8Array(0) }, 80))).toBe(
      "encodeJpegRgba: width and height must be integers in 1..65535",
    );
    expect(refusal(() => encodeJpegRgba({ ...ok, height: 0, data: new Uint8Array(0) }, 80))).toBe(
      "encodeJpegRgba: width and height must be integers in 1..65535",
    );
    expect(refusal(() => encodeJpegRgba({ width: 65_536, height: 1, data: new Uint8Array(65_536 * 4) }, 80))).toBe(
      "encodeJpegRgba: width and height must be integers in 1..65535",
    );
    expect(refusal(() => encodeJpegRgba({ ...ok, width: 2.5 }, 80))).toBe(
      "encodeJpegRgba: width and height must be integers in 1..65535",
    );
  });

  it("refuses pixels whose length is not width × height × 4 (the encoder would read undefined and write garbage)", () => {
    expect(refusal(() => encodeJpegRgba({ ...ok, data: new Uint8Array(4 * 4 * 4 - 1) }, 80))).toBe(
      "encodeJpegRgba: data must hold width × height × 4 bytes",
    );
    expect(refusal(() => encodeJpegRgba({ ...ok, data: new Uint8Array(4 * 4 * 3) }, 80))).toBe(
      "encodeJpegRgba: data must hold width × height × 4 bytes",
    );
  });

  it("refuses a quality that is not an integer in 1..100 (no silent clamping)", () => {
    for (const quality of [0, 101, 1.5, Number.NaN, -3]) {
      expect(
        refusal(() => encodeJpegRgba(ok, quality)),
        `quality ${quality}`,
      ).toBe("encodeJpegRgba: quality must be an integer in 1..100");
    }
  });
});
