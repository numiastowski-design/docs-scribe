import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { decode as decodeJpegJs, encode as encodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { decodeJpegRgba } from "./decode";
import { fixtureJpegs, isSubsampled, rgbaDistance } from "./jpeg-fixtures.test.support";
import { parseJpeg } from "./markers";

/** The options both product callers pass to jpeg-js (`image-reduction.ts`, `pptx-reader/picture.ts`). */
const CALLERS = {
  useTArray: true,
  formatAsRGBA: true,
  tolerantDecoding: false,
  maxResolutionInMP: 12,
  maxMemoryUsageInMB: 256,
} as const;
const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };
const FIXTURES_JPEG = join(dirname(fileURLToPath(import.meta.url)), "../../../fixtures/jpeg");

const reference = (bytes: Uint8Array, opts: { maxMemoryUsageInMB?: number } = {}) =>
  decodeJpegJs(bytes, { ...CALLERS, ...opts });

/** Chroma upsampling is libjpeg's fancy filter for 2:1 ratios (`color.ts`), no longer jpeg-js's replication: against jpeg-js an
 *  image with no subsampling is identical byte for byte, a subsampled one has the same size and a bounded distance.
 *  Measured 06/10/2026 on the 65 subsampled fixture JPEGs: worst level 108 (sharp coloured edge, 4:2:0), worst mean 5.66
 *  (the same image); the synthetic DC patterns below stay within the same bounds. The exact reading of the chroma step is checked against
 *  Pillow in `upsampling.test.ts`. Bounds are valeurs de départ à calibrer. */
const UPSAMPLING_WORST = 120;
const UPSAMPLING_MEAN = 6.5;
const sameAsJpegJs = (ours: Uint8Array, ref: Uint8Array, subsampled: boolean, label: string): void => {
  if (!subsampled) {
    expect(Buffer.from(ours).equals(Buffer.from(ref)), label).toBe(true);
    return;
  }
  const { worst, mean } = rgbaDistance(ours, ref);
  expect(worst, `${label}: worst ${worst}`).toBeLessThanOrEqual(UPSAMPLING_WORST);
  expect(mean, `${label}: mean ${mean}`).toBeLessThanOrEqual(UPSAMPLING_MEAN);
};

const refusal = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(ReaderError);
    expect((e as ReaderError).code).toBe("EXTRACT_FAILED");
    return (e as ReaderError).message;
  }
  throw new Error("expected a refusal");
};

// --- A DC-only baseline encoder, for samplings jpeg-js cannot write ----------

const u16 = (n: number): number[] => [(n >> 8) & 0xff, n & 0xff];
const seg = (marker: number, payload: number[]): number[] => [0xff, marker, ...u16(payload.length + 2), ...payload];
/** T.81 table K.3 (luminance DC) code of each category 0..11. */
const K3_CODES = [
  "00",
  "010",
  "011",
  "100",
  "101",
  "110",
  "1110",
  "11110",
  "111110",
  "1111110",
  "11111110",
  "111111110",
];
const K3_COUNTS = [0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0];

interface DcSpec {
  width: number;
  height: number;
  /** [id, H, V] of each component. */
  components: [number, number, number][];
  adobe?: number;
  /** The DC (quantisation 1) of block (row, col) of component `ci`. */
  dc: (ci: number, row: number, col: number) => number;
}

/** A baseline stream whose blocks carry only a DC (AC table: EOB alone),
 *  every component interleaved, quantisation 1. */
const dcOnlyJpeg = (spec: DcSpec): Uint8Array => {
  const maxH = Math.max(...spec.components.map((c) => c[1]));
  const maxV = Math.max(...spec.components.map((c) => c[2]));
  const mcusX = Math.ceil(spec.width / (8 * maxH));
  const mcusY = Math.ceil(spec.height / (8 * maxV));
  let bits = "";
  const pred = spec.components.map(() => 0);
  for (let my = 0; my < mcusY; my++) {
    for (let mx = 0; mx < mcusX; mx++) {
      spec.components.forEach(([, h, v], ci) => {
        for (let j = 0; j < v; j++) {
          for (let i = 0; i < h; i++) {
            const value = spec.dc(ci, my * v + j, mx * h + i);
            const diff = value - pred[ci]!;
            pred[ci] = value;
            const category = diff === 0 ? 0 : Math.floor(Math.log2(Math.abs(diff))) + 1;
            const extra = diff >= 0 ? diff : diff + (1 << category) - 1;
            bits += K3_CODES[category]! + (category ? extra.toString(2).padStart(category, "0") : "") + "0";
          }
        }
      });
    }
  }
  bits += "1".repeat((8 - (bits.length % 8)) % 8);
  const data: number[] = [];
  for (const byte of bits.match(/.{8}/g)!.map((b) => parseInt(b, 2)))
    data.push(...(byte === 0xff ? [0xff, 0x00] : [byte]));
  const adobe =
    spec.adobe === undefined ? [] : seg(0xee, [0x41, 0x64, 0x6f, 0x62, 0x65, 0x00, 0x64, 0, 0, 0, 0, spec.adobe]);
  return new Uint8Array([
    0xff,
    0xd8,
    ...adobe,
    ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
    ...seg(0xc0, [
      8,
      ...u16(spec.height),
      ...u16(spec.width),
      spec.components.length,
      ...spec.components.flatMap(([id, h, v]) => [id, (h << 4) | v, 0]),
    ]),
    ...seg(0xc4, [0x00, ...K3_COUNTS, ...K3_CODES.map((_, i) => i)]),
    ...seg(0xc4, [0x10, 1, ...new Array<number>(15).fill(0), 0x00]),
    ...seg(0xda, [spec.components.length, ...spec.components.flatMap(([id]) => [id, 0x00]), 0, 63, 0]),
    ...data,
    0xff,
    0xd9,
  ]);
};

/** A pattern of DCs that moves every 8 × 8 block and every component: samples 128 ± 120. */
const pattern = (ci: number, row: number, col: number): number => (((ci * 7 + row * 3 + col * 5) % 31) - 15) * 64;

const noiseRgba = (w: number, h: number, seed: number): Uint8Array => {
  const out = new Uint8Array(w * h * 4);
  let s = seed;
  for (let i = 0; i < out.length; i++) {
    s = (Math.imul(s, 1_103_515_245) + 12_345) >>> 0;
    out[i] = i % 4 === 3 ? 255 : s >>> 24;
  }
  return out;
};

// --- Tests ---------------------------------------------------------------------

describe("decodeJpegRgba against jpeg-js, pixel for pixel without subsampling (tolerance 0), bounded with it", () => {
  it("every byte of the RGBA of the 27 sequential JPEGs of the fixtures (the 3 unsubsampled ones identical), the others bounded", async () => {
    const sequential = (await fixtureJpegs()).filter(
      (f) => parseJpeg(f.bytes, { maxPixels: 12e6 }).frame.process !== "progressive",
    );
    // Counts re-measured on the public documents (25/11/2026).
    expect(sequential.length).toBe(27);
    let identical = 0;
    for (const { id, bytes } of sequential) {
      const ours = decodeJpegRgba(bytes, GUARD);
      const ref = reference(bytes);
      expect([ours.width, ours.height], id).toEqual([ref.width, ref.height]);
      const subsampled = isSubsampled(parseJpeg(bytes, { maxPixels: 12e6 }).frame);
      if (!subsampled) identical++;
      sameAsJpegJs(ours.data, ref.data, subsampled, id);
    }
    expect(identical).toBe(3);
  }, 300_000);

  it("JPEGs written by jpeg-js's encoder, sizes not multiple of 8, from 1 × 1 up", () => {
    for (const [w, h] of [
      [1, 1],
      [17, 9],
      [33, 65],
      [64, 8],
      [7, 300],
    ] as const) {
      const bytes = new Uint8Array(encodeJpegJs({ width: w, height: h, data: noiseRgba(w, h, w * 31 + h) }, 75).data);
      // jpeg-js's encoder writes 4:2:0: the chroma step is libjpeg's here, jpeg-js's replication there.
      sameAsJpegJs(decodeJpegRgba(bytes, GUARD).data, reference(bytes).data, true, `${w}x${h}`);
    }
  });

  it("samplings jpeg-js cannot write: 4:2:2, 4:4:0, 4:2:0, 4:1:1, 3 × 1, 1 × 3, 2 × 2 on chroma, grey (replication kept, identical, except the 2:1 cases)", () => {
    const cases: [string, [number, number, number][]][] = [
      [
        "4:2:2",
        [
          [1, 2, 1],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "4:4:0",
        [
          [1, 1, 2],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "4:2:0",
        [
          [1, 2, 2],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "4:1:1",
        [
          [1, 4, 1],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "3x1",
        [
          [1, 3, 1],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "1x3",
        [
          [1, 1, 3],
          [2, 1, 1],
          [3, 1, 1],
        ],
      ],
      [
        "chroma 2x2",
        [
          [1, 2, 2],
          [2, 2, 2],
          [3, 1, 1],
        ],
      ],
      ["grey", [[1, 1, 1]]],
    ];
    for (const [name, components] of cases) {
      for (const [width, height] of [
        [50, 37],
        [8, 8],
        [97, 13],
      ] as const) {
        const bytes = dcOnlyJpeg({ width, height, components, dc: pattern });
        const ours = decodeJpegRgba(bytes, GUARD);
        // 4:2:2, 4:2:0 and 2 × 2 luma over 1 × 1 chroma (a 2:1 chroma too) take libjpeg's fancy filter; every other ratio keeps jpeg-js's replication, byte for byte.
        const fancy = name === "4:2:2" || name === "4:2:0" || name === "chroma 2x2";
        if (!fancy)
          expect(Buffer.from(ours.data).equals(Buffer.from(reference(bytes).data)), `${name} ${width}x${height}`).toBe(
            true,
          );
        else sameAsJpegJs(ours.data, reference(bytes).data, true, `${name} ${width}x${height}`);
      }
    }
  });

  it("CMYK (APP14 transform 0) and YCCK (transform 2), four components, as jpeg-js turns them into RGB (identical without subsampling, bounded with)", () => {
    const four: [number, number, number][] = [
      [1, 1, 1],
      [2, 1, 1],
      [3, 1, 1],
      [4, 1, 1],
    ];
    for (const adobe of [0, 2]) {
      for (const components of [
        four,
        [
          [1, 2, 2],
          [2, 1, 1],
          [3, 1, 1],
          [4, 2, 2],
        ] as [number, number, number][],
      ]) {
        const bytes = dcOnlyJpeg({ width: 41, height: 23, components, adobe, dc: pattern });
        sameAsJpegJs(decodeJpegRgba(bytes, GUARD).data, reference(bytes).data, components !== four, `adobe ${adobe}`);
      }
    }
  });

  it("three components with APP14 transform 1 or 2 are YCbCr, as in jpeg-js", () => {
    for (const adobe of [1, 2]) {
      const bytes = dcOnlyJpeg({
        width: 24,
        height: 16,
        components: [
          [1, 1, 1],
          [2, 1, 1],
          [3, 1, 1],
        ],
        adobe,
        dc: pattern,
      });
      expect(Buffer.from(decodeJpegRgba(bytes, GUARD).data).equals(Buffer.from(reference(bytes).data))).toBe(true);
    }
  });
});

// Inventory of what we took from jpeg-js as is (06/10/2026), point 3: the CMYK conversion was copied from jpeg-js with no
// source (IJG renders CMYK as is and leaves RGB to the application). Two independent readings must agree: ours and Pillow's,
// on a real Photoshop CMYK JPEG (fixtures/jpeg/README.md). Measured: mean 0.3, at most 3 levels of 255 apart.
describe("a real Adobe CMYK JPEG, against a second, independent reading", () => {
  it("decodes within 3 levels of Pillow's reading, on every channel of every pixel", () => {
    const jpeg = new Uint8Array(readFileSync(join(FIXTURES_JPEG, "pillow-sample-cmyk.jpg")));
    const pillow = new Uint8Array(readFileSync(join(FIXTURES_JPEG, "pillow-sample-cmyk.pillow-rgb.bin")));
    const ours = decodeJpegRgba(jpeg, GUARD);
    expect([ours.width, ours.height]).toEqual([100, 100]);
    let worst = 0;
    for (let i = 0; i < ours.width * ours.height; i++) {
      for (let c = 0; c < 3; c++) worst = Math.max(worst, Math.abs(ours.data[i * 4 + c]! - pillow[i * 3 + c]!));
    }
    expect(worst).toBeLessThanOrEqual(3);
  });
});

describe("decodeJpegRgba refuses, with our own message", () => {
  it("three components with APP14 transform 0: RGB by the Adobe reading, YCbCr by jpeg-js's — two readings disagree, we abstain", () => {
    const bytes = dcOnlyJpeg({
      width: 16,
      height: 16,
      components: [
        [1, 1, 1],
        [2, 1, 1],
        [3, 1, 1],
      ],
      adobe: 0,
      dc: pattern,
    });
    expect(() => reference(bytes)).not.toThrow();
    expect(refusal(() => decodeJpegRgba(bytes, GUARD))).toBe("jpeg: adobe rgb colour transform not supported");
  });

  it("an image over the pixel guard, like jpeg-js", () => {
    const bytes = dcOnlyJpeg({ width: 4000, height: 3001, components: [[1, 1, 1]], dc: () => 0 });
    expect(() => reference(bytes)).toThrow();
    expect(refusal(() => decodeJpegRgba(bytes, GUARD))).toBe("jpeg: image over the pixel guard");
  });

  it("an image over the memory guard; and under 256 MB a 12 MP colour JPEG passes, as with jpeg-js", () => {
    const bytes = dcOnlyJpeg({
      width: 4000,
      height: 3000,
      components: [
        [1, 1, 1],
        [2, 1, 1],
        [3, 1, 1],
      ],
      dc: () => 0,
    });
    expect(decodeJpegRgba(bytes, GUARD).data.length).toBe(4000 * 3000 * 4);
    expect(() => reference(bytes)).not.toThrow();
    // We count coefficients, planes and RGBA (≈ 228 MB here); jpeg-js counts more (an extra buffer): what it accepts, we accept.
    expect(() => reference(bytes, { maxMemoryUsageInMB: 100 })).toThrow();
    expect(refusal(() => decodeJpegRgba(bytes, { maxResolutionInMP: 12, maxMemoryUsageInMB: 100 }))).toBe(
      "jpeg: image over the memory guard",
    );
  }, 60_000);

  it("a fixture cut in the middle of its scan, and one cut before its frame, like jpeg-js", async () => {
    const [first] = await fixtureJpegs();
    for (const length of [Math.floor(first!.bytes.length / 2), 20]) {
      const cut = first!.bytes.subarray(0, length);
      expect(() => reference(cut)).toThrow();
      const message = refusal(() => decodeJpegRgba(cut, GUARD));
      expect(message).toMatch(/^jpeg: [a-z ]+$/);
    }
  });

  // Review of 06/10/2026: refinement scans that only carry end-of-band runs walk every coefficient of every
  // block for a few bytes each; 255 of them on 12 MP took 7.5 s. The work of the progressive scans is summed from
  // their headers and refused BEFORE any scan is decoded: these scans carry no data at all, so decoding would
  // first refuse "truncated scan data". Fixtures: 1.03 to 2.66 times the image's coefficients; this stream: 9.9.
  it("🔴 progressive scans that would walk more than 8 times the image's coefficients, before decoding any", () => {
    const one = [0, 1, ...new Array<number>(14).fill(0)];
    const head = [
      0xff,
      0xd8,
      ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
      ...seg(0xc2, [8, 0, 64, 0, 64, 1, 1, 0x11, 0]),
      ...seg(0xc4, [0x00, ...one, 0]),
      ...seg(0xc4, [0x10, ...one, 0]),
    ];
    const scan = (ss: number, se: number, ahal: number): number[] => seg(0xda, [1, 1, 0x00, ss, se, ahal]);
    const refinements = Array.from({ length: 9 }, () => scan(1, 63, 0x10)).flat();
    const bytes = new Uint8Array([...head, ...scan(0, 0, 0x00), ...scan(1, 63, 0x01), ...refinements, 0xff, 0xd9]);
    expect(refusal(() => decodeJpegRgba(bytes, GUARD))).toBe("jpeg: scans over the work guard");
  });

  // Counter-review of 06/10/2026: a DC scan was counted as 1 coefficient per block, far below its real cost
  // (255 DC refinements on 12 MP, 4:4:4: 3 s under the guard, from 18 MB). Every scan now counts at least 8 per block.
  it("🔴 70 DC refinement scans count over the guard, though they walk one coefficient per block", () => {
    const one = [0, 1, ...new Array<number>(14).fill(0)];
    const head = [
      0xff,
      0xd8,
      ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
      ...seg(0xc2, [8, 0, 64, 0, 64, 1, 1, 0x11, 0]),
      ...seg(0xc4, [0x00, ...one, 0]),
      ...seg(0xc4, [0x10, ...one, 0]),
    ];
    const scan = (ahal: number): number[] => seg(0xda, [1, 1, 0x00, 0, 0, ahal]);
    const refinements = Array.from({ length: 70 }, () => scan(0x10)).flat();
    const bytes = new Uint8Array([...head, ...scan(0x01), ...refinements, 0xff, 0xd9]);
    expect(refusal(() => decodeJpegRgba(bytes, GUARD))).toBe("jpeg: scans over the work guard");
  });

  it("a caller's invalid guard is a bug: plain Error", () => {
    const bytes = dcOnlyJpeg({ width: 8, height: 8, components: [[1, 1, 1]], dc: () => 0 });
    expect(() => decodeJpegRgba(bytes, { maxResolutionInMP: Number.NaN, maxMemoryUsageInMB: 1 })).toThrow(
      /^decodeJpegRgba: /,
    );
    expect(() => decodeJpegRgba(bytes, { maxResolutionInMP: 1, maxMemoryUsageInMB: -1 })).toThrow(/^decodeJpegRgba: /);
  });
});
