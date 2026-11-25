import { decode as decodeJpegJs, encode as encodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { fixtureJpegs } from "./jpeg-fixtures.test.support";
import { parseJpeg, type JpegStructure } from "./markers";
import { MemoryBudget } from "./memory-budget";
import { allocateCoefficients, decodeSequentialScan, type ComponentCoefficients } from "./sequential-scan";

const GUARD = { maxPixels: 12_000_000 };
const MB = 1024 * 1024;

const u16 = (n: number): number[] => [(n >> 8) & 0xff, n & 0xff];
const seg = (marker: number, payload: number[]): number[] => [0xff, marker, ...u16(payload.length + 2), ...payload];
const dht = (tc: number, counts: number[], values: number[]): number[] =>
  seg(0xc4, [tc << 4, ...counts, ...new Array<number>(16 - counts.length).fill(0), ...values]);
/** A grey baseline stream of `w × h` with one-code tables: DC category
 *  `dcValue` (code "0"), AC symbol `acValue` (code "0"), and `data` as the
 *  entropy-coded segment. */
const greyStream = (w: number, h: number, dcValue: number, acValue: number, data: number[], dri = 0): Uint8Array =>
  new Uint8Array([
    0xff,
    0xd8,
    ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
    ...seg(0xc0, [8, ...u16(h), ...u16(w), 1, 1, 0x11, 0]),
    ...dht(0, [1], [dcValue]),
    ...dht(1, [1], [acValue]),
    ...(dri ? seg(0xdd, u16(dri)) : []),
    ...seg(0xda, [1, 1, 0x00, 0, 63, 0]),
    ...data,
    0xff,
    0xd9,
  ]);
const bits = (s: string): number[] =>
  (s + "1".repeat((8 - (s.length % 8)) % 8)).match(/.{8}/g)!.map((b) => parseInt(b, 2));

const decodeAll = (s: JpegStructure, budget = new MemoryBudget(256 * MB)): ComponentCoefficients[] => {
  const coefficients = allocateCoefficients(s.frame, budget);
  for (const scan of s.scans) decodeSequentialScan(s.frame, scan, coefficients);
  return coefficients;
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

/** The natural-order DC coefficient of block (row, col). */
const dcOf = (c: ComponentCoefficients, row: number, col: number): number => c.data[(row * c.stride + col) * 64]!;

describe("MemoryBudget", () => {
  it("counts what it allocates, and refuses BEFORE the allocation that would pass the guard", () => {
    const budget = new MemoryBudget(1000);
    expect(budget.int32(200).length).toBe(200);
    expect(budget.spent).toBe(800);
    expect(refusal(() => budget.uint8(201))).toBe("jpeg: image over the memory guard");
    expect(budget.spent).toBe(800);
    expect(budget.uint8(200).length).toBe(200);
    expect(budget.spent).toBe(1000);
  });

  it("a caller's invalid guard is a bug: plain Error", () => {
    expect(() => new MemoryBudget(Number.NaN)).toThrow(/^MemoryBudget: /);
  });
});

describe("allocateCoefficients", () => {
  it("allocates the MCU-padded block grid of each component (T.81 A.2.4), 4 bytes a coefficient, charged to the budget", () => {
    const s = parseJpeg(
      new Uint8Array([
        0xff,
        0xd8,
        ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
        ...seg(0xc0, [8, ...u16(17), ...u16(33), 3, 1, 0x22, 0, 2, 0x11, 0, 3, 0x11, 0]),
        ...dht(0, [1], [0]),
        ...dht(1, [1], [0]),
        ...seg(0xda, [3, 1, 0, 2, 0, 3, 0, 0, 63, 0]),
        0x00,
        0xff,
        0xd9,
      ]),
      GUARD,
    );
    const budget = new MemoryBudget(MB);
    const [y, cb] = allocateCoefficients(s.frame, budget);
    // 33 × 17 at 2×2: 3 × 2 MCUs; Y has 6 × 4 blocks allocated, 5 × 3 inside the image.
    expect([y!.stride, y!.rows, y!.blocksPerLine, y!.blocksPerColumn]).toEqual([6, 4, 5, 3]);
    expect([cb!.stride, cb!.rows, cb!.blocksPerLine, cb!.blocksPerColumn]).toEqual([3, 2, 3, 2]);
    expect(budget.spent).toBe((24 + 6 + 6) * 64 * 4);
  });

  it("refuses a frame whose coefficients would pass the memory guard, before allocating them", () => {
    const s = parseJpeg(greyStream(4000, 3000, 0, 0, [0x00]), GUARD);
    const budget = new MemoryBudget(10 * MB);
    expect(refusal(() => allocateCoefficients(s.frame, budget))).toBe("jpeg: image over the memory guard");
    expect(budget.spent).toBe(0);
  });
});

describe("decodeSequentialScan on synthetic streams", () => {
  it("decodes DC differences and keeps the prediction across blocks (T.81 F.2.1.3.1)", () => {
    // Two blocks, DC category 1 then bit "1" (+1), AC EOB: DC 1, then 2.
    const s = parseJpeg(greyStream(16, 8, 1, 0x00, bits("010" + "010")), GUARD);
    const [c] = decodeAll(s);
    expect([dcOf(c!, 0, 0), dcOf(c!, 0, 1)]).toEqual([1, 2]);
  });

  it("resets the prediction at each restart marker, which must come in sequence (T.81 E.2.4, F.2.2.5)", () => {
    const one = bits("010");
    const s = parseJpeg(greyStream(24, 8, 1, 0x00, [...one, 0xff, 0xd0, ...one, 0xff, 0xd1, ...one], 1), GUARD);
    const [c] = decodeAll(s);
    expect([dcOf(c!, 0, 0), dcOf(c!, 0, 1), dcOf(c!, 0, 2)]).toEqual([1, 1, 1]);
    const swapped = parseJpeg(greyStream(24, 8, 1, 0x00, [...one, 0xff, 0xd1, ...one, 0xff, 0xd0, ...one], 1), GUARD);
    expect(refusal(() => decodeAll(swapped))).toBe("jpeg: bad restart marker");
  });

  it("places AC coefficients in natural order through the zigzag (T.81 figure A.6)", () => {
    // DC cat 0; AC symbol 0x11 (run 1, size 1) "0" + bit "1": zigzag index 2 = natural 8, then zigzag 4 = natural 9, then EOB.
    const s = parseJpeg(
      new Uint8Array([
        0xff,
        0xd8,
        ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
        ...seg(0xc0, [8, 0, 8, 0, 8, 1, 1, 0x11, 0]),
        ...dht(0, [1], [0]),
        ...dht(1, [1, 1], [0x11, 0x00]),
        ...seg(0xda, [1, 1, 0, 0, 63, 0]),
        ...bits("0" + "01" + "01" + "10"),
        0xff,
        0xd9,
      ]),
      GUARD,
    );
    const [c] = decodeAll(s);
    const nonZero = [...c!.data.subarray(0, 64)].flatMap((v, i) => (v ? [[i, v]] : []));
    expect(nonZero).toEqual([
      [8, 1],
      [9, 1],
    ]);
  });

  it("refuses a coefficient past the 64th (run lengths that overflow the block)", () => {
    // AC symbol 0xF1: run 15, size 1; the fourth lands at index 64.
    const s = parseJpeg(greyStream(8, 8, 0, 0xf1, bits("0" + "01".repeat(4))), GUARD);
    expect(refusal(() => decodeAll(s))).toBe("jpeg: coefficient past the block");
  });

  it("refuses a DC category over 11 and an AC size over 10 (T.81 tables F.1 and F.2, 8-bit samples)", () => {
    expect(
      refusal(() => decodeAll(parseJpeg(greyStream(8, 8, 12, 0x00, bits("0" + "1".repeat(12) + "0")), GUARD))),
    ).toBe("jpeg: bad dc category");
    expect(
      refusal(() => decodeAll(parseJpeg(greyStream(8, 8, 0, 0x0b, bits("0" + "0" + "1".repeat(11))), GUARD))),
    ).toBe("jpeg: bad ac size");
  });

  it("refuses a segment that ends inside a block, or a restart marker where none is due", () => {
    expect(refusal(() => decodeAll(parseJpeg(greyStream(64, 64, 1, 0x00, bits("010")), GUARD)))).toBe(
      "jpeg: truncated scan data",
    );
    const one = bits("010");
    expect(refusal(() => decodeAll(parseJpeg(greyStream(16, 8, 1, 0x00, [...one, 0xff, 0xd0, ...one]), GUARD)))).toBe(
      "jpeg: unexpected marker in scan data",
    );
  });

  it("decodes a uniform image encoded by jpeg-js to blocks with no AC and the same DC", () => {
    const data = new Uint8Array(32 * 24 * 4).map((_, i) => (i % 4 === 3 ? 255 : 100));
    const s = parseJpeg(new Uint8Array(encodeJpegJs({ width: 32, height: 24, data }, 90).data), GUARD);
    const coefficients = decodeAll(s);
    for (const c of coefficients) {
      const dcs = new Set<number>();
      for (let block = 0; block < c.blocksPerLine * c.blocksPerColumn; block++) {
        const at = (Math.floor(block / c.blocksPerLine) * c.stride + (block % c.blocksPerLine)) * 64;
        dcs.add(c.data[at]!);
        expect(c.data.subarray(at + 1, at + 64).every((v) => v === 0)).toBe(true);
      }
      expect(dcs.size).toBe(1);
    }
  });
});

describe("decodeSequentialScan on every sequential JPEG of the fixtures", () => {
  /**
   * Cross-check against jpeg-js without its internals: with
   * `colorTransform: false` jpeg-js writes each component's samples, upsampled
   * by replication, as R, G, B. The mean of a block's samples is exactly
   * 128 + DC × Q0 / 8 in T.81's real IDCT (A.3.3: the AC basis functions sum
   * to 0 over a block), so every DC we decode must agree with the mean of
   * the region jpeg-js drew for that block, within 1 (integer IDCT rounding),
   * on every block whose region no sample clipped to 0 or 255. A desynchronised
   * AC decode would shift every later DC.
   */
  it("decodes every block of the 27 sequential JPEGs, each DC agreeing with jpeg-js's block mean", async () => {
    const sequential = (await fixtureJpegs())
      .map((f) => ({ ...f, s: parseJpeg(f.bytes, GUARD) }))
      .filter((f) => f.s.frame.process !== "progressive");
    // Count re-measured on the public documents (25/11/2026).
    expect(sequential.length).toBe(27);
    let checked = 0;
    for (const { id, bytes, s } of sequential) {
      const coefficients = decodeAll(s);
      // An APP14 transform forces jpeg-js to convert colours: its channels are not the components.
      if (s.adobe?.transform) continue;
      const ref = decodeJpegJs(bytes, {
        useTArray: true,
        formatAsRGBA: true,
        colorTransform: false,
        maxResolutionInMP: 12,
        maxMemoryUsageInMB: 256,
      });
      const { frame } = s;
      frame.components.forEach((component, ci) => {
        const c = coefficients[ci]!;
        const q0 = s.quantTables[component.tq]![0]!;
        const sx = frame.maxH / component.h;
        const sy = frame.maxV / component.v;
        for (let row = 0; row < c.blocksPerColumn; row++) {
          for (let col = 0; col < c.blocksPerLine; col++) {
            const x0 = col * 8 * sx;
            const y0 = row * 8 * sy;
            if (x0 + 8 * sx > frame.width || y0 + 8 * sy > frame.height) continue;
            let sum = 0;
            let clipped = false;
            for (let y = y0; y < y0 + 8 * sy; y++) {
              for (let x = x0; x < x0 + 8 * sx; x++) {
                const v = ref.data[(y * frame.width + x) * 4 + (frame.components.length === 1 ? 0 : ci)]!;
                clipped ||= v === 0 || v === 255;
                sum += v;
              }
            }
            if (clipped) continue;
            const mean = sum / (64 * sx * sy);
            expect(
              Math.abs(mean - (128 + (dcOf(c, row, col) * q0) / 8)),
              `${id} component ${ci} block ${row},${col}`,
            ).toBeLessThanOrEqual(1);
            checked++;
          }
        }
      });
    }
    // Measured on the fixtures: the check is not vacuous.
    expect(checked).toBeGreaterThan(100_000);
  }, 300_000);
});
