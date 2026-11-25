import { createHash } from "node:crypto";

import { decode as decodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { componentPlanes, inverseDctBlock, naturalQuantTable } from "./idct";
import { fixtureJpegs } from "./jpeg-fixtures.test.support";
import { parseJpeg, type JpegStructure } from "./markers";
import { MemoryBudget } from "./memory-budget";
import { allocateCoefficients, decodeSequentialScan, ZIGZAG } from "./sequential-scan";

const GUARD = { maxPixels: 12_000_000 };
const MB = 1024 * 1024;

/** T.81 A.3.3: the real inverse DCT of one block (natural order, already
 *  dequantised), level-shifted by 128 (A.3.1), rounded and clamped to 0..255. */
const referenceIdct = (block: Float64Array): number[] => {
  const out: number[] = [];
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      let sum = 0;
      for (let v = 0; v < 8; v++) {
        for (let u = 0; u < 8; u++) {
          const cu = u === 0 ? Math.SQRT1_2 : 1;
          const cv = v === 0 ? Math.SQRT1_2 : 1;
          sum +=
            cu *
            cv *
            block[v * 8 + u]! *
            Math.cos(((2 * x + 1) * u * Math.PI) / 16) *
            Math.cos(((2 * y + 1) * v * Math.PI) / 16);
        }
      }
      out.push(Math.min(255, Math.max(0, Math.round(sum / 4 + 128))));
    }
  }
  return out;
};

/** Deterministic pseudo-random integers (an LCG), like the repository's tests. */
const lcg = (seed: number) => {
  let s = seed >>> 0;
  return (): number => {
    s = (Math.imul(s, 1_103_515_245) + 12_345) >>> 0;
    return s / 2 ** 32;
  };
};

/** One call shape for every test below. */
const idct = (
  coefficients: Int32Array,
  at: number,
  quant: Int32Array,
  out: Uint8Array,
  outAt: number,
  stride: number,
): void => inverseDctBlock(coefficients, at, quant, { data: out, width: stride }, outAt);

const decodeCoefficients = (s: JpegStructure) => {
  const budget = new MemoryBudget(256 * MB);
  const coefficients = allocateCoefficients(s.frame, budget);
  for (const scan of s.scans) decodeSequentialScan(s.frame, scan, coefficients);
  return { coefficients, budget };
};

describe("naturalQuantTable", () => {
  it("puts the zigzag-ordered DQT values in natural order (T.81 B.2.4.1, figure A.6)", () => {
    const zigzag = new Uint16Array(64).map((_, k) => k + 1);
    const natural = naturalQuantTable(zigzag);
    for (let k = 0; k < 64; k++) expect(natural[ZIGZAG[k]!]).toBe(k + 1);
  });
});

describe("inverseDctBlock", () => {
  it("gives a flat block for a DC-only block: 128 + DC × Q / 8", () => {
    const out = new Uint8Array(64);
    const coefficients = new Int32Array(64);
    coefficients[0] = 40;
    idct(coefficients, 0, new Int32Array(64).fill(2), out, 0, 8);
    expect(new Set(out)).toEqual(new Set([138]));
  });

  it("writes the block at its place in a wider plane, row by row", () => {
    const out = new Uint8Array(16 * 8);
    const coefficients = new Int32Array(64);
    coefficients[0] = -1024;
    idct(coefficients, 0, new Int32Array(64).fill(1), out, 8, 16);
    for (let y = 0; y < 8; y++) {
      expect([...out.subarray(y * 16, y * 16 + 8)].every((v) => v === 0)).toBe(true);
      expect([...out.subarray(y * 16 + 8, y * 16 + 16)].every((v) => v === 0)).toBe(true);
    }
    coefficients[0] = 1024;
    idct(coefficients, 0, new Int32Array(64).fill(1), out, 8, 16);
    expect(out[8]).toBe(255);
    expect(out[0]).toBe(0);
  });

  it("stays within 1 of the real IDCT of T.81 A.3.3 on 2 000 pseudo-random blocks (stated tolerance)", () => {
    const random = lcg(7);
    const quant = new Int32Array(64).fill(1);
    let worst = 0;
    for (let n = 0; n < 2000; n++) {
      const coefficients = new Int32Array(64);
      // Magnitudes of a photograph: a DC anywhere, ACs shrinking with frequency.
      coefficients[0] = Math.round((random() - 0.5) * 1600);
      for (let i = 1; i < 64; i++)
        if (random() < 0.3) coefficients[i] = Math.round((random() - 0.5) * (400 / (1 + (i >> 3) + (i & 7))));
      const out = new Uint8Array(64);
      idct(coefficients, 0, quant, out, 0, 8);
      const ref = referenceIdct(Float64Array.from(coefficients));
      for (let i = 0; i < 64; i++) worst = Math.max(worst, Math.abs(out[i]! - ref[i]!));
    }
    expect(worst).toBeLessThanOrEqual(1);
  });
});

/**
 * Characterisation of the exact arithmetic (tranche 10C4, before `inverseDctBlock` was cut into its four steps).
 * A battery of some 20 000 blocks, deterministic, hashed over the whole target plane after each block so that a byte
 * written outside the 8 × 8 window, a scratch value leaking from the previous block, or one rounding off by one
 * changes the digest. It covers: every single AC coefficient at every position (the zero-row and zero-column
 * shortcuts and the full butterflies), a DC alone, a first row alone, a first column alone, dense photographic
 * blocks, extreme coefficients and quantisers (±2^15 × 65535, which wrap at 32 bits as jpeg-js's do), and zeros.
 */
describe("inverseDctBlock: the exact samples of a fixed battery", () => {
  const DIGEST = "ebfac304ddca1241f9747220d74e18290d41d09f967a43f1b6379c1e0ece1310";

  it("keeps the digest of the battery, window bytes and untouched bytes included", () => {
    const random = lcg(20261015);
    const hash = createHash("sha256");
    const STRIDE = 21;
    const AT = 7;
    const plane = new Uint8Array(AT + 8 * STRIDE + 9).fill(0xaa);
    const coefficients = new Int32Array(64 * 4);
    const quants = [
      new Int32Array(64).fill(1),
      new Int32Array(64).fill(65535),
      Int32Array.from({ length: 64 }, (_, i) => 1 + ((i * 37) % 255)),
    ];
    const run = (block: Int32Array, quant: Int32Array): void => {
      coefficients.fill(0x5a5a);
      coefficients.set(block, 64 * 2);
      idct(coefficients, 64 * 2, quant, plane, AT, STRIDE);
      hash.update(plane);
    };
    const one = (i: number, v: number): Int32Array => {
      const b = new Int32Array(64);
      b[i] = v;
      return b;
    };
    for (const quant of quants) {
      for (let i = 0; i < 64; i++) for (const v of [1, -1, 7, -300, 1023, -32768, 32767]) run(one(i, v), quant);
      for (let v = -2100; v <= 2100; v += 23) run(one(0, v), quant);
      run(new Int32Array(64), quant);
      for (let n = 0; n < 1500; n++) {
        const row = new Int32Array(64);
        for (let x = 0; x < 8; x++) row[x] = Math.round((random() - 0.5) * 2000);
        run(row, quant);
        const col = new Int32Array(64);
        for (let y = 0; y < 8; y++) col[y * 8] = Math.round((random() - 0.5) * 2000);
        run(col, quant);
        const dense = new Int32Array(64);
        dense[0] = Math.round((random() - 0.5) * 1600);
        for (let i = 1; i < 64; i++)
          if (random() < 0.4) dense[i] = Math.round((random() - 0.5) * (600 / (1 + (i >> 3) + (i & 7))));
        run(dense, quant);
        const wild = new Int32Array(64);
        for (let i = 0; i < 64; i++)
          wild[i] = random() < 0.5 ? (random() < 0.5 ? -32768 : 32767) : Math.round((random() - 0.5) * 2 ** 20);
        run(wild, quant);
      }
    }
    expect(hash.digest("hex")).toBe(DIGEST);
  });
});

describe("inverseDctBlock: a column left with only its first value (hostile first row)", () => {
  it("rounds that value as +8192 >> 14 does, also when 32-bit wrapping lands it on a sample boundary", () => {
    // Found by search: with +8191 instead of +8192 the third-from-left sample would be 35, not 36.
    const coefficients = Int32Array.from([0, 614918, 127905, 0, 1029418, -1031024, -392080, 2084303]);
    const out = new Uint8Array(64);
    idct(
      Int32Array.from([...coefficients, ...new Array<number>(56).fill(0)]),
      0,
      new Int32Array(64).fill(1),
      out,
      0,
      8,
    );
    for (let y = 0; y < 8; y++) expect([...out.subarray(y * 8, y * 8 + 8)]).toEqual([0, 0, 255, 36, 255, 255, 0, 255]);
  });
});

describe("componentPlanes", () => {
  it("allocates each plane through the budget (blocks inside the image × 8), and refuses past it", () => {
    const bytes = new Uint8Array([
      0xff,
      0xd8,
      0xff,
      0xdb,
      0,
      67,
      0,
      ...new Array<number>(64).fill(1),
      0xff,
      0xc0,
      0,
      11,
      8,
      0,
      9,
      0,
      17,
      1,
      1,
      0x11,
      0,
      0xff,
      0xc4,
      0,
      20,
      0x00,
      1,
      ...new Array<number>(15).fill(0),
      0,
      0xff,
      0xc4,
      0,
      20,
      0x10,
      1,
      ...new Array<number>(15).fill(0),
      0,
      0xff,
      0xda,
      0,
      8,
      1,
      1,
      0,
      0,
      63,
      0,
      0x00,
      0x0f,
      0xff,
      0xd9,
    ]);
    const s = parseJpeg(bytes, GUARD);
    const { coefficients, budget } = decodeCoefficients(s);
    const before = budget.spent;
    const [plane] = componentPlanes(s, coefficients, budget);
    expect([plane!.width, plane!.height]).toEqual([24, 16]);
    expect(budget.spent - before).toBe(24 * 16);
    expect(new Set(plane!.data)).toEqual(new Set([128]));
    const tight = new MemoryBudget(100);
    try {
      componentPlanes(s, coefficients, tight);
      throw new Error("expected a refusal");
    } catch (e) {
      expect(e).toBeInstanceOf(ReaderError);
      expect((e as ReaderError).message).toBe("jpeg: image over the memory guard");
    }
  });

  /**
   * Pixel for pixel: with `colorTransform: false` jpeg-js writes each
   * component's samples, upsampled by replication (pixel x reads sample
   * floor(x × H / Hmax)), as R, G, B. Every sample we compute must be the one
   * jpeg-js drew, exactly: same integer factorisation of the IDCT.
   */
  it("gives exactly jpeg-js's samples, every component of every pixel, on the 60 sequential JPEGs of the fixtures", async () => {
    const sequential = (await fixtureJpegs())
      .map((f) => ({ ...f, s: parseJpeg(f.bytes, GUARD) }))
      .filter((f) => f.s.frame.process !== "progressive");
    expect(sequential.length).toBe(61);
    let compared = 0;
    for (const { id, bytes, s } of sequential) {
      // An APP14 transform forces jpeg-js to convert colours: its channels are not the components (slice 4).
      if (s.adobe?.transform) continue;
      const { coefficients, budget } = decodeCoefficients(s);
      const planes = componentPlanes(s, coefficients, budget);
      const ref = decodeJpegJs(bytes, {
        useTArray: true,
        formatAsRGBA: true,
        colorTransform: false,
        maxResolutionInMP: 12,
        maxMemoryUsageInMB: 256,
      });
      const { frame } = s;
      let mismatches = 0;
      frame.components.forEach((component, ci) => {
        const plane = planes[ci]!;
        const channel = frame.components.length === 1 ? 0 : ci;
        for (let y = 0; y < frame.height; y++) {
          const row = Math.floor((y * component.v) / frame.maxV) * plane.width;
          for (let x = 0; x < frame.width; x++) {
            if (
              plane.data[row + Math.floor((x * component.h) / frame.maxH)] !==
              ref.data[(y * frame.width + x) * 4 + channel]
            )
              mismatches++;
          }
        }
        compared += frame.width * frame.height;
      });
      expect(mismatches, id).toBe(0);
    }
    expect(compared).toBeGreaterThan(10_000_000);
  }, 300_000);
});
