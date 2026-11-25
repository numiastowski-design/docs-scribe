import { decode as decodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { decodeJpegRgba } from "./decode";
import { fixtureJpegs, rgbaDistance } from "./jpeg-fixtures.test.support";
import { parseJpeg } from "./markers";

const CALLERS = {
  useTArray: true,
  formatAsRGBA: true,
  tolerantDecoding: false,
  maxResolutionInMP: 12,
  maxMemoryUsageInMB: 256,
} as const;
const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };

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

const u16 = (n: number): number[] => [(n >> 8) & 0xff, n & 0xff];
const seg = (marker: number, payload: number[]): number[] => [0xff, marker, ...u16(payload.length + 2), ...payload];
const bits = (s: string): number[] =>
  (s + "1".repeat((8 - (s.length % 8)) % 8)).match(/.{8}/g)!.map((b) => parseInt(b, 2));
/** One-code tables: DC category 0 ("0"), AC symbol `ac` ("0"). */
const tables = (ac: number): number[] => [
  ...seg(0xc4, [0x00, 1, ...new Array<number>(15).fill(0), 0]),
  ...seg(0xc4, [0x10, 1, ...new Array<number>(15).fill(0), ac]),
];
const YCC = [3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0];
/** A progressive 8 × 8 stream: `comps` (Nf then Ci, HV, Tq triples), then the given scans (header payload, data). */
const progressive = (comps: number[], ac: number, scans: [number[], number[]][]): Uint8Array =>
  new Uint8Array([
    0xff,
    0xd8,
    ...seg(0xdb, [0, ...new Array<number>(64).fill(1)]),
    ...seg(0xc2, [8, 0, 8, 0, 8, ...comps]),
    ...tables(ac),
    ...scans.flatMap(([header, data]) => [...seg(0xda, header), ...data]),
    0xff,
    0xd9,
  ]);
const GREY = [1, 1, 0x11, 0];
const DC_FIRST: [number[], number[]] = [[1, 1, 0x00, 0, 0, 0x01], bits("0")];

/** All 12 progressive fixtures are 4:2:0: since 06/10/2026 the chroma step is libjpeg's fancy filter (`color.ts`), not jpeg-js's
 *  replication, so pixels are bounded rather than equal. Measured against jpeg-js over the 12: worst level 67, worst mean 1.01;
 *  the exact chroma reading is checked against Pillow in `upsampling.test.ts`. Valeurs de départ à calibrer. */
const WORST = 75;
const MEAN = 1.5;

describe("decodeJpegRgba on progressive JPEGs, against jpeg-js (same size, bounded distance: 4:2:0)", () => {
  it("the RGBA of the 13 progressive JPEGs of the fixtures (the 13th, 4:4:4 Adobe, from docx-reel/11, added 06/10), whose scans exercise the four kinds of T.81 G.1.2", async () => {
    const progressives = (await fixtureJpegs())
      .map((f) => ({ ...f, s: parseJpeg(f.bytes, { maxPixels: 12e6 }) }))
      .filter((f) => f.s.frame.process === "progressive");
    expect(progressives.length).toBe(13);
    const kinds = new Set<string>();
    for (const { id, bytes, s } of progressives) {
      for (const scan of s.scans) kinds.add(`${scan.ss === 0 ? "dc" : "ac"} ${scan.ah === 0 ? "first" : "refine"}`);
      const ours = decodeJpegRgba(bytes, GUARD);
      const ref = decodeJpegJs(bytes, CALLERS);
      expect([ours.width, ours.height], id).toEqual([ref.width, ref.height]);
      const { worst, mean } = rgbaDistance(ours.data, ref.data);
      expect(worst, `${id}: worst ${worst}`).toBeLessThanOrEqual(WORST);
      expect(mean, `${id}: mean ${mean}`).toBeLessThanOrEqual(MEAN);
    }
    expect([...kinds].sort()).toEqual(["ac first", "ac refine", "dc first", "dc refine"]);
  }, 300_000);

  it("refuses a progressive fixture cut in the middle, like jpeg-js", async () => {
    const progressives = (await fixtureJpegs()).filter(
      (f) => parseJpeg(f.bytes, { maxPixels: 12e6 }).frame.process === "progressive",
    );
    const cut = progressives[0]!.bytes.subarray(0, Math.floor(progressives[0]!.bytes.length * 0.6));
    expect(() => decodeJpegJs(cut, CALLERS)).toThrow();
    expect(refusal(() => decodeJpegRgba(cut, GUARD))).toMatch(/^jpeg: [a-z ]+$/);
  });
});

describe("progressive scan parameters (T.81 B.2.3, G.1.1.1)", () => {
  it("accepts a DC-first scan then an AC-first scan of a grey 8 × 8 image", () => {
    const ok = progressive(GREY, 0x00, [DC_FIRST, [[1, 1, 0x00, 1, 63, 0x00], bits("0")]]);
    expect(decodeJpegRgba(ok, GUARD).data.length).toBe(256);
  });

  it("refuses a DC scan with Se ≠ 0, an AC scan over several components, Se < Ss, Se > 63", () => {
    const message = "jpeg: bad progressive scan parameters";
    expect(refusal(() => decodeJpegRgba(progressive(GREY, 0x00, [[[1, 1, 0x00, 0, 5, 0], bits("0")]]), GUARD))).toBe(
      message,
    );
    expect(
      refusal(() =>
        decodeJpegRgba(
          progressive(YCC, 0x00, [
            [[3, 1, 0, 2, 0, 3, 0, 0, 0, 0], bits("000")],
            [[2, 1, 0, 2, 0, 1, 63, 0], bits("00")],
          ]),
          GUARD,
        ),
      ),
    ).toBe(message);
    expect(
      refusal(() => decodeJpegRgba(progressive(GREY, 0x00, [DC_FIRST, [[1, 1, 0x00, 9, 3, 0], bits("0")]]), GUARD)),
    ).toBe(message);
    expect(
      refusal(() => decodeJpegRgba(progressive(GREY, 0x00, [DC_FIRST, [[1, 1, 0x00, 1, 64, 0], bits("0")]]), GUARD)),
    ).toBe(message);
  });

  it("refuses a refinement that is not one bit (Al ≠ Ah − 1), and Al over 13", () => {
    const message = "jpeg: bad progressive scan parameters";
    expect(
      refusal(() => decodeJpegRgba(progressive(GREY, 0x00, [DC_FIRST, [[1, 1, 0x00, 0, 0, 0x20], [0x80]]]), GUARD)),
    ).toBe(message);
    expect(refusal(() => decodeJpegRgba(progressive(GREY, 0x00, [[[1, 1, 0x00, 0, 0, 0x0e], bits("0")]]), GUARD))).toBe(
      message,
    );
  });

  it("refuses an AC coefficient that lands past the scan's band (jpeg-js writes it outside the band)", () => {
    // Band 1..5; symbol 0x71 = run 7, size 1: index 8.
    const bad = progressive(GREY, 0x71, [DC_FIRST, [[1, 1, 0x00, 1, 5, 0x00], bits("01")]]);
    expect(refusal(() => decodeJpegRgba(bad, GUARD))).toBe("jpeg: coefficient past the band");
  });

  it("refuses an AC refinement symbol of size other than 1 (T.81 G.1.2.3)", () => {
    const bad = progressive(GREY, 0x02, [DC_FIRST, [[1, 1, 0x00, 1, 63, 0x10], bits("011")]]);
    expect(refusal(() => decodeJpegRgba(bad, GUARD))).toBe("jpeg: bad ac refinement");
  });

  // Since the counter-review of 06/10/2026 a scan counts at least 8 coefficients per block (`decode.ts`): 256 DC
  // refinements now cross the work guard; 64 still pass. The 257th scan is refused while reading, first.
  it("bounds the number of scans: 257 are refused while reading, and 256 DC refinements cross the work guard (jpeg-js decodes them all)", () => {
    const refine: [number[], number[]] = [[1, 1, 0x00, 0, 0, 0x10], [0x7f]];
    const many = (n: number): Uint8Array =>
      progressive(GREY, 0x00, [DC_FIRST, ...new Array<[number[], number[]]>(n - 1).fill(refine)]);
    expect(decodeJpegRgba(many(64), GUARD).width).toBe(8);
    expect(refusal(() => decodeJpegRgba(many(256), GUARD))).toBe("jpeg: scans over the work guard");
    expect(() => decodeJpegJs(many(257), CALLERS)).not.toThrow();
    expect(refusal(() => decodeJpegRgba(many(257), GUARD))).toBe("jpeg: too many scans");
  });
});
