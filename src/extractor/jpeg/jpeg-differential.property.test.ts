/**
 * Slice 6 of the jpeg-js detachment: differential fuzz of our JPEG decoder
 * (`decodeJpegRgba`) against jpeg-js, with the options of the two product
 * callers (`image-reduction.ts`, `pptx-reader/picture.ts`).
 *
 * Damaged JPEGs (truncated, one bit flipped, one byte rewritten, a segment
 * length or a frame dimension rewritten — ITU-T T.81 B.1.1.4 segment length,
 * B.2.2 frame header) go through both decoders. The rules held:
 * - ours refuses only with our own `ReaderError("EXTRACT_FAILED", "jpeg: …")`;
 *   anything else (a TypeError, a RangeError) is a crash, and fails;
 * - ours NEVER accepts what jpeg-js refuses;
 * - when both accept, they give the same width, height and, without chroma subsampling, the same RGBA bytes. With subsampling
 *   the bytes are NOT compared: since 06/10/2026 our chroma step is libjpeg's fancy filter, jpeg-js's is replication, and on
 *   damaged data (chroma changing arbitrarily from sample to sample) the two can differ by the whole range, so no bound
 *   holds. The exact reading is checked on real JPEGs in `decode.test.ts` and `upsampling.test.ts`.
 * Ours may refuse more: it is stricter by design (each case is written in
 * the header of the file that refuses).
 *
 * Seeds: the smallest JPEG of each kind found in the fixtures (baseline
 * 4:2:0, 4:4:4, 4:2:2, grey, APP14, progressive), so that a run costs
 * milliseconds. fast-check only chooses WHERE and HOW to damage them.
 */
import fc from "fast-check";
import { decode as decodeJpegJs } from "jpeg-js";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { decodeJpegRgba } from "./decode";
import { fixtureJpegs, isSubsampled } from "./jpeg-fixtures.test.support";
import { parseJpeg } from "./markers";

vi.setConfig({ testTimeout: fuzzTestTimeout(60_000, process.env) });

const CALLERS = {
  useTArray: true,
  formatAsRGBA: true,
  tolerantDecoding: false,
  maxResolutionInMP: 12,
  maxMemoryUsageInMB: 256,
} as const;
const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };

type Verdict =
  | { readonly ok: false }
  | { readonly ok: true; readonly width: number; readonly height: number; readonly data: Uint8Array };

const ours = (bytes: Uint8Array): Verdict => {
  try {
    return { ok: true, ...decodeJpegRgba(bytes, GUARD) };
  } catch (e) {
    // Only our own refusal is a refusal; anything else is a crash and propagates.
    if (e instanceof ReaderError && e.code === "EXTRACT_FAILED" && e.message.startsWith("jpeg: ")) return { ok: false };
    throw e;
  }
};

const theirs = (bytes: Uint8Array): Verdict => {
  try {
    const image = decodeJpegJs(bytes, CALLERS);
    return { ok: true, width: image.width, height: image.height, data: image.data };
  } catch {
    // Both callers catch every exception of jpeg-js: any throw is a refusal.
    return { ok: false };
  }
};

let seeds: Uint8Array[] = [];

beforeAll(async () => {
  const smallest = new Map<string, Uint8Array>();
  for (const { bytes } of await fixtureJpegs()) {
    const { frame, adobe } = parseJpeg(bytes, { maxPixels: 12e6 });
    const kind = `${frame.process} ${frame.components.map((c) => `${c.h}${c.v}`).join("")} ${adobe ? "adobe" : ""}`;
    const known = smallest.get(kind);
    if (!known || bytes.length < known.length) smallest.set(kind, bytes);
  }
  seeds = [...smallest.values()];
}, 120_000);

/** The offsets of the two-byte length of every marker segment before the first scan's data. */
const lengthOffsets = (b: Uint8Array): number[] => {
  const out: number[] = [];
  let at = 2;
  while (at + 3 < b.length && b[at] === 0xff) {
    const marker = b[at + 1]!;
    out.push(at + 2);
    if (marker === 0xda) break;
    at += 2 + ((b[at + 2]! << 8) | b[at + 3]!);
  }
  return out;
};

/** The offset of the frame header's height (then width) field. */
const frameOffset = (b: Uint8Array): number | undefined => {
  let at = 2;
  while (at + 3 < b.length && b[at] === 0xff) {
    const marker = b[at + 1]!;
    if (marker >= 0xc0 && marker <= 0xc2) return at + 5;
    if (marker === 0xda) return undefined;
    at += 2 + ((b[at + 2]! << 8) | b[at + 3]!);
  }
  return undefined;
};

type Damage =
  | { kind: "truncate"; where: number }
  | { kind: "flip"; where: number; bit: number }
  | { kind: "byte"; where: number; value: number }
  | { kind: "length"; which: number; value: number }
  | { kind: "dimension"; which: number; value: number };

const fraction = fc.double({ min: 0, max: 1, noNaN: true });
const damage: fc.Arbitrary<Damage> = fc.oneof(
  fc.record({ kind: fc.constant("truncate" as const), where: fraction }),
  fc.record({ kind: fc.constant("flip" as const), where: fraction, bit: fc.integer({ min: 0, max: 7 }) }),
  fc.record({ kind: fc.constant("byte" as const), where: fraction, value: fc.integer({ min: 0, max: 255 }) }),
  fc.record({ kind: fc.constant("length" as const), which: fc.nat(), value: fc.integer({ min: 0, max: 0xffff }) }),
  fc.record({
    kind: fc.constant("dimension" as const),
    which: fc.integer({ min: 0, max: 1 }),
    value: fc.integer({ min: 0, max: 0xffff }),
  }),
);

const apply = (seed: Uint8Array, d: Damage): Uint8Array => {
  const b = seed.slice();
  const at = (x: number): number => Math.min(b.length - 1, Math.floor(x * b.length));
  switch (d.kind) {
    case "truncate":
      return b.subarray(0, at(d.where));
    case "flip":
      b[at(d.where)] = b[at(d.where)]! ^ (1 << d.bit);
      return b;
    case "byte":
      b[at(d.where)] = d.value;
      return b;
    case "length": {
      const offsets = lengthOffsets(b);
      const o = offsets[d.which % offsets.length]!;
      b[o] = d.value >> 8;
      b[o + 1] = d.value & 0xff;
      return b;
    }
    case "dimension": {
      const o = frameOffset(b);
      if (o === undefined) return b;
      b[o + 2 * d.which] = d.value >> 8;
      b[o + 2 * d.which + 1] = d.value & 0xff;
      return b;
    }
  }
};

describe("decodeJpegRgba against jpeg-js on damaged JPEGs (differential fuzz)", () => {
  it("has seeds of every kind of the fixtures, each accepted by both", () => {
    expect(seeds.length).toBe(7); // the 7th kind, progressive 4:4:4 Adobe, from fixtures/docx-reel/11 (06/10)
    for (const seed of seeds) {
      const a = ours(seed);
      const b = theirs(seed);
      expect(a.ok && b.ok).toBe(true);
    }
  });

  it("never crashes, never accepts what jpeg-js refuses, and agrees pixel for pixel when both accept and nothing is subsampled", () => {
    fc.assert(
      fc.property(fc.nat(), damage, (pick, d) => {
        const bytes = apply(seeds[pick % seeds.length]!, d);
        const a = ours(bytes);
        const b = theirs(bytes);
        if (a.ok) {
          expect(b.ok).toBe(true);
          if (!b.ok) return;
          expect([a.width, a.height]).toEqual([b.width, b.height]);
          // The frame is re-read from the damaged bytes: it is the one our decoder accepted.
          if (!isSubsampled(parseJpeg(bytes, { maxPixels: 12e6 }).frame))
            expect(Buffer.from(a.data).equals(Buffer.from(b.data))).toBe(true);
        }
      }),
      { numRuns: fuzzRuns(300, process.env), seed: 20261005 },
    );
  });
});
