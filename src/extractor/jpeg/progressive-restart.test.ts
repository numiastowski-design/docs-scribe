import { decode as decodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { decodeJpegRgba } from "./decode";
import { BLOCK_AC1, BLOCK_DC, HEIGHT, progressiveWithRestarts, WIDTH } from "./progressive-restart.test.support";

const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };

/** Mean grey level of a block, and its left-column minus right-column level (sign of coefficient 1). */
const blockStats = (rgba: Uint8Array, block: number): { mean: number; lean: number } => {
  const bx = (block % 4) * 8;
  const by = Math.floor(block / 4) * 8;
  let sum = 0;
  let left = 0;
  let right = 0;
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      const v = rgba[((by + y) * WIDTH + bx + x) * 4]!;
      sum += v;
      if (x === 0) left += v;
      if (x === 7) right += v;
    }
  }
  return { mean: sum / 64, lean: (left - right) / 8 };
};

describe("progressive JPEG with a restart interval (DRI), generated: DC first + AC first", () => {
  for (const overlongRun of [false, true]) {
    it(`decodes to the pixels fixed at generation (${overlongRun ? "end-of-band run reaching past a restart: dropped" : "runs inside their interval"})`, () => {
      const bytes = progressiveWithRestarts({ overlongRun });
      const out = decodeJpegRgba(bytes, GUARD);
      expect([out.width, out.height]).toEqual([WIDTH, HEIGHT]);
      for (let b = 0; b < BLOCK_DC.length; b++) {
        const { mean, lean } = blockStats(out.data, b);
        expect(mean, `block ${b} mean`).toBeCloseTo(128 + 8 * BLOCK_DC[b]!, 0);
        // Coefficient 1 = 64 on the left/right columns: about +-22 level difference when set, 0 otherwise.
        if (BLOCK_AC1[b] === 0) expect(Math.abs(lean), `block ${b} lean`).toBeLessThan(1);
        else expect(Math.sign(lean) * BLOCK_AC1[b]!, `block ${b} lean`).toBe(1);
        if (BLOCK_AC1[b] !== 0) expect(Math.abs(lean), `block ${b} strength`).toBeGreaterThan(15);
      }
    });
  }

  for (const overlongRun of [false, true]) {
    it(`agrees with jpeg-js on the same stream (${overlongRun ? "end-of-band run reaching past a restart" : "runs inside their interval"})`, () => {
      const bytes = progressiveWithRestarts({ overlongRun });
      const ours = decodeJpegRgba(bytes, GUARD);
      const ref = decodeJpegJs(bytes, {
        useTArray: true,
        formatAsRGBA: true,
        tolerantDecoding: false,
        maxResolutionInMP: 12,
        maxMemoryUsageInMB: 256,
      });
      let worst = 0;
      for (let i = 0; i < ours.data.length; i++) worst = Math.max(worst, Math.abs(ours.data[i]! - ref.data[i]!));
      expect(worst).toBeLessThanOrEqual(2);
    });
  }
});
