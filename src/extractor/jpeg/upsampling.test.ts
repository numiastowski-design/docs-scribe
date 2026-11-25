import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { decodeJpegRgba } from "./decode";
import { fixtureJpegs } from "./jpeg-fixtures.test.support";

/**
 * Chroma upsampling against a second, independent reading (06/10/2026): Pillow 12.3.0, which decodes with libjpeg-turbo,
 * "fancy" (triangle) upsampling on, as `fixtures/jpeg/README.md` records. libjpeg's inverse DCT is not ours, so the
 * samples differ by a level or two before any upsampling; what the references show is the chroma step.
 * Measured 06/10/2026 on the three references (worst level, mean level, every channel of every compared pixel):
 *   replication (jpeg-js's, the decoder before this lot): 4:2:0 baseline 108 / 5.761, 4:2:0 progressive 7 / 0.389, 4:2:2 6 / 0.602;
 *   fancy (libjpeg-turbo's h2v1 and h2v2, now): 3 / 0.476, 3 / 0.290, 3 / 0.561.
 * The remaining distance is not the chroma step: our inverse DCT and our truncating YCbCr to RGB (libjpeg rounds) differ from
 * libjpeg's. Bounds: worst 3 levels, mean 0.7 (valeur de départ à calibrer: the three measured means are 0.29 to 0.56, the
 * replication's worst 108 on the 4:2:0 baseline is far outside).
 */
const GUARD = { maxResolutionInMP: 12, maxMemoryUsageInMB: 256 };
const DIR = join(dirname(fileURLToPath(import.meta.url)), "../../../fixtures/jpeg");

interface Reference {
  readonly name: string;
  /** sha256 of the source JPEG (found among the fixtures: the document paths are not this test's business). */
  readonly sha: string;
  readonly sampling: string;
  readonly width: number;
  readonly height: number;
  /** Rows kept at the top and at the bottom (all rows when undefined). */
  readonly edge?: number;
}

const REFERENCES: readonly Reference[] = [
  {
    name: "pillow-422-baseline-2144x1424",
    sha: "6f37dcc460b72fd0b3e82db008b1e812326b273b6718e5a868af8f44472480f1",
    sampling: "4:2:2 baseline",
    width: 2144,
    height: 1424,
    edge: 16,
  },
];

/** Worst and mean absolute difference of the RGB channels, over the rows the reference holds. */
const compare = async (ref: Reference): Promise<{ worst: number; mean: number; pixels: number }> => {
  const fixture = (await fixtureJpegs()).find((f) => createHash("sha256").update(f.bytes).digest("hex") === ref.sha);
  if (!fixture) throw new Error(`fixture not found: ${ref.name}`);
  const pillow = new Uint8Array(readFileSync(join(DIR, `${ref.name}.bin`)));
  const ours = decodeJpegRgba(fixture.bytes, GUARD);
  expect([ours.width, ours.height]).toEqual([ref.width, ref.height]);
  const rows =
    ref.edge === undefined
      ? Array.from({ length: ref.height }, (_, y) => y)
      : [...Array(ref.edge).keys(), ...Array.from({ length: ref.edge }, (_, i) => ref.height - ref.edge! + i)];
  expect(pillow.length).toBe(rows.length * ref.width * 3);
  let worst = 0;
  let sum = 0;
  rows.forEach((y, k) => {
    for (let x = 0; x < ref.width; x++) {
      for (let c = 0; c < 3; c++) {
        const d = Math.abs(ours.data[(y * ref.width + x) * 4 + c]! - pillow[(k * ref.width + x) * 3 + c]!);
        worst = Math.max(worst, d);
        sum += d;
      }
    }
  });
  return { worst, mean: sum / (rows.length * ref.width * 3), pixels: rows.length * ref.width };
};

describe("chroma upsampling against Pillow 12.3.0 (libjpeg-turbo, fancy upsampling)", () => {
  for (const ref of REFERENCES) {
    it(`${ref.sampling}, ${ref.width}x${ref.height}: within the stated distance of Pillow's reading`, async () => {
      const { worst, mean } = await compare(ref);
      expect(worst, `worst ${worst}, mean ${mean.toFixed(3)}`).toBeLessThanOrEqual(3);
      expect(mean, `worst ${worst}, mean ${mean.toFixed(3)}`).toBeLessThanOrEqual(0.7);
    }, 60_000);
  }
});

// Review of 06/10/2026: none of the three references above has a height that is a multiple of 16, the case where the
// bottom context row of a 4:2:0 chroma plane must be clamped to the component's last real row (else it reads the
// block padding: 255 levels off, measured on a mutant). A 64x48 picture written by our own encoder (4:2:0, gradients
// and a sharp red band on the last rows), and Pillow 12.3.0's reading of it (fixtures/jpeg/README.md).
describe("chroma upsampling at a height that is a multiple of 16", () => {
  it("64x48, 4:2:0: within 3 levels of Pillow's reading, last rows included", () => {
    const ours = decodeJpegRgba(new Uint8Array(readFileSync(join(DIR, "ours-420-64x48.jpg"))), GUARD);
    const pillow = new Uint8Array(readFileSync(join(DIR, "ours-420-64x48.pillow-rgb.bin")));
    expect([ours.width, ours.height]).toEqual([64, 48]);
    let worst = 0;
    for (let i = 0; i < 64 * 48; i++) {
      for (let c = 0; c < 3; c++) worst = Math.max(worst, Math.abs(ours.data[i * 4 + c]! - pillow[i * 3 + c]!));
    }
    expect(worst).toBeLessThanOrEqual(3);
  });
});
