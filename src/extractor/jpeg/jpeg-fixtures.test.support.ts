/** Test support: every JPEG embedded in the public Office documents of `fixtures/` (`publicFixtures`),
 *  read with our own zip reader, deduplicated by content. A document the zip
 *  reader refuses (the hostile fixtures) is skipped: it carries no JPEG we
 *  could compare.
 *  NOT here: any decoding or comparison (the `jpeg/*.test.ts` files).
 *  Called by: the tests of `src/extractor/jpeg/`. */

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { byCodeUnit } from "doc-ir/internal";
import { publicFixturePath, publicFixtures } from "../public-fixtures.test.support";
import { openZip } from "../zip/zip-package";
import type { JpegFrame } from "./markers";

export interface FixtureJpeg {
  /** `<document path under fixtures/>!<entry name>`: a technical id, never content. */
  readonly id: string;
  readonly bytes: Uint8Array;
}

const LIMITS = { partMaxBytes: 64 * 1024 * 1024, packageMaxBytes: 512 * 1024 * 1024 };

/** Measurement corpora versed on 14/10/2026 for the heading rules (25 Word, 36 workbooks): the encoder totals below were calibrated
 *  on 06/10 without them, and their pictures pushed the test past its time limit in CI. They measure headings, not the encoder. */
/** Same for the corpus cases versed on 19/10/2026 (`docx-corpus-cases`, `pdf-corpus-cases`) and the real Word files of the layout-table lot
 *  (`docx-layout-tables`, 24/10/2026): they measure words kept and structure, not the encoder. */
const MEASUREMENT_ONLY = [
  "fixtures/docx-heading-reel/",
  "fixtures/xlsx-heading-reel/",
  "fixtures/docx-corpus-cases/",
  "fixtures/pdf-corpus-cases/",
  "fixtures/docx-layout-tables/",
];

/** The public Office documents, measurement corpora left out, as paths from the root of the repository. */
const officeDocuments = (): string[] =>
  publicFixtures(".docx", ".pptx", ".xlsx").filter(
    (path) => !MEASUREMENT_ONLY.some((folder) => path.startsWith(folder)),
  );

let cache: Promise<FixtureJpeg[]> | undefined;

const load = async (): Promise<FixtureJpeg[]> => {
  const seen = new Set<string>();
  const out: FixtureJpeg[] = [];
  for (const path of officeDocuments().sort(byCodeUnit)) {
    let zip;
    try {
      zip = openZip(new Uint8Array(readFileSync(publicFixturePath(path))), LIMITS);
    } catch {
      continue;
    }
    for (const name of zip.names) {
      let bytes: Uint8Array | undefined;
      try {
        bytes = await zip.readBytes(name);
      } catch {
        continue;
      }
      if (!bytes || bytes.length < 3 || bytes[0] !== 0xff || bytes[1] !== 0xd8 || bytes[2] !== 0xff) continue;
      const sha = createHash("sha256").update(bytes).digest("hex");
      if (seen.has(sha)) continue;
      seen.add(sha);
      out.push({ id: `${path.slice("fixtures/".length)}!${name}`, bytes });
    }
  }
  return out;
};

/** The distinct JPEGs of the fixtures, read once per test file. */
export const fixtureJpegs = (): Promise<FixtureJpeg[]> => (cache ??= load());

/** True when a component has fewer samples than the image has pixels (chroma subsampling): the pixels our decoder fills
 *  between samples by libjpeg's fancy filter (2:1 ratios) differ from jpeg-js's replication. */
export const isSubsampled = (frame: JpegFrame): boolean =>
  frame.components.some((c) => c.h !== frame.maxH || c.v !== frame.maxV);

/** Worst and mean absolute difference between two RGBA buffers of one size, over the three colour channels. */
export const rgbaDistance = (a: Uint8Array, b: Uint8Array): { worst: number; mean: number } => {
  let worst = 0;
  let sum = 0;
  for (let i = 0; i < a.length; i += 4) {
    for (let c = 0; c < 3; c++) {
      const d = Math.abs(a[i + c]! - b[i + c]!);
      if (d > worst) worst = d;
      sum += d;
    }
  }
  return { worst, mean: sum / ((a.length / 4) * 3) };
};
