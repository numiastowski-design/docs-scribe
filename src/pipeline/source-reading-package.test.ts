/**
 * `structureOnly`: the isolation of a read only the structure control makes. A refusal of the PART (damaged, over a ceiling: an `ReaderError`) leaves the
 * structure unreadable; a defect of our own code is not swallowed.
 * Run: npx vitest run --project=main src/pipeline/source-reading-package.test.ts
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { PartDamagedError } from "../extractor/ooxml/package";
import { structureOnly } from "./source-reading-package";

describe("structureOnly", () => {
  it("gives the value of a read that succeeds", async () => {
    // FAILS WITH: unreadable true.
    expect(await structureOnly(async () => 7)).toStrictEqual({ value: 7 });
  });

  it("a damaged part and a ceiling refusal leave the structure unreadable", async () => {
    // FAILS WITH: the error thrown through (the whole reading refused).
    expect(
      await structureOnly(async () =>
        Promise.reject(new PartDamagedError(new ReaderError("EXTRACT_FAILED", "damaged"))),
      ),
    ).toStrictEqual({ unreadable: true });
    expect(
      await structureOnly(async () => Promise.reject(new ReaderError("EXTRACT_FAILED", "over the ceiling"))),
    ).toStrictEqual({ unreadable: true });
  });

  it("a programming error is NOT swallowed", async () => {
    // FAILS WITH: { unreadable: true } (a defect of ours read as a damaged file).
    await expect(structureOnly(async () => Promise.reject(new TypeError("our own bug")))).rejects.toThrow(TypeError);
  });
});
