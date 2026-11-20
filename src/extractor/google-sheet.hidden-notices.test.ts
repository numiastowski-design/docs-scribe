import { describe, expect, it } from "vitest";

import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";

// Decision of 08/10/2026: filtered rows (hiddenByFilter) that carry at least one non-empty value and are
// dropped are SAID in { kind: "hidden-dropped", what: "filtered-rows", count } (never their values). A row hidden by
// the user only is expected and says nothing. Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/google-sheet.hidden-notices.test.ts

type RowFlags = { hiddenByUser?: boolean; hiddenByFilter?: boolean };

const noticesOf = async (rowMetadata: RowFlags[], rows: string[][]): Promise<unknown[]> => {
  const meta: GoogleSheetMeta = { sheets: [{ properties: { title: "S", index: 0 }, rowMetadata }] };
  const { notices } = await extractGoogleSheet(meta, new Map([["S", rows]]));
  return notices.filter((n) => n.kind === "hidden-dropped");
};

describe("N3 — filtered rows are counted in a hidden-dropped/filtered-rows notice", () => {
  it("🔴 one filtered row with values: { hidden-dropped, filtered-rows, 1 }", async () => {
    // Failure value (before): no notice.
    expect(await noticesOf([{}, { hiddenByFilter: true }, {}], [["h"], ["secretvalue"], ["k"]])).toEqual([
      { kind: "hidden-dropped", what: "filtered-rows", count: 1 },
    ]);
  });

  it("🔴 three filtered rows (the last one too) give ONE entry with count 3, and no value in it", async () => {
    const found = await noticesOf(
      [{}, { hiddenByFilter: true }, { hiddenByFilter: true }, {}, { hiddenByFilter: true }],
      [["a"], ["secret1"], ["secret2"], ["b"], ["secret3"]],
    );
    expect(found).toEqual([{ kind: "hidden-dropped", what: "filtered-rows", count: 3 }]);
    expect(JSON.stringify(found)).not.toContain("secret");
  });

  it("🔴 a filtered row with one non-empty cell among empty ones counts", async () => {
    expect(
      await noticesOf(
        [{}, { hiddenByFilter: true }],
        [
          ["a", "b"],
          ["", "x"],
        ],
      ),
    ).toEqual([{ kind: "hidden-dropped", what: "filtered-rows", count: 1 }]);
  });

  it("🔴 a row both hidden by the user and filtered is counted once", async () => {
    expect(await noticesOf([{}, { hiddenByUser: true, hiddenByFilter: true }, {}], [["a"], ["both"], ["c"]])).toEqual([
      { kind: "hidden-dropped", what: "filtered-rows", count: 1 },
    ]);
  });

  it("🔴 an empty filtered row does not count: 1 with values + 1 empty gives count 1", async () => {
    expect(await noticesOf([{}, { hiddenByFilter: true }, { hiddenByFilter: true }], [["a"], ["", ""], ["v"]])).toEqual(
      [{ kind: "hidden-dropped", what: "filtered-rows", count: 1 }],
    );
  });

  it("non-regression: only an empty filtered row gives no notice", async () => {
    expect(await noticesOf([{}, { hiddenByFilter: true }], [["a"], ["", ""]])).toEqual([]);
  });

  it("non-regression: a row hidden by the user only gives NO notice (expected, not reported)", async () => {
    expect(await noticesOf([{}, { hiddenByUser: true }, {}], [["a"], ["usersecret"], ["c"]])).toEqual([]);
  });

  it("non-regression: no filtered row (false or absent) gives no notice", async () => {
    expect(
      await noticesOf(
        [{ hiddenByFilter: false }, {}, { hiddenByFilter: false, hiddenByUser: false }],
        [["a"], ["b"], ["c"]],
      ),
    ).toEqual([]);
  });
});
