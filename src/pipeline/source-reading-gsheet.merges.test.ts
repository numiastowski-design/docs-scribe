/** Characterisation of the Google Sheet control reading's merge rule (ours, written twice with `sheet-grid.ts#tableOfKept`): the
 *  first merge wins, a merge anchored on an anchor or a covered cell is ignored whole, a merge needs a kept first cell and at
 *  least two kept lines. Ranges are half-open and 0-based on this side.
 *  Run: npx vitest run --project=main src/pipeline/source-reading-gsheet.merges.test.ts */
import { describe, expect, it } from "vitest";

import type { GoogleSheetMeta } from "../extractor/google-sheet";
import { isDeclaredFormatSource } from "./conservation-source";
import { readGoogleSheetConservationSource } from "./source-reading-gsheet";

type Merge = { startRowIndex: number; endRowIndex: number; startColumnIndex: number; endColumnIndex: number };

const expectedWords = (merges: Merge[], rows: string[][]): string[] => {
  const meta: GoogleSheetMeta = { sheets: [{ properties: { title: "S", index: 0 }, merges }] };
  const source = readGoogleSheetConservationSource(meta, new Map([["S", rows]]));
  if (!isDeclaredFormatSource(source)) throw new Error("read failed");
  return source.units[0]!.tokens.map((t) => t.text);
};
const merge = (r1: number, r2: number, c1: number, c2: number): Merge => ({
  startRowIndex: r1,
  endRowIndex: r2,
  startColumnIndex: c1,
  endColumnIndex: c2,
});

describe("Google Sheet control reading: which merges cover cells", () => {
  it("the first merge wins; a later one on the same anchor is ignored whole", () => {
    // Failure value (second merge wins): ["aone", "btwo"] — b shown, the second row's a covered.
    expect(
      expectedWords(
        [merge(0, 1, 0, 2), merge(0, 2, 0, 1)],
        [
          ["aone", "btwo"],
          ["athree", "bfour"],
        ],
      ),
    ).toEqual(["aone", "athree", "bfour"]);
  });

  it("a merge anchored on a covered cell is ignored whole", () => {
    // Failure value (honoured): the second row's middle cell covered too.
    expect(
      expectedWords(
        [merge(0, 1, 0, 3), merge(0, 2, 1, 2)],
        [
          ["aone", "btwo", "cthree"],
          ["dfour", "efive", "fsix"],
        ],
      ),
    ).toEqual(["aone", "dfour", "efive", "fsix"]);
  });

  it("a merge over a single cell covers nothing and does not take its anchor", () => {
    expect(expectedWords([merge(0, 1, 0, 1)], [["aone", "btwo"]])).toEqual(["aone", "btwo"]);
    // Failure value (the single-cell merge claims the anchor): ["aone", "btwo"] — the next merge on that anchor is then ignored.
    expect(expectedWords([merge(0, 1, 0, 1), merge(0, 1, 0, 2)], [["aone", "btwo"]])).toEqual(["aone"]);
  });

  it("a merge spanning two rows of which only one is kept still covers across kept columns", () => {
    expect(
      expectedWords(
        [merge(0, 2, 0, 2)],
        [
          ["aone", "btwo"],
          ["", ""],
        ],
      ),
    ).toEqual(["aone"]);
  });

  it("a merge spanning two rows with one kept row and one kept column covers nothing", () => {
    expect(
      expectedWords(
        [merge(0, 2, 0, 1)],
        [
          ["aone", "btwo"],
          ["", ""],
        ],
      ),
    ).toEqual(["aone", "btwo"]);
  });

  it("a merge whose first cell is not kept is ignored", () => {
    expect(
      expectedWords(
        [merge(0, 2, 0, 2)],
        [
          ["", "btwo"],
          ["", "dfour"],
        ],
      ),
    ).toEqual(["btwo", "dfour"]);
  });
});
