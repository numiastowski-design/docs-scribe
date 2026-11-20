import { describe, expect, it } from "vitest";

import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";

// Decision of 08/10/2026: a row hidden by a filter (rowMetadata[r].hiddenByFilter, Sheets API
// DimensionProperties) is not published, like a row hidden by the user. Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/google-sheet.filtered-rows.test.ts

type RowFlags = { hiddenByUser?: boolean; hiddenByFilter?: boolean };

const metaOf = (rowMetadata: RowFlags[]): GoogleSheetMeta => ({
  sheets: [{ properties: { title: "S", index: 0 }, rowMetadata }],
});

const published = async (rowMetadata: RowFlags[], rows: string[][]): Promise<string[]> => {
  const { doc } = await extractGoogleSheet(metaOf(rowMetadata), new Map([["S", rows]]));
  const tabs = doc.blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  const table = tabs.tabs[0]!.blocks[0];
  if (table?.kind !== "table") throw new Error("expected table");
  return table.rows.flatMap((r) =>
    r.cells.flatMap((c) => c.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text) : []))),
  );
};

describe("extractGoogleSheet — rows hidden by a filter", () => {
  it("🔴 a filtered row among visible rows is not published", async () => {
    const texts = await published(
      [{}, { hiddenByFilter: true }, {}],
      [
        ["head1", "head2"],
        ["filteredsecret1", "filteredsecret2"],
        ["keep1", "keep2"],
      ],
    );
    // Failure value (before): the filtered values are published.
    expect(texts).not.toContain("filteredsecret1");
    expect(texts).not.toContain("filteredsecret2");
    expect(texts).toEqual(["head1", "head2", "keep1", "keep2"]);
  });

  it("🔴 several filtered rows, including the last one, are all dropped", async () => {
    const texts = await published(
      [{}, { hiddenByFilter: true }, { hiddenByFilter: true }, {}, { hiddenByFilter: true }],
      [["a"], ["f1"], ["f2"], ["b"], ["f3"]],
    );
    expect(texts).toEqual(["a", "b"]);
  });

  it("🔴 a row both hidden by the user and by a filter is dropped once: its neighbours are untouched", async () => {
    const texts = await published([{}, { hiddenByUser: true, hiddenByFilter: true }, {}], [["a"], ["both"], ["c"]]);
    expect(texts).toEqual(["a", "c"]);
  });

  it("non-regression: hiddenByFilter false or absent publishes the row", async () => {
    const texts = await published(
      [{ hiddenByFilter: false }, {}, { hiddenByFilter: false, hiddenByUser: false }],
      [["a"], ["b"], ["c"]],
    );
    expect(texts).toEqual(["a", "b", "c"]);
  });

  it("non-regression: a row hidden by the user only is still dropped", async () => {
    const texts = await published([{}, { hiddenByUser: true }, {}], [["a"], ["b"], ["c"]]);
    expect(texts).toEqual(["a", "c"]);
  });
});
