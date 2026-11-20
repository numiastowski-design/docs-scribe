import { describe, expect, it } from "vitest";

import { MAX_TABLE_COLS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";

// 08/10/2026: ONE cap of 12 columns for worksheet tabs (MAX_TABLE_COLS was 10), Google Sheets included.
// Run: npx vitest run --project=main src/extractor/google-sheet.cols-12.test.ts
const META: GoogleSheetMeta = {
  sheets: [{ properties: { title: "Wide", index: 0 } }, { properties: { title: "Small", index: 1 } }],
};
const rowOf = (n: number, prefix: string): string[] => Array.from({ length: n }, (_, i) => `${prefix}${i + 1}`);
const run = async (
  wideRows: string[][],
): Promise<{ notices: readonly { kind: string }[]; titles: string[]; wideCols: number | undefined }> => {
  const { doc, notices } = await extractGoogleSheet(
    META,
    new Map([
      ["Wide", wideRows],
      ["Small", [rowOf(2, "s")]],
    ]),
    PAGE_SIZED_LIMITS,
  );
  const tabs = doc.blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  const table = tabs.tabs[0]?.blocks[0];
  return {
    notices,
    titles: tabs.tabs.map((t) => t.title),
    wideCols: table?.kind === "table" ? table.rows[0]?.cells.length : undefined,
  };
};

describe("Google Sheet — a tab of more than the column cap is withheld (sheet-withheld), a tab AT the cap is published", () => {
  it("a tab of MAX_TABLE_COLS filled columns is published (it was withheld from 11 on with the cap of 10)", async () => {
    const r = await run([rowOf(MAX_TABLE_COLS, "w")]);
    expect(r.titles).toEqual(["Wide", "Small"]);
    expect(r.wideCols).toBe(MAX_TABLE_COLS);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });

  it("a tab of MAX_TABLE_COLS + 1 filled columns carries no tab, with the existing remark; the next tab stays", async () => {
    const r = await run([rowOf(MAX_TABLE_COLS + 1, "w")]);
    expect(r.titles).toEqual(["Small"]);
    expect(r.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Wide", rows: 1, cols: MAX_TABLE_COLS + 1 });
  });

  it("entirely empty columns do not count: MAX_TABLE_COLS + 1 columns of which one is empty is published", async () => {
    const r = rowOf(MAX_TABLE_COLS + 1, "e").map((t, i) => (i === 4 ? "" : t));
    const out = await run([r, r.map((t) => (t === "" ? "" : `${t}b`))]);
    expect(out.titles).toEqual(["Wide", "Small"]);
    expect(out.wideCols).toBe(MAX_TABLE_COLS);
    expect(out.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });
});
