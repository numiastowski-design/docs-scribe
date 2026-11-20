import { describe, expect, it } from "vitest";

import { MAX_SHEET_ROWS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";

// 08/10/2026: ONE row cap for worksheet tabs (MAX_SHEET_ROWS), 55 instead of 100, Google Sheets included.
// Run: npx vitest run --project=main src/extractor/google-sheet.rows-55.test.ts
const META: GoogleSheetMeta = {
  sheets: [{ properties: { title: "Long", index: 0 } }, { properties: { title: "Small", index: 1 } }],
};
const rowsOf = (n: number, prefix: string): string[][] => Array.from({ length: n }, (_, i) => [`${prefix}${i + 1}x`]);
const run = async (
  longRows: string[][],
): Promise<{ notices: readonly { kind: string }[]; titles: string[]; json: string }> => {
  const { doc, notices } = await extractGoogleSheet(
    META,
    new Map([
      ["Long", longRows],
      ["Small", rowsOf(2, "s")],
    ]),
    PAGE_SIZED_LIMITS,
  );
  const tabs = doc.blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected tabs");
  return { notices, titles: tabs.tabs.map((t) => t.title), json: JSON.stringify(doc) };
};

describe("Google Sheet — a tab of more than MAX_SHEET_ROWS kept rows is withheld (sheet-withheld), a tab AT the cap is published", () => {
  it("a tab of MAX_SHEET_ROWS filled rows is published whole, no remark", async () => {
    const r = await run(rowsOf(MAX_SHEET_ROWS, "l"));
    expect(r.titles).toEqual(["Long", "Small"]);
    expect(r.json).toContain(`l${MAX_SHEET_ROWS}x`);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });

  it("a tab of MAX_SHEET_ROWS + 1 filled rows carries no tab, with the existing remark (rows 56); the next tab stays", async () => {
    const r = await run(rowsOf(MAX_SHEET_ROWS + 1, "l"));
    // Failure value (before, cap 100): the tab is published.
    expect(r.titles).toEqual(["Small"]);
    expect(r.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Long", rows: MAX_SHEET_ROWS + 1, cols: 1 });
  });

  it("non-regression: entirely empty rows do not count (MAX_SHEET_ROWS filled + empty rows is published)", async () => {
    const empties = [...Array.from({ length: 40 }, () => [] as string[]), ...Array.from({ length: 5 }, () => [""])];
    const r = await run([...rowsOf(MAX_SHEET_ROWS, "g"), ...empties]);
    expect(r.titles).toEqual(["Long", "Small"]);
    expect(r.json).toContain(`g${MAX_SHEET_ROWS}x`);
    expect(r.notices.filter((n) => n.kind === "sheet-withheld")).toEqual([]);
  });
});

describe("the sheet row cap constant", () => {
  it("MAX_SHEET_ROWS is 55 (the fixtures above are built from it; at the old 100 they would pass vacuously)", () => {
    expect(MAX_SHEET_ROWS).toBe(55);
  });
});
