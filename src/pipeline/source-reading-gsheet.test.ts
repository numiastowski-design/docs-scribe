/** Behaviour tests for `readGoogleSheetConservationSource`
 *  (MIKSUDOC « Spécification technique » T11): a malformed API response becomes
 *  a named failure — never a thrown error, never a reason not to publish
 *  (21/09/2026) — and the declared removals (tab cap, size-cap
 *  withhold, hidden sheet, OBJECT tab) never surface as missing words.
 *  That a real API response yields a usable declared-format source is held by the publisher's application, on a
 *  real sheet this package does not carry (298 rows × 6 columns, withheld, its one unit excluded and empty). */
import { describe, expect, it } from "vitest";

import {
  MAX_SHEET_COLS,
  MAX_SHEET_ROWS,
  MAX_WORKBOOK_TABS,
  PAGE_SIZED_LIMITS,
} from "../extractor/page-caps.test.support";
import type { GoogleSheetMeta } from "../extractor/google-sheet";
import { isDeclaredFormatSource, isReadFailure } from "./conservation-source";
import { readGoogleSheetConservationSource } from "./source-reading-gsheet";

const sheetOf = (
  title: string,
  index: number,
  extra: Partial<GoogleSheetMeta["sheets"][number]["properties"]> = {},
): GoogleSheetMeta["sheets"][number] => ({
  properties: { title, index, ...extra },
});

const allWords = (source: ReturnType<typeof readGoogleSheetConservationSource>): string[] => {
  if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
  return source.units.flatMap((u) => u.tokens.map((t) => t.text));
};

describe("readGoogleSheetConservationSource — malformed responses", () => {
  it("🔴 GUARD BITES: a response with no sheets array — a named failure, never a thrown error", () => {
    const source = readGoogleSheetConservationSource({} as GoogleSheetMeta, new Map(), PAGE_SIZED_LIMITS);
    expect(isReadFailure(source)).toBe(true);
  });
});

describe("readGoogleSheetConservationSource — the declared removals never surface as missing", () => {
  it("a sheet past the tab cap: its unit is excluded, never a source of tokens", () => {
    const meta: GoogleSheetMeta = {
      sheets: Array.from({ length: MAX_WORKBOOK_TABS + 2 }, (_, i) => sheetOf(`Sheet${i + 1}`, i)),
    };
    const values = new Map(meta.sheets.map((s, i) => [s.properties.title, [[`word${i + 1}unique`]]]));
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(MAX_WORKBOOK_TABS + 2);
    expect([...source.excludedUnits].sort((a, b) => a - b)).toEqual([MAX_WORKBOOK_TABS, MAX_WORKBOOK_TABS + 1]);
    expect(allWords(source)).toContain("word1unique");
    expect(allWords(source)).not.toContain(`word${MAX_WORKBOOK_TABS + 1}unique`);
  });

  it("a sheet over MAX_SHEET_ROWS: withheld whole, its unit excluded and empty", () => {
    const bigRows = Array.from({ length: MAX_SHEET_ROWS + 2 }, (_, i) => [`r${i}unique`]);
    const meta: GoogleSheetMeta = { sheets: [sheetOf("Big", 0), sheetOf("Small", 1)] };
    const values = new Map([
      ["Big", bigRows],
      ["Small", [["next sheet unique"]]],
    ]);
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).not.toContain("r0unique");
    expect(allWords(source)).toContain("unique");
  });

  it("a sheet over MAX_SHEET_COLS: withheld whole, its unit excluded and empty", () => {
    const wideRow = Array.from({ length: MAX_SHEET_COLS + 2 }, (_, i) => `c${i}unique`);
    const meta: GoogleSheetMeta = { sheets: [sheetOf("Wide", 0), sheetOf("Small", 1)] };
    const values = new Map([
      ["Wide", [wideRow]],
      ["Small", [["next sheet unique"]]],
    ]);
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).not.toContain("c0unique");
    expect(allWords(source)).toContain("unique");
  });

  it("a hidden sheet: a unit of its own, excluded, never counted toward the cap", () => {
    const meta: GoogleSheetMeta = {
      sheets: [sheetOf("Shown1", 0), sheetOf("Hidden", 1, { hidden: true }), sheetOf("Shown2", 2)],
    };
    const values = new Map([
      ["Shown1", [["shown one unique"]]],
      ["Hidden", [["hidden words unique"]]],
      ["Shown2", [["shown two unique"]]],
    ]);
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect(source.units.length).toBe(3);
    expect([...source.excludedUnits]).toEqual([1]);
    expect(allWords(source)).not.toContain("hidden");
    expect(allWords(source)).toContain("one");
    expect(allWords(source)).toContain("two");
  });

  it("an OBJECT tab: excluded, never a source of tokens (no grid was ever fetched for it)", () => {
    const meta: GoogleSheetMeta = { sheets: [sheetOf("Chart", 0, { sheetType: "OBJECT" }), sheetOf("Data", 1)] };
    const values = new Map([["Data", [["kept word unique"]]]]);
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    expect([...source.excludedUnits]).toEqual([0]);
    expect(allWords(source)).toContain("kept");
  });

  it("a hidden row: its word never read, the rest of the sheet still is", () => {
    const meta: GoogleSheetMeta = {
      sheets: [{ properties: { title: "S", index: 0 }, rowMetadata: [{ hiddenByUser: true }, {}] }],
    };
    const values = new Map([["S", [["masked"], ["kept"]]]]);
    const source = readGoogleSheetConservationSource(meta, values, PAGE_SIZED_LIMITS);
    if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
    const words = allWords(source);
    expect(words).toContain("kept");
    expect(words).not.toContain("masked");
  });
});
