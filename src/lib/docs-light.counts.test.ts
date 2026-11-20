/**
 * The arithmetic on structure counts, through the published light door (`docs-scribe/light`): `emptyCounts`,
 * `addCounts`, `subtractCounts`, `isZeroCounts`. A client that compares the structure of a source with the structure
 * of its page adds what it counts and subtracts one side from the other; every expected number below is written by
 * hand.
 */
import { describe, expect, it } from "vitest";

import { addCounts, emptyCounts, isZeroCounts, subtractCounts, type StructureCounts } from "./docs-light";

const SOURCE: StructureCounts = {
  headings: 5,
  headingsByLevel: [1, 2, 2, 0, 0, 0],
  tables: 2,
  cells: 12,
  items: 7,
  rows: 4,
};
const PAGE: StructureCounts = {
  headings: 3,
  headingsByLevel: [1, 1, 0, 1, 0, 0],
  tables: 3,
  cells: 10,
  items: 7,
  rows: 9,
};

describe("emptyCounts", () => {
  it("is all zeros, six heading levels", () => {
    expect(emptyCounts()).toEqual({
      headings: 0,
      headingsByLevel: [0, 0, 0, 0, 0, 0],
      tables: 0,
      cells: 0,
      items: 0,
      rows: 0,
    });
  });

  it("is a FRESH object each time: a caller that increments one changes no other", () => {
    const first = emptyCounts();
    first.headings += 1;
    first.headingsByLevel[0] += 1;
    // Failure value: { headings: 1, headingsByLevel: [1, …] }, one shared object handed out twice.
    expect(emptyCounts()).toEqual({
      headings: 0,
      headingsByLevel: [0, 0, 0, 0, 0, 0],
      tables: 0,
      cells: 0,
      items: 0,
      rows: 0,
    });
  });
});

describe("addCounts", () => {
  it("adds field by field and level by level", () => {
    expect(addCounts(SOURCE, PAGE)).toEqual({
      headings: 8,
      headingsByLevel: [2, 3, 2, 1, 0, 0],
      tables: 5,
      cells: 22,
      items: 14,
      rows: 13,
    });
  });

  it("changes neither of its arguments", () => {
    const before = JSON.stringify([SOURCE, PAGE]);
    addCounts(SOURCE, PAGE);
    expect(JSON.stringify([SOURCE, PAGE])).toBe(before);
  });
});

describe("subtractCounts", () => {
  it("is the SIGNED difference, first minus second, field by field and level by level", () => {
    // Failure values: the absolute difference (rows 5, tables 1), or the second minus the first (headings -2).
    expect(subtractCounts(SOURCE, PAGE)).toEqual({
      headings: 2,
      headingsByLevel: [0, 1, 2, -1, 0, 0],
      tables: -1,
      cells: 2,
      items: 0,
      rows: -5,
    });
  });

  it("undoes addCounts", () => {
    expect(subtractCounts(addCounts(SOURCE, PAGE), PAGE)).toEqual(SOURCE);
  });
});

describe("isZeroCounts", () => {
  it("is true of the empty counts and of a count minus itself", () => {
    expect(isZeroCounts(emptyCounts())).toBe(true);
    expect(isZeroCounts(subtractCounts(SOURCE, SOURCE))).toBe(true);
  });

  it.each(["headings", "tables", "cells", "items", "rows"] as const)(
    "is false when %s alone is not zero, whatever its sign",
    (field) => {
      expect(isZeroCounts({ ...emptyCounts(), [field]: 1 })).toBe(false);
      expect(isZeroCounts({ ...emptyCounts(), [field]: -1 })).toBe(false);
    },
  );

  it("is false when one heading LEVEL alone is not zero (two levels that cancel in the total are still a gap)", () => {
    // One heading of level 2 became a heading of level 3: the totals agree, the levels do not.
    const moved = subtractCounts(
      { ...emptyCounts(), headings: 1, headingsByLevel: [0, 1, 0, 0, 0, 0] },
      { ...emptyCounts(), headings: 1, headingsByLevel: [0, 0, 1, 0, 0, 0] },
    );
    expect(moved.headings).toBe(0);
    expect(isZeroCounts(moved)).toBe(false);
  });
});
