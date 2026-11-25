import { describe, expect, it } from "vitest";

import type { PptxShape } from "../extractor/pptx-reader/index";
import { layoutSlide, type SlideElement, type SlideLayout, type SlideLayoutInput } from "./slide-layout";

// Made-up slides, for the rules of the layout.

const PT = 12_700;

/** A made-up element, its box in points. */
const el = (
  rank: number,
  box: [number, number, number, number],
  extra: Partial<Omit<SlideElement<number>, "item" | "box" | "spTreeIndex">> = {},
): SlideElement<number> => ({
  item: rank,
  box: { x: box[0] * PT, y: box[1] * PT, cx: box[2] * PT, cy: box[3] * PT },
  picture: false,
  spTreeIndex: rank,
  ...extra,
});

const SLIDE = { cx: 960 * PT, cy: 540 * PT };

const layout = (elements: SlideElement<number>[], more: Partial<SlideLayoutInput<number>> = {}): SlideLayout<number> =>
  layoutSlide({ slide: SLIDE, elements, drawnGeometries: [], connectors: 0, ...more });

/** Each row as its cells, "RxC:items" with R the rowspan, C the colspan, the
 *  items by rank. */
const cellsOf = (l: SlideLayout<number | PptxShape>): string[][] => {
  if (l.form !== "table") return [[l.form]];
  return l.rows.map((row) =>
    row.map(
      (c) => `${c.rowspan}x${c.colspan}:${c.items.map((i) => (typeof i === "number" ? i : i.spTreeIndex)).join(",")}`,
    ),
  );
};

describe("slide layout — rules on made-up slides", () => {
  it("left edges chaining at 2.5, 12, 21, 30 pt make no column: the guard abstains", () => {
    const slide = (lefts: number[]): SlideElement<number>[] => [
      ...lefts.map((x, i) => el(i + 1, [x, 50 * i, 100, 30])),
      el(9, [300, 0, 100, 30]),
    ];
    // Control: the same slide with aligned edges is a 4x2 table.
    const aligned = layout(slide([2.5, 2.5, 2.5, 2.5]));
    expect(aligned.form === "table" ? `${aligned.rows.length}x${aligned.columns}` : aligned.form).toBe("4x2");
    // Failure value: without the guard, the chain is one column 27.5 pt wide
    // beside the 300 pt one, the same 4x2 table.
    expect(layout(slide([2.5, 12, 21, 30]))).toEqual({ form: "flow", reason: "chained-edges", items: [1, 9, 2, 3, 4] });
  });

  it("a colspan meeting another cell's rowspan stops there (USA 26's boxes): nothing lost", () => {
    // A picture top right; a text under it reaching across both columns.
    const l = layout([el(1, [787.5, 5.1, 124.2, 97.7], { picture: true }), el(2, [83, 69.1, 839.8, 383])]);
    // Failure value: two overlapping spans, and `buildMergeMap` drops #2.
    expect(cellsOf(l)).toEqual([["1x1:", "2x1:1"], ["1x1:2"]]);
  });

  it("an element whose centre lies on a picture goes into the picture's cell, after it", () => {
    // The digit is declared BEFORE the picture and sits at its bottom right:
    // its own edges would make a column and a row of their own.
    const l = layout([
      el(1, [250, 250, 20, 20]),
      el(5, [100, 100, 200, 200], { picture: true }),
      el(6, [400, 100, 200, 50]),
    ]);
    expect(cellsOf(l)).toEqual([["1x1:5,1", "1x1:6"]]);
  });

  it("two elements at the same place come out by their rank in the shape tree", () => {
    const l = layout([el(7, [0, 0, 100, 30]), el(3, [0, 0, 100, 30]), el(9, [300, 0, 100, 30])]);
    // Failure value: the given order, 7 then 3.
    expect(cellsOf(l)).toEqual([["1x1:3,7", "1x1:9"]]);
    expect(layout([el(7, [0, 0, 100, 30]), el(3, [0, 0, 100, 30])])).toEqual({
      form: "flow",
      reason: "one-column",
      items: [3, 7],
    });
  });

  it("a block arrow makes a diagram, drawn only or with text; an arrow callout does not", () => {
    const two = (): SlideElement<number>[] => [el(1, [0, 0, 100, 30]), el(2, [300, 0, 100, 30])];
    expect(layout(two(), { drawnGeometries: ["upArrow"] }).form).toBe("diagram");
    expect(layout([...two(), el(3, [150, 0, 50, 30], { geometry: "rightArrow" })]).form).toBe("diagram");
    expect(layout(two(), { connectors: 1 }).form).toBe("diagram");
    expect(layout(two(), { drawnGeometries: ["rightArrowCallout", "rect", undefined] }).form).toBe("table");
  });

  it("a group is one element: its box the union, its members in geometric order", () => {
    // Members declared bottom first.
    const l = layout([
      el(2, [0, 100, 100, 30], { group: 1 }),
      el(3, [0, 0, 100, 30], { group: 1 }),
      el(4, [300, 0, 100, 30]),
    ]);
    expect(cellsOf(l)).toEqual([["1x1:3,2", "1x1:4"]]);
  });

  it("a background picture takes a first row across the table; an element without a box makes a flow", () => {
    const l = layout([el(1, [0, 0, 960, 540], { picture: true }), el(2, [0, 0, 100, 30]), el(3, [300, 0, 100, 30])]);
    expect(cellsOf(l)).toEqual([["1x2:1"], ["1x1:2", "1x1:3"]]);
    const noBox: SlideElement<number> = { item: 4, picture: false, spTreeIndex: 4 };
    expect(layout([el(2, [0, 0, 100, 30]), noBox, el(3, [300, 0, 100, 30])])).toEqual({
      form: "flow",
      reason: "no-box",
      items: [2, 3, 4],
    });
  });
});

describe("🔴 review pptx 27/09, finding 3 — the union of a group of 200 000 boxes", () => {
  it("is taken by a loop, never Math.min(...boxes) (RangeError past ~124k arguments), and gives the right rectangle", () => {
    // 200 000 members in ONE group, tiled inside [0, 0, 109, 17] pt; a lone
    // element at x 300 pt: with the union right, the slide is a 1x2 table.
    const members = Array.from({ length: 200_000 }, (_, i) => el(i + 1, [i % 100, (i % 7) * 2, 10, 5], { group: 1 }));
    let l: SlideLayout<number> | undefined;
    // Failure value: RangeError: Maximum call stack size exceeded.
    expect(() => {
      l = layoutSlide({
        slide: SLIDE,
        elements: [...members, el(300_001, [300, 0, 100, 30])],
        drawnGeometries: [],
        connectors: 0,
      });
    }).not.toThrow();
    if (l?.form !== "table") throw new Error(`expected a table, got ${l?.form}`);
    expect([l.rows.length, l.columns]).toEqual([1, 2]);
    expect(l.rows[0]![0]!.items).toHaveLength(200_000);
    expect(l.rows[0]![1]!.items).toEqual([300_001]);
  });
});
