import { describe, expect, it } from "vitest";

import type { Box } from "../layout/slide-elements";
import type { PptxTextShape } from "./pptx-reader";
import { renderSlideSvg } from "./slide-svg";

// Characterisation of the steps `drawText` is cut into (outline, text rectangle, run sizes, line layout, line emission). Written green on
// the code before the cut. Every expected number is worked out by hand from the preset's rectangle, the insets (7.2 pt left and right, 3.6 pt top
// and bottom by default), 1.2 em lines and the advance classes (a lower-case letter 0.52 em, a narrow one 0.25 em, m and w 0.8 em).
const SLIDE = { cx: 12_700_000, cy: 6_350_000 }; // 1 000 × 500 pt
const box = (x: number, y: number, w: number, h: number): Box => ({
  x: x * 12_700,
  y: y * 12_700,
  cx: w * 12_700,
  cy: h * 12_700,
});
let rank = 0;
const text = (t: string, extra: Partial<PptxTextShape> = {}): PptxTextShape => ({
  kind: "text",
  spTreeIndex: rank++,
  box: box(10, 10, 200, 100),
  geometry: "rect",
  fill: "none",
  hasVisibleLine: true,
  paragraphs: [{ runs: [{ text: t }] }],
  ...extra,
});
const sized = (t: string, size: number, extra: Partial<PptxTextShape> = {}): PptxTextShape =>
  text(t, { paragraphs: [{ runs: [{ text: t, size }] }], ...extra });

/** The `<text>` elements: [x, y, font-size, content]. */
const texts = (svg: string): string[][] =>
  [
    ...svg.matchAll(/<text x="([^"]*)" y="([^"]*)" font-family="sans-serif" font-size="([^"]*)"[^>]*>(.*?)<\/text>/g),
  ].map((m) => [m[1]!, m[2]!, m[3]!, m[4]!]);

describe("drawText stages: the outline", () => {
  it("a rectangle has no corner; a round rectangle's radius is its adj of the shorter side, clamped to [0, 50 000]", () => {
    const shape = (geometry: string, adj?: number) =>
      sized("hi", 10, { geometry, box: box(100, 50, 200, 100), ...(adj === undefined ? {} : { adj }) });
    const rect = renderSlideSvg([shape("rect")], SLIDE).svg;
    expect(rect).toContain(
      '<rect x="100" y="50" width="200" height="100" fill="none" stroke="#000" stroke-width="1"/>',
    );
    expect(renderSlideSvg([shape("roundRect")], SLIDE).svg).toContain(
      '<rect x="100" y="50" width="200" height="100" rx="16.67" fill="none"',
    );
    expect(renderSlideSvg([shape("roundRect", 80_000)], SLIDE).svg).toContain('rx="50" fill="none"');
    expect(renderSlideSvg([shape("roundRect", -100)], SLIDE).svg).toContain('rx="0" fill="none"');
    expect(renderSlideSvg([shape("roundRect", 30_000)], SLIDE).svg).toContain('rx="30" fill="none"');
  });

  it("an ellipse is drawn with its radii", () => {
    const { svg } = renderSlideSvg([sized("hi", 10, { geometry: "ellipse", box: box(100, 50, 200, 100) })], SLIDE);
    expect(svg).toContain('<ellipse cx="200" cy="100" rx="100" ry="50" fill="none" stroke="#000" stroke-width="1"/>');
  });
});

describe("drawText stages: the text rectangle", () => {
  it("an ellipse's text starts (1 − cos 45°) / 2 of its width and of its height inside it, on each axis", () => {
    // x: 100 + 200 × 0.14645 + 7.2 = 136.49; first baseline: 50 + 100 × 0.14645 + 3.6 + 10 = 78.24.
    const { svg } = renderSlideSvg([sized("hi", 10, { geometry: "ellipse", box: box(100, 50, 200, 100) })], SLIDE);
    expect(texts(svg)).toEqual([["136.49", "78.24", "10", "hi"]]);
  });

  it("a round rectangle's text starts 29.289 % of its radius inside it", () => {
    // Radius 100 × 0.16667 = 16.667; inset 4.8816: x = 100 + 4.8816 + 7.2 = 112.08; baseline 50 + 4.8816 + 3.6 + 10 = 68.48.
    const { svg } = renderSlideSvg([sized("hi", 10, { geometry: "roundRect", box: box(100, 50, 200, 100) })], SLIDE);
    expect(texts(svg)).toEqual([["112.08", "68.48", "10", "hi"]]);
  });

  it("the right inset narrows the line and the bottom inset lowers the bottom anchor, each by its own default", () => {
    // t.w = 58.4 − 7.2 − 7.2 = 44.0 and « aaaa bbbb » measures 44.38: two lines (with a right inset of 3.6 it would be one).
    const two = renderSlideSvg([sized("aaaa bbbb", 10, { box: box(100, 50, 58.4, 40) })], SLIDE).svg;
    expect(texts(two).map((t) => t[3])).toEqual(["aaaa", "bbbb"]);
    // Bottom anchor: t.y 13.6, t.h 100 − 3.6 − 3.6 = 92.8, line 12 high: top 94.4, baseline 104.4 (bottom inset 7.2 would give 100.8).
    const bottom = renderSlideSvg([sized("hi", 10, { anchor: "b" })], SLIDE).svg;
    expect(texts(bottom)[0]![1]).toBe("104.4");
  });

  it("an ellipse's room for the text is its inscribed rectangle: a line leaves it before it leaves the box", () => {
    const room = (extra: Partial<PptxTextShape>, content: string): number =>
      renderSlideSvg([sized(content, 10, { geometry: "ellipse", ...extra })], SLIDE).overflowingText;
    // Room: y 24.64 to 95.36 (the box is 10 to 110), width 200 − 2 × 29.29 = 141.42. One 10 pt line at the bottom anchor has its middle at 85.75.
    expect(room({ anchor: "b" }, "hi")).toBe(0);
    // Seven lines from the top: the last middle is 28.24 + 72 + 6 = 106.24, past 95.36 but inside the box.
    expect(room({}, Array.from({ length: 7 }, () => "w").join("\n"))).toBe(1);
    expect(room({}, Array.from({ length: 5 }, () => "w").join("\n"))).toBe(0);
    // Eighteen « m » (0.8 em at 10 pt) make a word 144 pt wide: past 141.42, inside the 200 pt box. A word is never cut.
    expect(room({}, "m".repeat(18))).toBe(1);
    expect(room({}, "m".repeat(17))).toBe(0);
  });
});

describe("drawText stages: run sizes", () => {
  it("a blank run with no size does not leave the text out", () => {
    const shape = text("x", { paragraphs: [{ runs: [{ text: "  " }, { text: "hi", size: 10 }] }] });
    const { svg, omitted } = renderSlideSvg([shape], SLIDE);
    expect(omitted).toEqual({});
    expect(texts(svg).map((t) => t[3])).toEqual(["  hi"]);
    const visible = text("x", { paragraphs: [{ runs: [{ text: "hi" }] }] });
    expect(renderSlideSvg([visible], SLIDE).omitted).toEqual({ "size-not-declared": 1 });
  });
});

describe("drawText stages: line layout", () => {
  const baselines = (shape: PptxTextShape, defaults: Record<number, number> = {}): string[] =>
    texts(renderSlideSvg([shape], SLIDE, defaults).svg).map((t) => t[1]!);

  it("a line with no visible text is as high as its paragraph's largest known size", () => {
    // « a », a blank line, « b » at 10 pt, level-1 default 20 pt: the blank line is 12 high, not 24. Baselines 13.6 + 10, 25.6 + 12 + 10.
    expect(baselines(sized("a\n\nb", 10), { 0: 20 })).toEqual(["23.6", "47.6"]);
    // Runs of 10 and 20 pt: the blank line takes the LARGEST, 24 high: 13.6 + 12 + 24 + 20.
    const mixed = text("x", {
      paragraphs: [
        {
          runs: [
            { text: "a\n", size: 10 },
            { text: "\nb", size: 20 },
          ],
        },
      ],
    });
    expect(baselines(mixed)).toEqual(["23.6", "69.6"]);
  });

  it("an empty paragraph is as high as its level's default size, else the shape's first sized run", () => {
    const shape = text("x", {
      paragraphs: [{ runs: [{ text: "a", size: 10 }] }, { runs: [] }, { runs: [{ text: "b", size: 10 }] }],
    });
    expect(baselines(shape, { 0: 20 })).toEqual(["23.6", "59.6"]);
    expect(baselines(shape)).toEqual(["23.6", "47.6"]);
  });

  it("the shrink factor scales the blank line too, and a line spacing reduction never makes a line negative", () => {
    // fontScale 0.5: a 5 pt « a » (6 high), a blank line 6 high, « b » at 5 pt: 13.6 + 5, 25.6 + 5.
    expect(baselines(sized("a\n\nb", 10, { fontScale: 0.5 }))).toEqual(["18.6", "30.6"]);
    // 20 % spacing less a 50 % reduction would be −30 %: the line is 0 high, both baselines sit at 23.6.
    const squeezed = text("x", {
      lineSpaceReduction: 0.5,
      paragraphs: [{ runs: [{ text: "one\ntwo", size: 10 }], lineSpacing: { percent: 0.2 } }],
    });
    expect(baselines(squeezed)).toEqual(["23.6", "23.6"]);
  });

  it("a line is as high as the largest VISIBLE character on it: a large blank in the middle counts for nothing", () => {
    const shape = text("x", {
      paragraphs: [
        {
          runs: [
            { text: "a", size: 10 },
            { text: "   ", size: 30 },
            { text: "b", size: 10 },
          ],
        },
      ],
    });
    const [line] = texts(renderSlideSvg([shape], SLIDE).svg);
    expect([line![1], line![2]]).toEqual(["23.6", "10"]);
  });

  it("a line exactly as wide as the room is not an overflow", () => {
    // « iiii » at 8 pt: 4 × 0.25 × 8 = 8 pt in a box 8 pt wide with no written inset.
    const box8 = { box: box(100, 50, 8, 40), insets: { left: 0, right: 0 } };
    expect(renderSlideSvg([sized("iiii", 8, box8)], SLIDE).overflowingText).toBe(0);
    expect(renderSlideSvg([sized("iiiii", 8, box8)], SLIDE).overflowingText).toBe(1);
  });
});

describe("drawText stages: line emission", () => {
  it("a blank line draws nothing", () => {
    const { svg } = renderSlideSvg([sized("a\n\nb", 10)], SLIDE);
    expect(texts(svg).map((t) => t[3])).toEqual(["a", "b"]);
  });

  it("text that rises above its box at the bottom anchor is an overflow, though nothing falls below it", () => {
    // Four 7 pt lines (33.6 pt) bottom-anchored in a 20 pt box: the first line's middle is at −3, above the box (10).
    const lines = Array.from({ length: 4 }, () => "w").join("\n");
    expect(renderSlideSvg([sized(lines, 7, { box: box(10, 10, 200, 20), anchor: "b" })], SLIDE).overflowingText).toBe(
      1,
    );
    expect(renderSlideSvg([sized(lines, 7, { box: box(10, 10, 200, 20), anchor: "ctr" })], SLIDE).overflowingText).toBe(
      1,
    );
    expect(renderSlideSvg([sized("w", 7, { box: box(10, 10, 200, 20), anchor: "b" })], SLIDE).overflowingText).toBe(0);
  });

  it("centred and right-aligned text that spills past a narrow box widens the drawing to where the text really goes", () => {
    const wide = (align: "ctr" | "r") =>
      text("x", {
        box: box(100, 50, 20, 20),
        isUnwrapped: true,
        paragraphs: [{ runs: [{ text: "mmmmmmmm", size: 10 }], align }],
      });
    // 64 pt wide. Centred on 110 (the text rectangle's middle): 78 to 142; the box 100 to 120; a 6 pt margin: x from 72.
    expect(renderSlideSvg([wide("ctr")], SLIDE).svg).toContain('viewBox="72 44 76 32"');
    // Right-aligned at 112.8: 48.8 to 112.8: x from 42.8.
    expect(renderSlideSvg([wide("r")], SLIDE).svg).toContain('viewBox="42.8 44 83.2 32"');
  });

  it("pieces of another size are tspans; one piece is the text itself", () => {
    const shape = text("x", {
      paragraphs: [
        {
          runs: [
            { text: "big ", size: 20 },
            { text: "small", size: 10 },
          ],
        },
      ],
    });
    expect(renderSlideSvg([shape], SLIDE).svg).toContain(
      'font-size="20" xml:space="preserve"><tspan>big </tspan><tspan font-size="10">small</tspan></text>',
    );
  });
});

describe("drawText stages: colours", () => {
  it("a text of known colour on a box whose fill is unknown stays black, even inside a box drawn filled before it", () => {
    const under = text("x", {
      box: box(10, 10, 300, 200),
      fill: "#FFFF00",
      paragraphs: [{ runs: [{ text: "A", size: 10, color: "#000000" }] }],
    });
    const over = (fill: PptxTextShape["fill"]) =>
      text("x", {
        box: box(50, 50, 100, 50),
        fill,
        paragraphs: [{ runs: [{ text: "B", size: 10, color: "#FF0000" }] }],
      });
    const unknown = renderSlideSvg([under, over("unknown")], SLIDE).svg;
    expect(unknown).toMatch(/<text [^>]*>B<\/text>/);
    expect(unknown).not.toContain('fill="#FF0000"');
    // Fill « none » inside a filled box keeps the colour.
    expect(renderSlideSvg([under, over("none")], SLIDE).svg).toContain('fill="#FF0000"');
  });
});
