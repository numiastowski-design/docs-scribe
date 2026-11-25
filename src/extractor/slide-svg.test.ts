import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { XMLParser, XMLValidator } from "fast-xml-parser";
import { beforeAll, describe, expect, it } from "vitest";

import type { Box } from "../layout/slide-elements";
import { encodePng } from "doc-ir/internal";
import { extractPptx } from "./pptx";
import { type PptxReadResult, type PptxShape, type PptxSlide, type PptxTextShape, readPptx } from "./pptx-reader";
import { MAX_SVG_BYTES, renderSlideSvg } from "./slide-svg";

// Witnesses: the turned and grouped pictures of `fixtures/pptx/04` and `05`. Every value below was read in the slide's XML
// (ppt/slides/slide1.xml) before the test was written.
const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));
const TURNED = "pptx/04-Images-tournees.pptx";
const GROUPS = "pptx/05-Groupes-tournes.pptx";

/** Set D2C_SLIDE_SVG_OUT to a directory to write the witnesses' SVGs there
 *  for a look (never into the repository). */
const OUT = process.env["D2C_SLIDE_SVG_OUT"];

interface El {
  tag: string;
  attrs: Record<string, string>;
  text: string;
}

/** Every element of the SVG, parsed by fast-xml-parser (so it is XML), in
 *  document order; `text` is all the character data inside it, a `tspan`'s
 *  included. */
const elementsOf = (svg: string): El[] => {
  expect(XMLValidator.validate(svg)).toBe(true);
  const tree = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    parseTagValue: false,
    trimValues: false,
  }).parse(svg) as unknown[];
  const out: El[] = [];
  const walk = (nodes: unknown[]): void => {
    for (const node of nodes as Record<string, unknown>[]) {
      const tag = Object.keys(node).find((k) => k !== ":@")!;
      if (tag === "#text") continue;
      const kids = node[tag] as Record<string, unknown>[];
      const deep = (nodes: Record<string, unknown>[]): string =>
        nodes
          .map((k) =>
            k["#text"] !== undefined
              ? String(k["#text"] as string | number)
              : k["tspan"] !== undefined
                ? deep(k["tspan"] as Record<string, unknown>[])
                : "",
          )
          .join("");
      const text = deep(kids);
      out.push({ tag, attrs: (node[":@"] as Record<string, string> | undefined) ?? {}, text });
      walk(kids);
    }
  };
  walk(tree);
  return out;
};

/** Each deck read once for the whole file. */
const decks = new Map<string, Promise<PptxReadResult>>();
const deckOf = (path: string): Promise<PptxReadResult> => {
  if (!decks.has(path)) decks.set(path, readPptx(fixture(path)));
  return decks.get(path)!;
};

/** Cold loading happens here, not in the first test that awaits a deck: only this load gets 30 s, a starting value to
 *  calibrate. */
beforeAll(async () => {
  await Promise.all([deckOf(TURNED), deckOf(GROUPS)]);
}, 30_000);

const slideOf = async (path: string, n: number): Promise<{ slide: PptxSlide; size: { cx: number; cy: number } }> => {
  const { slides, slideSize } = await deckOf(path);
  return { slide: slides.find((s) => s.number === n)!, size: slideSize! };
};

const squash = (s: string): string => s.replace(/\s+/g, " ").trim();

/** The font size a drawn line of text is written at. */
const sizeOfText = (els: readonly El[], text: string): string | undefined =>
  els.find((e) => e.tag === "text" && e.text === text)?.attrs["font-size"];

/** What the drawing shows, counted. */
const tally = (els: readonly El[]) => ({
  lines: els.filter((e) => e.tag === "line").length,
  polylines: els.filter((e) => e.tag === "polyline").length,
  endArrows: els.filter((e) => (e.tag === "line" || e.tag === "polyline") && e.attrs["marker-end"] === "url(#e)")
    .length,
  startArrows: els.filter((e) => (e.tag === "line" || e.tag === "polyline") && e.attrs["marker-start"] === "url(#s)")
    .length,
  // The first rect is the white background.
  outlines: els.filter((e) => e.tag === "rect" || e.tag === "ellipse").length - 1,
  images: els.filter((e) => e.tag === "image").length,
});

const drawnText = (els: readonly El[]): string =>
  squash(
    els
      .filter((e) => e.tag === "text")
      .map((e) => e.text)
      .join(" "),
  );

const writeOut = (name: string, svg: string): void => {
  if (OUT === undefined) return;
  mkdirSync(OUT, { recursive: true });
  writeFileSync(join(OUT, name), svg);
};

describe("slide-svg — turned and mirrored pictures (25/09/2026)", () => {
  it("04-Images-tournees: the six copies of one picture, each with the transform computed by hand from slide1.xml", async () => {
    const { slide, size } = await slideOf(TURNED, 1);
    const { svg, omitted, picturesNotDrawn } = renderSlideSvg(
      slide.shapes,
      size,
      (await deckOf(TURNED)).defaultTextSizes,
    );
    writeOut("04-Images-tournees.svg", svg);
    const images = elementsOf(svg).filter((e) => e.tag === "image");
    // Every `a:ext` is 2 160 000 × 1 440 000 EMU: 170.08 × 113.39 pt
    // (EMU / 12 700, two decimals). Each image stands at its `a:off`, the
    // frame BEFORE rotation; the centre is off + ext / 2.
    // 1. 0°: off (3 240 000, 1 440 000) → (255.12, 113.39). No transform.
    // 2. rot="1800000": 1 800 000 / 60 000 = 30°. off (6 120 000, 1 440 000)
    //    → (481.89, 113.39); centre (7 200 000, 2 160 000) → (566.93, 170.08).
    // 3. rot="5400000": 90°. off (9 000 000, 1 440 000) → (708.66, 113.39);
    //    centre (10 080 000, 2 160 000) → (793.7, 170.08).
    // 4. flipH: off (3 240 000, 4 140 000) → (255.12, 325.98); centre
    //    (4 320 000, 4 860 000) → (340.16, 382.68). The mirror about x = cx is
    //    x ↦ 2·cx − x: translate(2·cx 0) scale(-1 1), 2·cx = 8 640 000 / 12 700
    //    = 680.31.
    // 5. rot="1800000" flipH: off (6 120 000, 4 140 000) → (481.89, 325.98);
    //    centre (7 200 000, 4 860 000) → (566.93, 382.68); mirror first
    //    (2·cx = 14 400 000 / 12 700 = 1 133.86), then 30° about the centre:
    //    SVG applies the list right to left, so rotate is written first.
    // 6. flipV: off (9 000 000, 4 140 000) → (708.66, 325.98); centre
    //    (10 080 000, 4 860 000) → (793.7, 382.68); y ↦ 2·cy − y,
    //    2·cy = 9 720 000 / 12 700 = 765.35.
    // Failure value: 1 image (the five turned or flipped ones left out).
    expect(
      images.map((e) => [e.attrs["x"], e.attrs["y"], e.attrs["width"], e.attrs["height"], e.attrs["transform"]]),
    ).toEqual([
      ["255.12", "113.39", "170.08", "113.39", undefined],
      ["481.89", "113.39", "170.08", "113.39", "rotate(30 566.93 170.08)"],
      ["708.66", "113.39", "170.08", "113.39", "rotate(90 793.7 170.08)"],
      ["255.12", "325.98", "170.08", "113.39", "translate(680.31 0) scale(-1 1)"],
      ["481.89", "325.98", "170.08", "113.39", "rotate(30 566.93 382.68) translate(1133.86 0) scale(-1 1)"],
      ["708.66", "325.98", "170.08", "113.39", "translate(0 765.35) scale(1 -1)"],
    ]);
    // Failure value: turned 5.
    expect(omitted.turned).toBeUndefined();
    expect(omitted["picture-group-stretched"]).toBeUndefined();
    expect(picturesNotDrawn).toEqual([]);
  });

  it("05-Groupes-tournes: (a) and (c) drawn through their group's matrix computed by hand from slide1.xml, (b) a marker, (a)'s text left out but kept in the collapsed text", async () => {
    const { slide, size } = await slideOf(GROUPS, 1);
    const { svg, omitted, picturesNotDrawn } = renderSlideSvg(
      slide.shapes,
      size,
      (await deckOf(GROUPS)).defaultTextSizes,
    );
    writeOut("05-Groupes-tournes.svg", svg);
    const els = elementsOf(svg);
    const images = els.filter((e) => e.tag === "image");
    // Each image stands at its `a:off`/`a:ext` as written, in the group's
    // coordinates, EMU / 12 700, two decimals; the matrix carries it to the
    // slide, its own rot and flips included (a, b, c, d six decimals; e, f
    // in points, two decimals).
    // (a) Group 5: off = chOff (2 520 000, 1 800 000), ext = chExt
    //     (2 160 000, 1 872 000): no scale. rot="1800000" (30°) flipH about
    //     its centre (3 600 000, 2 736 000). Mirror first (x ↦ −x), then 30°
    //     clockwise: a = cos 30°·(−1) = −0.866025, b = sin 30°·(−1) = −0.5,
    //     c = −sin 30° = −0.5, d = cos 30° = 0.866025.
    //     e = cx − a·cx − c·cy = 3 600 000 + 0.8660254 × 3 600 000
    //       + 0.5 × 2 736 000 = 8 085 691.45 EMU → 636.67 pt;
    //     f = cy − b·cx − d·cy = 2 736 000 + 0.5 × 3 600 000
    //       − 0.8660254 × 2 736 000 = 2 166 554.5 EMU → 170.59 pt.
    //     The picture itself (rank 5): off (2 520 000, 1 800 000) → (198.43,
    //     141.73), ext (2 160 000, 1 440 000) → 170.08 × 113.39, no rot.
    // (c) Group 12: ext cx 3 240 000 for chExt cx 2 160 000: sx = 1.5, sy = 1;
    //     off = chOff = (9 360 000, 1 800 000): x ↦ 1.5·x − 4 680 000.
    //     Picture 13: rot="10800000" (180°) about its centre (10 440 000,
    //     2 520 000): (x, y) ↦ (20 880 000 − x, 5 040 000 − y), diagonal, so
    //     drawn. Composed: a = 1.5 × −1 = −1.5, b = c = 0, d = −1;
    //     e = 1.5 × 20 880 000 − 4 680 000 = 26 640 000 EMU → 2 097.64 pt;
    //     f = 5 040 000 EMU → 396.85 pt. The image: off (9 360 000,
    //     1 800 000) → (737.01, 141.73). Check: its x edges 737.01 and 907.09
    //     land on 992.13 and 737.01, 255.12 pt wide = 3 240 000 EMU.
    // (b) Group 9: the same stretch around picture 10 turned 90°: below the
    //     uneven level the matrix is not diagonal → no image, a marker.
    // Failure value (until c98b204): (a) left out ("group-turned"); (b) drawn
    // in its stretched frame, 255.12 × 113.39, then "rotate(90 581.1
    // 198.43)": a tall picture where PowerPoint shows a square; (c) the same
    // way, "rotate(180 864.57 198.43)".
    expect(
      images.map((e) => [e.attrs["x"], e.attrs["y"], e.attrs["width"], e.attrs["height"], e.attrs["transform"]]),
    ).toEqual([
      ["198.43", "141.73", "170.08", "113.39", "matrix(-0.866025 -0.5 -0.5 0.866025 636.67 170.59)"],
      ["737.01", "141.73", "170.08", "113.39", "matrix(-1.5 0 0 -1 2097.64 396.85)"],
    ]);
    // Ranks: title 0, boxes 1–2, connector 3, group 4 (picture 5, text box
    // 6), caption 7, group 8 (picture 9), caption 10, group 11 (picture 12).
    expect(picturesNotDrawn).toEqual([{ spTreeIndex: 9, reason: "group-stretched" }]);
    // The title is the tab's heading, not a shape of the drawing. The text
    // box of (a), turned and mirrored by its group, is left out whole:
    // PowerPoint turns its text but never mirrors the glyphs (the maintainer's
    // screenshot, 25/09). Failure value: { "picture-group-turned": 1 }.
    expect(omitted).toEqual({ turned: 1, "picture-group-stretched": 1 });
    expect(drawnText(els)).not.toContain("texte dans le groupe");
    expect(drawnText(els)).toContain("Départ");
    // Its words stay in the collapsed text — hoisted AFTER the tabs since
    // option A (27/09/2026), in the section of its slide (one expand for
    // the deck since 29/09/2026).
    const { doc } = await extractPptx(fixture(GROUPS));
    const collapsed = doc.blocks.find((b) => b.kind === "diagramText");
    const section = collapsed?.kind === "diagramText" ? collapsed.sections[0] : undefined;
    expect(section?.inlines.map((i) => i.text).join("")).toContain("texte dans le groupe");
    expect(section?.label).toBe("Slide 1");
  });
});

// ---------------------------------------------------------------------------
// Text calibration witnesses (conception du 27/09/2026, purgée
// 27/09/2026): what PowerPoint fits, we fit; what overflows in PowerPoint
// still counts. Ground truth: PowerPoint's own export of each witness slide
// (COM, 27/09). Each shape is drawn ALONE: the overflow of a text shape does
// not depend on the others.

describe("slide-svg — text calibration, synthetic", () => {
  it("written insets place the text; a:lnSpc sets the line height, in percent or in points", () => {
    const two = (lineSpacing: { percent: number } | { points: number }) => [
      { runs: [{ text: "one\ntwo", size: 10 }], lineSpacing },
    ];
    const { svg } = renderSlideSvg(
      [
        // Left and top insets 0: x = 10, first baseline 10 + 0 + 10 = 20;
        // 150 %: 10 × 1.2 × 1.5 = 18 below it, 38. Failure value (insets
        // not read, 1.2 flat): x 17.2, baselines 23.6 and 35.6.
        text("x", { insets: { left: 0, top: 0 }, paragraphs: two({ percent: 1.5 }) }),
        // Exactly 30 pt between baselines: 23.6 and 53.6. Failure value: 35.6.
        text("x", { box: box(300, 10, 200, 100), paragraphs: two({ points: 30 }) }),
      ],
      SLIDE,
    );
    expect(
      elementsOf(svg)
        .filter((e) => e.tag === "text")
        .map((e) => [e.text, e.attrs["x"], e.attrs["y"]]),
    ).toEqual([
      ["one", "10", "20"],
      ["two", "10", "38"],
      ["one", "307.2", "23.6"],
      ["two", "307.2", "53.6"],
    ]);
  });

  it("a line is as high as the largest text on it, not in its paragraph; a smaller run on it is a tspan of its own size", () => {
    const { svg } = renderSlideSvg(
      [
        text("x", {
          paragraphs: [
            {
              runs: [
                { text: "Big ", size: 20 },
                { text: "small", size: 10 },
                { text: "\n", size: 10 },
                { text: "tail", size: 10 },
              ],
            },
          ],
        }),
      ],
      SLIDE,
    );
    // Line 1 at 20 pt: baseline 10 + 3.6 + 20 = 33.6, 24 pt high. Line 2 at
    // 10 pt: 13.6 + 24 + 10 = 47.6. Failure value (the paragraph's 20 pt on
    // every line): « tail » at 20 pt, baseline 57.6, no tspan size.
    const els = elementsOf(svg);
    expect(els.filter((e) => e.tag === "text").map((e) => [e.text, e.attrs["font-size"], e.attrs["y"]])).toEqual([
      ["Big small", "20", "33.6"],
      ["tail", "10", "47.6"],
    ]);
    expect(els.filter((e) => e.tag === "tspan").map((e) => [e.text, e.attrs["font-size"]])).toEqual([
      ["Big ", undefined],
      ["small", "10"],
    ]);
  });

  it("a line is out when its middle leaves the shape: a third of a line past each edge is not counted, more than half is", () => {
    // IT 3 « OTA… »'s case: 7 pt lines (8.4 pt) centred in a 20 pt box.
    // Three lines, 25.2 pt: 2.6 pt out at each end, the middles inside.
    // Four, 33.6 pt: 6.8 pt out, the first line's middle 2.6 pt above the
    // box. Failure value: [1, 1] (the whole height compared with the box,
    // or with the text rectangle).
    const centred = (lines: number) =>
      text("x", {
        box: box(10, 10, 200, 20),
        anchor: "ctr",
        paragraphs: [{ runs: [{ text: Array.from({ length: lines }, () => "w").join("\n"), size: 7 }] }],
      });
    expect([
      renderSlideSvg([centred(3)], SLIDE).overflowingText,
      renderSlideSvg([centred(4)], SLIDE).overflowingText,
    ]).toEqual([0, 1]);
    // IT's legend labels: 9 pt at the top of a 17.2 pt box whose written
    // insets (7.2 pt each) leave 2.8 pt. The line eats the bottom padding and
    // ends 0.8 pt past the box. Failure value: 1 (the text rectangle as the limit).
    const label = text("x", {
      box: box(10, 10, 58.3, 17.2),
      insets: { left: 91_425, top: 91_425, right: 91_425, bottom: 91_425 },
      paragraphs: [{ runs: [{ text: "Digital", size: 9 }] }],
    });
    expect(renderSlideSvg([label], SLIDE).overflowingText).toBe(0);
  });

  it("a:spAutoFit: the written box is already the grown one, its height is not counted as overflowing", () => {
    // Three 10 pt lines, 36 pt, in a 30 pt high box.
    const tall = { box: box(10, 10, 200, 30), paragraphs: [{ runs: [{ text: "one\ntwo\nthree", size: 10 }] }] };
    // Failure value: 1 for the second (spAutoFit not read).
    expect([
      renderSlideSvg([text("x", tall)], SLIDE).overflowingText,
      renderSlideSvg([text("x", { ...tall, hasSpAutoFit: true })], SLIDE).overflowingText,
    ]).toEqual([1, 0]);
  });
});

// ---------------------------------------------------------------------------
// Synthetic shapes, for what the witnesses do not exercise.

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
/** The shape without one of its optional fields. */
const without = (shape: PptxTextShape, key: "geometry" | "box"): PptxTextShape => {
  const copy = { ...shape };
  delete copy[key];
  return copy;
};
const png = encodePng(1, 1, "rgb-24", [0, 0, 0])!;

describe("arbitrage 28/09 (7c) — the wrap width is bounded: a degenerate box cannot cost quadratic time", () => {
  it("in a box wider than MAX_WRAP_WIDTH_PT the text still wraps there: lines close, the cost stays linear", () => {
    // The defect is a SYNCHRONOUS quadratic loop, so a timeout witness
    // would hang the worker instead of failing; the witness is the SHAPE
    // of the output instead — with the bound the lines close (which is
    // exactly what keeps the measuring windows small: 40 000 words in
    // this box ran past 120 s unbounded, measured 28/09/2026), without
    // it everything sits on ONE line. ⚠️ The run MUST carry a size:
    // without one the text is never drawn at all (« a size is never
    // invented ») and this witness exercises nothing — seen
    // green-for-the-wrong-reason while writing it (28/09/2026).
    const words = Array.from({ length: 5_000 }, (_, i) => `w${i}`).join(" ");
    const huge = text("x", {
      box: { x: 0, y: 0, cx: 1_000_000_000_000, cy: 12_700_000 },
      paragraphs: [{ runs: [{ text: words, size: 18 }] }],
    });
    const { svg } = renderSlideSvg([huge], SLIDE);
    expect(svg).toContain("w4999");
    expect((svg.match(/<text[ >]/g) ?? []).length).toBeGreaterThan(1);
  });
});

describe("arbitrage 28/09 (8b) — a paragraph's endSize saves a box whose runs carry no size", () => {
  it("a placeholder box, runs sizeless: drawn at endSize; without it, omitted size-not-declared as before", () => {
    const box8b = (endSize?: number) =>
      text("x", {
        placeholder: { type: "body" },
        paragraphs: [{ runs: [{ text: "chatroom words" }], ...(endSize === undefined ? {} : { endSize }) }],
      });
    // Failure value: the text omitted although the file DECLARES its size
    // on the paragraph's end mark — the Chatroom box, whole text lost from
    // the drawing.
    const saved = renderSlideSvg([box8b(10)], SLIDE);
    expect(saved.svg).toContain("chatroom words");
    expect(saved.omitted["size-not-declared"]).toBeUndefined();
    const control = renderSlideSvg([box8b()], SLIDE);
    expect(control.svg).not.toContain("chatroom words");
    expect(control.omitted["size-not-declared"]).toBe(1);
  });
});

describe("slide-svg — synthetic", () => {
  it("escapes text: & < > and quotes stay text", () => {
    const { svg } = renderSlideSvg([text('R&D <core> "quoted"', { box: box(10, 10, 400, 100) })], SLIDE, { 0: 18 });
    expect(svg).toContain("R&amp;D &lt;core&gt; &quot;quoted&quot;");
    expect(drawnText(elementsOf(svg))).toBe('R&D <core> "quoted"');
  });

  it("every omitted case is counted, nothing of it drawn", () => {
    const shapes: PptxShape[] = [
      text("arrow words", { geometry: "rightArrow" }),
      without(text("free words"), "geometry"),
      text("turned words", { rot: 5_400_000 }),
      // Turned or mirrored by an enclosing group: left out whole too.
      text("group words", { isTurnedByGroup: true }),
      text("vertical words", { isTextTurned: true }),
      text("upside words", { flipV: true }),
      // A placeholder whose run writes no size: outline drawn, text left out.
      text("placeholder words", { placeholder: { type: "body", idx: "1" } }),
      // A free shape without a size, and no default declared (no third argument).
      text("sizeless words"),
      without(text("no box words"), "box"),
      {
        kind: "connector",
        spTreeIndex: rank++,
        box: box(0, 0, 10, 10),
        startArrow: false,
        endArrow: true,
        geometry: "curvedConnector2",
        hasVisibleLine: true,
        ends: { x1: 0, y1: 0, x2: 127_000, y2: 127_000 },
      },
      {
        kind: "connector",
        spTreeIndex: rank++,
        box: box(0, 0, 10, 10),
        startArrow: false,
        endArrow: true,
        geometry: "curvedConnector3",
        hasVisibleLine: true,
        ends: { x1: 0, y1: 0, x2: 127_000, y2: 127_000 },
      },
      { kind: "table", spTreeIndex: rank++, box: box(0, 0, 10, 10), table: { rows: [] } } as unknown as PptxShape,
      { kind: "graphic", spTreeIndex: rank++, box: box(0, 0, 10, 10), graphic: "smartart" },
      { kind: "picture", spTreeIndex: rank++, box: box(0, 0, 10, 10), notPublished: "external-link" },
      {
        kind: "picture",
        spTreeIndex: rank++,
        box: box(0, 0, 10, 10),
        image: { bytes: new Uint8Array([1, 0, 0, 0, 0x45, 0x4d, 0x46]), contentType: "image/x-emf" },
      },
      {
        kind: "picture",
        spTreeIndex: rank++,
        box: box(0, 0, 10, 10),
        groupStretched: true,
        image: { bytes: png, contentType: "image/png" },
      },
      // A picture with a box but no frame: never from the reader, counted all the same.
      { kind: "picture", spTreeIndex: rank++, box: box(0, 0, 10, 10), image: { bytes: png, contentType: "image/png" } },
    ];
    const { svg, omitted } = renderSlideSvg(shapes, SLIDE);
    expect(omitted).toEqual({
      geometry: 2,
      turned: 2,
      "turned-text": 2,
      "size-not-declared": 2,
      "no-box": 2,
      "connector-geometry": 2,
      frame: 2,
      "picture-not-published": 2,
      "picture-group-stretched": 1,
    });
    const els = elementsOf(svg);
    // The vertical, upside-down and sizeless shapes keep their outline; no text at all.
    expect(tally(els)).toEqual({ lines: 0, polylines: 0, endArrows: 0, startArrows: 0, outlines: 4, images: 0 });
    expect(drawnText(els)).toBe("");
    // No arrow drawn: no markers written.
    expect(svg).not.toContain("<marker");
  });

  it("round rectangle corner, ellipse, invisible edge and line, straight line and its heads", () => {
    const { svg } = renderSlideSvg(
      [
        text("", { geometry: "roundRect", adj: 50_000, box: box(0, 0, 100, 40) }),
        text("", { geometry: "roundRect", box: box(0, 0, 100, 40) }),
        text("", { geometry: "ellipse", box: box(0, 0, 100, 40) }),
        text("", { hasVisibleLine: false }),
        {
          kind: "connector",
          spTreeIndex: rank++,
          box: box(0, 0, 10, 10),
          startArrow: true,
          endArrow: false,
          geometry: "line",
          hasVisibleLine: true,
          ends: { x1: 127_000, y1: 0, x2: 0, y2: 254_000 },
        },
        {
          kind: "connector",
          spTreeIndex: rank++,
          box: box(0, 0, 10, 10),
          startArrow: true,
          endArrow: true,
          geometry: "straightConnector1",
          hasVisibleLine: false,
          ends: { x1: 0, y1: 0, x2: 1, y2: 1 },
        },
      ],
      SLIDE,
    );
    const els = elementsOf(svg);
    // Corner: half the shorter side at adj 50 000; 16.667 % by default.
    expect(els.filter((e) => e.tag === "rect").map((e) => e.attrs["rx"])).toEqual([undefined, "20", "6.67"]);
    expect(els.find((e) => e.tag === "ellipse")!.attrs).toMatchObject({ cx: "50", cy: "20", rx: "50", ry: "20" });
    // The line runs from its start (10, 0) to its end (0, 20); an invisible
    // one draws nothing and is no omission.
    expect(els.filter((e) => e.tag === "line").map((e) => e.attrs)).toEqual([
      { x1: "10", y1: "0", x2: "0", y2: "20", stroke: "#000", "stroke-width": "1", "marker-start": "url(#s)" },
    ]);
  });

  it("text: written size, wrapping inside the box, alignment, anchor, overflow drawn and counted", () => {
    // Box 200 pt wide, text area 200 − 2 × 7.2 = 185.6 pt. Advances by class
    // (calibration of 27/09/2026): « alpha beta gamma delta epsilon zeta
    // eta » is 15.56 em of letters and 6 spaces of 0.278 em, 17.23 em =
    // 172.3 pt at 10 pt; with « theta », 195.7 pt: it wraps there. Before,
    // at a flat 0.55 em, 33 characters fit: « …epsilon » / « zeta… ».
    const long = "alpha beta gamma delta epsilon zeta eta theta iota kappa";
    const { svg, overflowingText } = renderSlideSvg(
      [
        text(long, { paragraphs: [{ runs: [{ text: long, size: 10 }], align: "ctr" }], anchor: "b" }),
        // 18 pt declared default, 40 pt box: three lines cannot fit.
        text("x", { box: box(300, 10, 100, 40), paragraphs: [{ runs: [{ text: "one\ntwo\nthree" }] }] }),
      ],
      SLIDE,
      { 0: 18 },
    );
    const lines = elementsOf(svg).filter((e) => e.tag === "text");
    expect(lines.map((e) => e.text)).toEqual([
      "alpha beta gamma delta epsilon zeta eta",
      "theta iota kappa",
      "one",
      "two",
      "three",
    ]);
    expect(lines.map((e) => [e.attrs["font-size"], e.attrs["text-anchor"]])).toEqual([
      ["10", "middle"],
      ["10", "middle"],
      ["18", undefined],
      ["18", undefined],
      ["18", undefined],
    ]);
    // Anchored at the bottom: the last baseline 12 pt (one line) above the
    // text area's bottom plus the size: 110 − 3.6 − 12 + 10 = 104.4.
    expect(lines[1]!.attrs["y"]).toBe("104.4");
    // The second shape runs past its box and is drawn whole.
    expect(overflowingText).toBe(1);
  });

  it("a picture that would pass the cap is left out and counted; the next ones still fit", () => {
    const big = new Uint8Array(MAX_SVG_BYTES);
    big.set(png);
    const pic = (bytes: Uint8Array): PptxShape => ({
      kind: "picture",
      spTreeIndex: rank++,
      box: box(0, 0, 10, 10),
      frame: box(0, 0, 10, 10),
      image: { bytes, contentType: "image/png" },
    });
    const { svg, omitted } = renderSlideSvg([pic(png), pic(big), pic(png)], SLIDE);
    expect(omitted).toEqual({ "picture-over-cap": 1 });
    expect(tally(elementsOf(svg)).images).toBe(2);
    expect(Buffer.byteLength(svg)).toBeLessThanOrEqual(MAX_SVG_BYTES);
  });
});

/** 29/09/2026: the drawing is cropped to what it draws, plus a
 *  margin of 6 pt, shown at most 744 px wide and at most twice the scale
 *  of the whole slide (here 744 / 1 000 = 0.744 px per pt, so 1.488). */
describe("slide-svg — the drawing cropped to its useful area", () => {
  const frameOf = (
    svg: string,
  ): { viewBox: string; width: string; height: string; background: Record<string, string> } => {
    const els = elementsOf(svg);
    return {
      viewBox: els[0]!.attrs["viewBox"]!,
      width: els[0]!.attrs["width"]!,
      height: els[0]!.attrs["height"]!,
      background: els[1]!.attrs,
    };
  };

  it("a small drawing: viewBox = its box plus 6 pt, enlarged at most 2× (1.488 px per pt), the height derived; the white background fills the viewBox", () => {
    const { svg } = renderSlideSvg([text("", { box: box(100, 100, 200, 100) })], SLIDE);
    // Failure value (until 29/09/2026): "0 0 1000 500", 744 × 372.
    expect(frameOf(svg)).toEqual({
      viewBox: "94 94 212 112",
      width: "315.46",
      height: "166.66",
      background: { x: "94", y: "94", width: "212", height: "112", fill: "#fff" },
    });
  });

  it("a wide drawing is shown 744 px wide, never more; the viewBox never leaves the slide", () => {
    expect(frameOf(renderSlideSvg([text("", { box: box(0, 200, 1000, 20) })], SLIDE).svg)).toMatchObject({
      viewBox: "0 194 1000 32",
      width: "744",
      height: "23.81",
    });
    expect(frameOf(renderSlideSvg([text("", { box: box(0, 0, 1000, 500) })], SLIDE).svg)).toMatchObject({
      viewBox: "0 0 1000 500",
      width: "744",
      height: "372",
    });
  });

  it("text running past its box is inside the viewBox (never clipped): its estimated line, not its box", () => {
    // « alpha beta gamma delta epsilon » at 10 pt: 13.572 em = 135.72 pt
    // (ADVANCE_EM), from 100 + 7.2: it ends at 242.92, far past the box's
    // 150. The box itself is not drawn (no line, no fill). Top: 100 + 3.6;
    // bottom: the baseline 113.6 plus a 0.25 em descent.
    const shape = text("x", {
      box: box(100, 100, 50, 20),
      hasVisibleLine: false,
      isUnwrapped: true,
      paragraphs: [{ runs: [{ text: "alpha beta gamma delta epsilon", size: 10 }] }],
    });
    const { svg, overflowingText } = renderSlideSvg([shape], SLIDE);
    expect(overflowingText).toBe(1);
    expect(frameOf(svg).viewBox).toBe("101.2 97.6 147.72 24.5");
  });

  it("a turned picture counts by its turned corners; a connector by its ends", () => {
    const turned: PptxShape = {
      kind: "picture",
      spTreeIndex: rank++,
      box: box(100, 100, 200, 100),
      frame: box(100, 100, 200, 100),
      rot: 5_400_000,
      image: { bytes: png, contentType: "image/png" },
    };
    expect(frameOf(renderSlideSvg([turned], SLIDE).svg).viewBox).toBe("144 44 112 212");
    const line: PptxShape = {
      kind: "connector",
      spTreeIndex: rank++,
      box: box(0, 0, 10, 10),
      startArrow: false,
      endArrow: true,
      geometry: "line",
      hasVisibleLine: true,
      ends: { x1: 500 * 12_700, y1: 100 * 12_700, x2: 700 * 12_700, y2: 300 * 12_700 },
    };
    expect(frameOf(renderSlideSvg([line], SLIDE).svg).viewBox).toBe("494 94 212 212");
  });

  it("a straight connector with a box but no ends is counted, never left out in silence", () => {
    // Never from the reader (it gives `ends` whenever it gives a box): counted all the same, like a bent one without a path.
    const line: PptxShape = {
      kind: "connector",
      spTreeIndex: rank++,
      box: box(0, 0, 10, 10),
      startArrow: false,
      endArrow: true,
      geometry: "line",
      hasVisibleLine: true,
    };
    const { svg, omitted } = renderSlideSvg([line], SLIDE);
    // Failure value: `{}`, the line neither drawn nor counted.
    expect(omitted).toEqual({ "connector-geometry": 1 });
    expect(svg).not.toContain("<line");
    expect(svg).not.toContain("<defs>");
  });

  it("an invisible connector is neither drawn nor counted, whatever its route; a visible one that cannot be traced is counted", () => {
    const untraceable = (geometry: string, hasVisibleLine: boolean): PptxShape => ({
      kind: "connector",
      spTreeIndex: rank++,
      box: box(0, 0, 10, 10),
      startArrow: false,
      endArrow: true,
      geometry,
      hasVisibleLine,
    });
    // A straight one without ends, a bent one without vertices, a curved one: none can be traced.
    const routes = ["line", "bentConnector3", "curvedConnector3"];
    const hidden = renderSlideSvg(
      routes.map((geometry) => untraceable(geometry, false)),
      SLIDE,
    );
    // Failure value: `{ "connector-geometry": 2 }`, the bent and the curved ones counted though no line shows.
    expect(hidden.omitted).toEqual({});
    expect(tally(elementsOf(hidden.svg))).toMatchObject({ lines: 0, polylines: 0 });
    const shown = renderSlideSvg(
      routes.map((geometry) => untraceable(geometry, true)),
      SLIDE,
    );
    expect(shown.omitted).toEqual({ "connector-geometry": 3 });
    expect(tally(elementsOf(shown.svg))).toMatchObject({ lines: 0, polylines: 0 });
  });

  it("nothing drawn: the whole slide, as before", () => {
    expect(frameOf(renderSlideSvg([text("arrow words", { geometry: "rightArrow" })], SLIDE).svg)).toMatchObject({
      viewBox: "0 0 1000 500",
      width: "744",
      height: "372",
    });
  });
});

describe("slide-svg — step 7 bis, synthetic", () => {
  it("a bent connector is a polyline with its heads; a repeated vertex is dropped; without a path it is omitted", () => {
    const bent = (path: { x: number; y: number }[] | undefined): PptxShape => ({
      kind: "connector",
      spTreeIndex: rank++,
      box: box(0, 0, 10, 10),
      startArrow: true,
      endArrow: true,
      geometry: "bentConnector3",
      hasVisibleLine: true,
      ends: { x1: 0, y1: 0, x2: 1, y2: 1 },
      ...(path === undefined ? {} : { path }),
    });
    // adj1 0: the first two vertices coincide; the start head must follow
    // the first real segment. Failure value: 4 points.
    const pts = [
      { x: 0, y: 0 },
      { x: 0, y: 0 },
      { x: 0, y: 254_000 },
      { x: 127_000, y: 254_000 },
    ];
    const { svg, omitted } = renderSlideSvg([bent(pts), bent(undefined)], SLIDE);
    expect(
      elementsOf(svg)
        .filter((e) => e.tag === "polyline")
        .map((e) => e.attrs),
    ).toEqual([
      {
        points: "0,0 0,20 10,20",
        fill: "none",
        stroke: "#000",
        "stroke-width": "1",
        "marker-start": "url(#s)",
        "marker-end": "url(#e)",
      },
    ]);
    expect(omitted).toEqual({ "connector-geometry": 1 });
    expect(svg).toContain("<marker");
  });

  it("fontScale shrinks the size, lnSpcReduction the line height; wrap none keeps one line", () => {
    const long = "alpha beta gamma delta epsilon zeta eta theta iota kappa";
    const { svg } = renderSlideSvg(
      [
        // 20 pt × 0.5 = 10 pt; line height 10 × 1.2 × (1 − 0.2) = 9.6.
        // First baseline 10 + 3.6 + 10 = 23.6, second 33.2. Failure value:
        // font-size 20, second baseline 23.6 + 24 = 47.6.
        text("x", {
          paragraphs: [{ runs: [{ text: "one\ntwo", size: 20 }] }],
          fontScale: 0.5,
          lineSpaceReduction: 0.2,
        }),
        // 10 pt text 185.6 pt wide would wrap in two (see above); not here.
        text("x", {
          box: box(300, 10, 200, 100),
          paragraphs: [{ runs: [{ text: long, size: 10 }] }],
          isUnwrapped: true,
        }),
      ],
      SLIDE,
    );
    const lines = elementsOf(svg).filter((e) => e.tag === "text");
    expect(lines.map((e) => [e.text, e.attrs["font-size"], e.attrs["y"]])).toEqual([
      ["one", "10", "23.6"],
      ["two", "10", "33.2"],
      [long, "10", "23.6"],
    ]);
  });

  it("an unknown size: a free shape takes defaultTextStyle by level (else level 1's); a placeholder takes the size it inherits (29/09/2026), times fontScale; without one it is left out", () => {
    const para = (t: string, level?: number) => ({ runs: [{ text: t }], ...(level === undefined ? {} : { level }) });
    const { svg, omitted } = renderSlideSvg(
      [
        text("x", { box: box(10, 10, 400, 200), paragraphs: [para("zero"), para("one", 1), para("two", 2)] }),
        text("hidden size", { placeholder: { idx: "3" } }),
        text("x", {
          box: box(500, 10, 400, 100),
          placeholder: { type: "body" },
          paragraphs: [{ runs: [{ text: "sized", size: 12 }] }],
        }),
        // The reader's inherited sizes (`pptx-reader/text-size.ts`), never
        // `p:defaultTextStyle`'s for a placeholder: 20 × 0.5 at level 0.
        // Failure value (until 29/09/2026): omitted, size-not-declared 2.
        text("x", {
          box: box(500, 200, 400, 100),
          placeholder: { type: "subTitle", idx: "1" },
          inheritedSizes: { 0: 20 },
          fontScale: 0.5,
          paragraphs: [para("inherited")],
        }),
      ],
      SLIDE,
      // lvl1pPr and lvl3pPr declare a size; lvl2pPr does not.
      { 0: 14, 2: 10 },
    );
    const els = elementsOf(svg);
    // Failure values: 18 everywhere (the old guess), or 10 for "one".
    expect(["zero", "one", "two", "sized", "inherited"].map((t) => sizeOfText(els, t))).toEqual([
      "14",
      "14",
      "10",
      "12",
      "10",
    ]);
    expect(drawnText(els)).not.toContain("hidden size");
    expect(omitted).toEqual({ "size-not-declared": 1 });
    // The placeholder's outline is still drawn: 4 shapes, 4 outlines.
    expect(tally(els).outlines).toBe(4);
  });
});

// ---------------------------------------------------------------------------
// Step 7 ter — exact colours (25/09/2026; decision 12, MIKSUDOC « Spécification fonctionnelle » §9.3.3).

/** Each drawn line of text with its colour (`undefined`: black by default,
 *  or several colours in tspans). */
const textColours = (els: readonly El[]): { text: string; fill: string | undefined }[] =>
  els.filter((e) => e.tag === "text").map((e) => ({ text: e.text, fill: e.attrs["fill"] }));

describe("slide-svg — step 7 ter, colours, synthetic", () => {
  const run = (t: string, color?: `#${string}`) => ({ text: t, size: 10, ...(color === undefined ? {} : { color }) });

  it("a filled box is filled only when the colour of all its visible text is known, or it has no text; its outline only when its line shows", () => {
    const { svg } = renderSlideSvg(
      [
        // Filled, every run known, no line: fill, no stroke, white text.
        text("x", { fill: "#351C75", hasVisibleLine: false, paragraphs: [{ runs: [run("Okta", "#FFFFFF")] }] }),
        // Filled, a run of unknown colour: as today (outline, black text).
        text("x", {
          box: box(300, 10, 200, 100),
          fill: "#3C78D8",
          hasVisibleLine: false,
          paragraphs: [{ runs: [run("GCP")] }],
        }),
        // Filled, no text, line shown: fill and black stroke.
        text("", { box: box(600, 10, 100, 40), fill: "#B0D300", hasVisibleLine: true, paragraphs: [] }),
        // Unknown fill, no line: outline as today.
        text("", { box: box(600, 100, 100, 40), fill: "unknown", hasVisibleLine: false, paragraphs: [] }),
        // No fill, no line: nothing.
        text("", { box: box(600, 200, 100, 40), fill: "none", hasVisibleLine: false, paragraphs: [] }),
        // A filled ellipse without text.
        text("", {
          box: box(600, 300, 100, 40),
          geometry: "ellipse",
          fill: "#FFFFFF",
          hasVisibleLine: false,
          paragraphs: [],
        }),
      ],
      SLIDE,
    );
    const els = elementsOf(svg);
    // Failure value: every rect `fill="none"` (no colour drawn).
    expect(
      els
        .filter((e) => e.tag === "rect" || e.tag === "ellipse")
        .slice(1)
        .map((e) => [e.attrs["fill"], e.attrs["stroke"]]),
    ).toEqual([
      ["#351C75", undefined],
      ["none", "#000"],
      ["#B0D300", "#000"],
      ["none", "#000"],
      ["#FFFFFF", undefined],
    ]);
    expect(textColours(els)).toEqual([
      { text: "Okta", fill: "#FFFFFF" },
      { text: "GCP", fill: undefined },
    ]);
  });

  it("a line of runs in several colours is written in tspans; a known black stays written", () => {
    const { svg } = renderSlideSvg(
      [
        text("x", {
          box: box(10, 10, 400, 100),
          fill: "#EEEEEE",
          hasVisibleLine: false,
          paragraphs: [{ runs: [run("Level ", "#000000"), run("two", "#142420")] }],
        }),
      ],
      SLIDE,
    );
    const els = elementsOf(svg);
    expect(drawnText(els)).toBe("Level two");
    expect(els.filter((e) => e.tag === "tspan").map((e) => [e.text, e.attrs["fill"]])).toEqual([
      ["Level ", "#000000"],
      ["two", "#142420"],
    ]);
  });

  it("text without its own fill keeps its colour only on a filled box painted before it that contains its text rectangle; else black", () => {
    const bar = (x: number, fill: `#${string}` | "unknown" = "#3C78D8") =>
      text("", { box: box(x, 0, 200, 60), fill, hasVisibleLine: false, paragraphs: [] });
    const label = (x: number, w = 100) =>
      text("x", {
        box: box(x, 10, w, 40),
        fill: "none",
        hasVisibleLine: false,
        paragraphs: [{ runs: [run(`L${x}`, "#FFFFFF")] }],
      });
    const { svg } = renderSlideSvg(
      [
        bar(0),
        label(20), // on the bar painted before: white
        label(150), // runs past the bar's right edge (150 + 100 > 200): black
        label(320), // its bar is painted AFTER it: black
        bar(300),
        bar(600, "unknown"),
        label(620), // on a bar whose fill is not drawn: black
      ],
      SLIDE,
    );
    // Failure value: every label "#FFFFFF" (white on white for 150, 320, 620).
    expect(textColours(elementsOf(svg))).toEqual([
      { text: "L20", fill: "#FFFFFF" },
      { text: "L150", fill: undefined },
      { text: "L320", fill: undefined },
      { text: "L620", fill: undefined },
    ]);
  });
});

describe("🔴 review pptx 27/09, finding 3 — a paragraph of 200 000 runs is drawn without a spread", () => {
  it("the largest run size is found by a loop, never Math.max(...runs) (RangeError past ~124k arguments)", () => {
    const runs = Array.from({ length: 200_000 }, () => ({ text: "w", size: 10 }));
    const shape = text("x", { box: box(10, 10, 200, 100), isUnwrapped: true, paragraphs: [{ runs }] });
    let svg = "";
    // Failure value: RangeError: Maximum call stack size exceeded.
    expect(() => {
      svg = renderSlideSvg([shape], SLIDE).svg;
    }).not.toThrow();
    expect(drawnText(elementsOf(svg))).toBe("w".repeat(200_000));
  });
});

// SmartArt shapes turned by any angle (20/10/2026): a SmartArt shape of any angle is drawn as its UNTURNED frame under `rotate(θ cx cy)`.
describe("slide-svg — SmartArt shapes turned by any angle", () => {
  it("a bar 400 x 80 pt turned 45 degrees grows the drawing by its turned corners: (400 + 80) cos 45 = 339.41 pt plus the 6 pt margin on each side", () => {
    const side = 480 * Math.SQRT1_2; // 339.41 pt: the bounds of the turned bar, as `computeVisibleBox` gives them
    const frame = box(300, 150, 400, 80);
    const shape = text("", {
      box: box(500 - side / 2, 190 - side / 2, side, side),
      smartArt: { frame, turn: 2_700_000, textBox: frame, textTurn: 0 },
    });
    const { svg } = renderSlideSvg([shape], SLIDE);
    const viewBox = elementsOf(svg)[0]!.attrs["viewBox"]!.split(" ").map(Number);
    // Failure value: 412 (the unturned 400 pt plus the margin), or the whole slide when the turned shape is left out.
    expect(viewBox[2]).toBeCloseTo(side + 12, 1);
    expect(viewBox[3]).toBeCloseTo(side + 12, 1);
    // The outline is the UNTURNED frame under a rotate about its centre (500, 190).
    const rect = elementsOf(svg).find((e) => e.tag === "rect" && e.attrs["fill"] !== "#fff")!;
    expect(rect.attrs).toMatchObject({
      x: "300",
      y: "150",
      width: "400",
      height: "80",
      transform: "rotate(45 500 190)",
    });
  });
});

describe("slide-svg — the `chord` preset of a SmartArt shape, synthetic", () => {
  // A 200 × 100 pt frame at (100, 100): the ellipse's centre is (200, 150), its radii 100 and 50.
  const frame = box(100, 100, 200, 100);
  const chord = (guides: { adj1?: number; adj2?: number }): string => {
    const shape = text("", {
      box: frame,
      geometry: "chord",
      smartArt: { frame, turn: 0, textBox: frame, textTurn: 0, ...guides },
    });
    return renderSlideSvg([shape], SLIDE).svg;
  };
  const pathOf = (svg: string): string | undefined => /<path d="([^"]*)"/.exec(svg)?.[1];

  it("the default guides: from 45 to 270 degrees, the large arc, clockwise, closed", () => {
    // The point at 45 degrees is the ellipse's parameter atan2(100 sin 45, 50 cos 45) = 63.43 degrees: (244.72, 194.72).
    expect(pathOf(chord({}))).toBe("M244.72,194.72A100 50 0 1 1 200,100Z");
  });

  it("a half turn exactly is the small arc; one unit more is the large one", () => {
    expect(pathOf(chord({ adj1: 0, adj2: 10_800_000 }))).toBe("M300,150A100 50 0 0 1 100,150Z");
    // Failure value: the same flag 0, were the half turn compared with `>=` or another number.
    expect(pathOf(chord({ adj1: 0, adj2: 10_800_001 }))).toBe("M300,150A100 50 0 1 1 100,150Z");
  });

  it("guides out of range are pinned to [0, a full turn less one unit]: still a path, not the whole ellipse", () => {
    // Failure value: `undefined` (an ellipse), were the upper pin a full turn.
    expect(pathOf(chord({ adj1: -5, adj2: 30_000_000 }))).toBe("M300,150A100 50 0 1 1 300,150Z");
  });

  it("equal guides are a whole turn: the ellipse, no path", () => {
    const svg = chord({ adj1: 16_200_000, adj2: 16_200_000 });
    expect(pathOf(svg)).toBeUndefined();
    expect(svg).toContain('<ellipse cx="200" cy="150" rx="100" ry="50"');
  });
});
