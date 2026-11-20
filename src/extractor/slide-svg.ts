/** ONE slide drawn as an SVG picture, from what the reader declares: round
 *  and square rectangles and ellipses (their fill when its colour is exact,
 *  a black outline), the text of those shapes inside their box, straight connectors and bent ones
 *  (`bentConnector2` to `5`, a polyline) with their arrow heads, and the
 *  slide's pictures at their place and size, turned and mirrored as
 *  declared (`pictureTransform`), or through their groups' matrix
 *  (`pictureMatrix`). The shapes of a SmartArt's saved drawing
 *  (`PptxTextShape#smartArt`, `extractor/smartart-drawing.ts`) are drawn the same way, plus the `hexagon` and `chord` presets, outlines
 *  turned by any angle and text turned with them (`transform="rotate"` on its `<text>` lines); the words they draw are returned
 *  (`smartArtText`) for the control (`findUndrawnWords`). Everything else is
 *  OMITTED, never approximated (what cannot be drawn right is not drawn), and
 *  counted by reason, so the caller can say what is not drawn; for a deck,
 *  the text itself goes to the diagram text the conservation control reads
 *  (`extractor/pptx.ts`).
 *  Colours (MIKSUDOC « Spécification fonctionnelle » §9.3.3): a box's fill and its
 *  text's colour are drawn only when EXACT (`pptx-reader/color.ts`), a box
 *  is filled only when the colour of all its text is known, and text of
 *  known colour without a fill of its own keeps it only on a box drawn
 *  filled before it; anything not known is black on white, never white.
 *  Strokes and arrows stay black, the background white; the slide's own
 *  background is not read. The picture is an image, insensitive to the
 *  page's theme.
 *  The picture is cropped to what it draws, text past its box included
 *  (never clipped), plus a margin, inside the slide; shown at most
 *  `MAX_WIDTH_PX` wide and at most `MAX_ENLARGEMENT` times the whole slide's
 *  scale.
 *  Deterministic: the same shapes give the same string, byte for byte (no
 *  random id; the two arrow markers are always "s" and "e").
 *  NOT here: which slide is a diagram (`layout/slide-layout.ts`), reading the
 *  shapes (`extractor/pptx-reader/`), attaching the picture and the text in
 *  flow (`extractor/pptx.ts`), nor Google's own picture of a Google Slides
 *  deck (MIKSUDOC « Spécification fonctionnelle » §9.3.4, step 8).
 *  Called by: `extractor/pptx.ts`, `extractor/docx-reader/shapes.ts`, `extractor/docx-reader/drawing-objects.ts`. */

import { utf8Length } from "../lib/utf8";
import { escapeAttr } from "../lib/xml";
import type { Box, Size } from "../layout/slide-elements";
import { sniffImage } from "doc-ir/internal";
import type {
  PptxConnectorShape,
  PptxPictureNotPublished,
  PptxPictureShape,
  PptxShape,
  PptxTextShape,
} from "./pptx-reader";
import type { Hex } from "./pptx-reader/color";
import { EMU_PER_PT, FULL_TURN, QUARTER_TURN, type Transform } from "./pptx-reader/geometry";
import { applyMatrix } from "doc-ir/internal";
import { holdsRect, RECT_RULES, rawFromXywh } from "doc-ir/internal";

/** Why a shape is missing from the drawing, in whole or in part. */
export type SlideSvgOmission =
  /** A text shape whose geometry is not "rect", "roundRect" or "ellipse":
   *  another preset (an arrow, a flowchart disk…), a free-form `a:custGeom`,
   *  or none written. Outline and text both left out: the text rectangle of
   *  another geometry is not the box. */
  | "geometry"
  /** A text shape turned (`rot`), or turned or mirrored by an enclosing
   *  group (`isTurnedByGroup`). Left out whole: its outline is not its visible
   *  box, its text turns (and glyphs are never mirrored, so the group's
   *  matrix would draw them wrong). A turned or flipped picture is drawn. */
  | "turned"
  /** A text shape whose text is turned (`a:bodyPr vert` or `rot`) or upside
   *  down (`flipV`): its outline is drawn, its text left out. */
  | "turned-text"
  /** A VISIBLE connector that cannot be traced: other than a straight or a
   *  bent one (curved, unknown), a bent one the reader gave no vertices, or a
   *  straight one it gave no ends. An invisible one is never counted. */
  | "connector-geometry"
  /** A text shape whose runs write no size where none can be READ: a
   *  placeholder when neither its own list style, its layout, its master nor
   *  the master's text styles write one at that level
   *  (`pptx-reader/text-size.ts`), or a free shape when `p:defaultTextStyle`
   *  declares none. Its outline is drawn, its text left out: a size is never
   *  invented. */
  | "size-not-declared"
  /** A shape without a box on the slide (a placeholder inheriting it). */
  | "no-box"
  /** A table, chart, SmartArt, OLE object or other graphic frame. */
  | "frame"
  /** A picture the reader does not publish (external, missing, video,
   *  crop it cannot apply), or a file format a browser does not draw in an
   *  SVG (only PNG, JPEG and GIF are embedded). */
  | "picture-not-published"
  /** A picture that would take the SVG past `MAX_SVG_BYTES`. */
  | "picture-over-cap"
  /** A picture turned inside a group stretched unevenly
   *  (`pptx-reader/geometry.ts#placePicture`): left out, not approximated. */
  | "picture-group-stretched";

/** Why ONE picture is not in the drawing: the reader's reason, its format
 *  ("format": not PNG, JPEG nor GIF), the cap ("over-cap"), a turn inside a
 *  group stretched unevenly ("group-stretched"), or no place on the slide
 *  ("no-box"). For a deck, the caller says each one as a `picture-not-drawn`
 *  notice (`extractor/pptx.ts`). */
type PictureNotDrawn = PptxPictureNotPublished | "format" | "over-cap" | "group-stretched" | "no-box";

export interface SlideSvg {
  svg: string;
  /** Counts by reason; a reason that never occurred is absent. */
  omitted: Partial<Record<SlideSvgOmission, number>>;
  /** Text shapes whose text, as estimated here, runs past their box: a line
   *  wider than the box, or a line whose middle lies outside it (a height
   *  under `a:spAutoFit` is never counted); `textRectangle` says why.
   *  Informative only. Drawn whole all the same (never clipped): a cut word
   *  would be a lost one. */
  overflowingText: number;
  /** Every picture left out of the drawing, in shape-tree order. */
  picturesNotDrawn: { spTreeIndex: number; reason: PictureNotDrawn }[];
  /** The text lines drawn for the shapes of a SmartArt's drawing, in paint order: what the drawing shows, compared with the model's words
   *  (`extractor/smartart-drawing.ts#findUndrawnWords`). A line of any other shape is not here. */
  smartArtText: string[];
}

// One viewBox unit is one point (EMU_PER_PT, imported above).

/** Widest picture on the page, in CSS pixels: the width `renderer/storage.ts`
 *  gives every published image. A 960 pt slide is drawn 744 px wide —
 *  the inner width of a native-tab frame, measured 27/09/2026 (`renderer/storage.ts#TABLE_WIDTH_PX`):
 *  at 760 the drawing overflowed its tab by 12-16 px and put a horizontal
 *  scrollbar on the page. */
const MAX_WIDTH_PX = 744;

/** Largest enlargement of a cropped drawing over the scale of its whole
 *  slide. Starting value to calibrate: the drawing is
 *  cropped to what it draws, so a drawing that fills a third of its slide
 *  could be shown three times larger; at 2× a 960 pt slide shows 12 pt text
 *  at 18.6 px instead of 9.3, and a narrow or tall drawing does not become
 *  a page-high picture. No population measured on either side. */
const MAX_ENLARGEMENT = 2;

/** White space kept around the drawn area, in points: the half stroke of
 *  an outline (0.5) and the half width of an arrow head (3) fit inside it,
 *  with room for our text-width estimate's error. Starting value to
 *  calibrate. */
const CROP_MARGIN_PT = 6;

/** Depth of a sans-serif descender below the baseline, in em: Arial's
 *  0.212, rounded up. Only the crop reads it. */
const DESCENT_EM = 0.25;

/** Largest SVG, in bytes, pictures included. Starting value to calibrate:
 *  measured on the three witness slides, 26 KB (USA 20, 3 pictures), 129 KB
 *  (IT 2, 10) and 346 KB (USA 9, 14), each with every picture. 1 MB is about
 *  2.9 times the largest; no second population (a slide too heavy to
 *  attach or to display) was measured.
 *  Pictures are taken in shape-tree order; one that would pass the cap is
 *  left out and counted, the next ones still tried. */
export const MAX_SVG_BYTES = 1_000_000;

/** The text insets taken for a side `a:bodyPr` writes none of (`lIns`,
 *  `rIns`, `tIns`, `bIns`): 0.1 inch left and right, 0.05 inch top and
 *  bottom. A written inset is read (`PptxTextShape.insets`). */
const INSET_X_EMU = 91_440;
const INSET_Y_EMU = 45_720;

/** Widest line the wrap may build, in points. Starting value to calibrate:
 *  ten times the widest slide seen
 *  (1 000 pt) — a guard against the degenerate box, not a separation of
 *  two populations: an ordinary box sits far under it and its rendering
 *  does not move a byte. Unbounded, a box ECMA still allows (up to
 *  ~2×10⁹ pt) never lets a line close, the measuring window never resets,
 *  and the wrap's cost turns quadratic — 40 000 words ran past 120 s,
 *  measured 28/09/2026. */
const MAX_WRAP_WIDTH_PT = 10_000;

/** Estimated advance of one character, in em of its written size, by class.
 *  Starting values to calibrate. No real font is measured: the SVG is drawn in the READER's
 *  sans-serif, so an exact fit is impossible by nature; the aim is « what
 *  PowerPoint fits, we fit », not the pixel. The values are Arial's
 *  (Windows' sans-serif) widths rounded per class: narrow i l t f j . , : ;
 *  ' | ! 0.222-0.278, lower case 0.5-0.556, m w M W 0.72-0.94, capitals
 *  0.61-0.78, digits 0.556, space 0.278. Checked on the witnesses of §5
 *  (slide-svg.test.ts) against PowerPoint's export: IT 3 « RMS… PREPROD »
 *  wraps in 4 lines and « OTA… » in 3, as in PowerPoint; USA 24 in 5 where
 *  PowerPoint has 4 (its decks are in Calibri, narrower than Arial), still
 *  inside its box. Known miss: IT 2-4 « Locks (Dormakaba, Omnitec, Assa
 *  Abloy) », 5 pt, 3 lines here and 2 in PowerPoint, still counted.
 *  Tried and rejected: a flat 0.55 em per character, which made IT 3's
 *  « (Accurate then Duetto) » 5 lines and « Client » (6 pt) 19.8 pt wide
 *  where this table gives 14.7 pt. */
const ADVANCE_EM = { narrow: 0.25, lower: 0.52, wide: 0.8, capital: 0.66, digit: 0.556, space: 0.278 } as const;
const NARROW = new Set("iltfj.,:;'|!");
const WIDE = new Set("mwMW");

/** One character's estimated advance, in em. Any other character (another
 *  punctuation, another script) counts as a lower-case letter. */
const advanceEm = (character: string): number => {
  if (NARROW.has(character)) return ADVANCE_EM.narrow;
  if (WIDE.has(character)) return ADVANCE_EM.wide;
  if (character === " " || character === "\t") return ADVANCE_EM.space;
  if (character >= "0" && character <= "9") return ADVANCE_EM.digit;
  if (character !== character.toLowerCase()) return ADVANCE_EM.capital;
  return ADVANCE_EM.lower;
};

/** A SmartArt's text is laid out by Office in the theme's font before it is saved (Calibri in both fixtures), whose advances are
 *  narrower than the Arial-based table above: a line Office keeps whole (« Process 1 », « Proce 1.1 » at 33 pt in a 133 pt wide text
 *  rectangle) must not wrap here. Starting value to calibrate, two measures against Office's PDFs (20/10/2026): Calibri / table =
 *  125.6 / 144.0 = 0.87 for « Proce 1.1 » and 125.4 / 152.8 = 0.82 for « Process 1 »; 0.85 keeps both on one line and still wraps
 *  « Point 1 » (20 pt) in its 44 pt hexagon text rectangle, as the PDF does (57 pt of Calibri, 51 here). */
const SMARTART_ADVANCE = 0.85;

/** Two centres closer than this (points) are one: below the 0.01 pt an SVG number is written with. */
const SAME_CENTRE_PT = 0.005;

/** Baseline-to-baseline distance of single spacing, as a fraction of the
 *  font size: the usual one of a sans-serif face. A paragraph's `a:lnSpc`
 *  scales it (`spcPct`) or replaces it (`spcPts`); `a:normAutofit
 *  lnSpcReduction` is subtracted from the percentage. */
const LINE_HEIGHT_EM = 1.2;

/** "roundRect" corner when `adj` is not written: ECMA-376's preset default. */
const DEFAULT_ROUND_ADJ = 16_667;

/** One side of the text rectangle of an ellipse inside its box, per side:
 *  (1 − cos 45°) / 2 of the box, the preset's inscribed rectangle. */
const ELLIPSE_TEXT_INSET = (1 - Math.SQRT1_2) / 2;

/** A roundRect's text rectangle is inset by 29.289 % of its corner radius
 *  (the preset's `il = x1 · 29289 / 100000`). */
const ROUND_TEXT_INSET = 0.29289;

/** Exported for `docx-reader/shapes.ts#hasNonTextGraphics`: deciding whether a Word
 *  drawing is published as a picture (MIKSUDOC « Spécification fonctionnelle » §9.5) needs to
 *  know, BEFORE calling this file, whether a shape draws a box at all — restating the set there would drift. */
export const DRAWN_GEOMETRIES = new Set(["rect", "roundRect", "ellipse"]);
/** Presets drawn for a SmartArt's saved drawing ONLY: its `dsp:txXfrm` gives the text rectangle, which a shape of a slide would need the
 *  preset's own formula for. */
const SMARTART_GEOMETRIES = new Set(["hexagon", "chord"]);
/** The `hexagon` preset's defaults (ECMA-376 presetShapeDefinitions): `adj` 25 000 (inset of the slanted sides, 100 000ths of the shorter
 *  side) and `vf` 115 470 (vertical factor). */
const HEXAGON_ADJ = 25_000;
const HEXAGON_VF = 115_470;
const STRAIGHT_CONNECTORS = new Set(["line", "straightConnector1"]);
const BENT_CONNECTORS = new Set(["bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"]);

/** What the drawing does with a connector that has a box: its line is traced ("drawn"), it cannot be and is counted
 *  `connector-geometry` ("not-traceable"), or its line does not show and nothing is said ("invisible"). */
export type ConnectorFate = "drawn" | "not-traceable" | "invisible";

/** THE definition of a connector's fate, read by `drawConnector` below and by `docx-reader/shapes.ts#hasNonTextGraphics`,
 *  which decides whether a Word drawing gives a picture at all: one copy, so the two cannot disagree.
 *  ONE rule for every route: a line that does not show is no omission, so it is neither drawn nor counted; a line that
 *  shows is drawn when the reader gave what traces it — the ends of a straight one (its box alone does not say which
 *  diagonal it runs along once turned), the vertices of a bent one — and counted otherwise, a curved or unknown preset
 *  included. */
export const connectorFate = (shape: PptxConnectorShape): ConnectorFate => {
  if (!shape.hasVisibleLine) return "invisible";
  const geometry = shape.geometry ?? "";
  const traceable = STRAIGHT_CONNECTORS.has(geometry)
    ? shape.ends !== undefined
    : BENT_CONNECTORS.has(geometry) && shape.path !== undefined;
  return traceable ? "drawn" : "not-traceable";
};

/** A number of points written with at most two decimals; never "-0". */
const formatNumber = (value: number): string => String(Math.round(value * 100) / 100 + 0);
const toPoints = (emu: number): number => emu / EMU_PER_PT;

const STROKE = 'fill="none" stroke="#000" stroke-width="1"';

/** The MIME type the bytes themselves say, for the three formats every
 *  browser draws inside an SVG; `undefined` otherwise (EMF, TIFF, SVG…). */
const drawableType = (bytes: Uint8Array): string | undefined => {
  const format = sniffImage(bytes);
  return format === undefined ? undefined : `image/${format}`;
};

/** The SVG `transform` of a picture turned by `rot` and mirrored by
 *  `flipH`/`flipV`, its image placed at `frame` (points); "" when neither.
 *  Order: the flips FIRST, about the frame's centre, then the rotation,
 *  clockwise, about the same centre: the same order as
 *  `pptx-reader/geometry.ts#turnAbout`.
 *  An SVG transform list applies right to left: `scale` (the flip, carried
 *  back to the centre by `translate(2·cx 2·cy)`) runs first, `rotate` last. */
const pictureTransform = (frame: Box, rot: number | undefined, flipH: boolean, flipV: boolean): string => {
  const cx = toPoints(frame.x + frame.cx / 2);
  const cy = toPoints(frame.y + frame.cy / 2);
  const steps: string[] = [];
  if (rot !== undefined && rot !== 0)
    steps.push(`rotate(${formatNumber(rot / 60_000)} ${formatNumber(cx)} ${formatNumber(cy)})`);
  if (flipH || flipV)
    steps.push(
      `translate(${formatNumber(flipH ? 2 * cx : 0)} ${formatNumber(flipV ? 2 * cy : 0)}) scale(${flipH ? -1 : 1} ${flipV ? -1 : 1})`,
    );
  return steps.length > 0 ? ` transform="${steps.join(" ")}"` : "";
};

/** The SVG `transform` of a picture its groups turn, mirror or stretch: the
 *  reader's whole map (`PptxPictureShape.matrix`, EMU) with its move in
 *  points. The coefficients take six decimals, not two: on a 170 pt picture
 *  −0.87 for cos 30° would move a corner by 0.7 pt. */
const pictureMatrix = (matrix: Transform): string => {
  const formatCoefficient = (value: number): string => String(Math.round(value * 1e6) / 1e6 + 0);
  return ` transform="matrix(${formatCoefficient(matrix.a)} ${formatCoefficient(matrix.b)} ${formatCoefficient(matrix.c)} ${formatCoefficient(matrix.d)} ${formatNumber(toPoints(matrix.e))} ${formatNumber(toPoints(matrix.f))})"`;
};

/** The area a picture covers once placed: its frame's four corners through
 *  its groups' matrix, or turned about its centre (a flip moves no corner
 *  out of the frame's own bounds). */
const growByPicture = (drawn: Extent, frame: Box, rot: number | undefined, matrix: Transform | undefined): void => {
  const x = toPoints(frame.x);
  const y = toPoints(frame.y);
  const width = toPoints(frame.cx);
  const height = toPoints(frame.cy);
  if (matrix !== undefined) {
    const placed = { ...matrix, e: toPoints(matrix.e), f: toPoints(matrix.f) };
    for (const [cornerX, cornerY] of [
      [x, y],
      [x + width, y],
      [x, y + height],
      [x + width, y + height],
    ] as const) {
      const at = applyMatrix(placed, cornerX, cornerY);
      grow(drawn, at.x, at.y, at.x, at.y);
    }
    return;
  }
  const angle = ((rot ?? 0) / 60_000) * (Math.PI / 180);
  // |cos| and |sin| rounded to 1e-9: a quarter turn gives exactly 0 and 1.
  const cos = Math.round(Math.abs(Math.cos(angle)) * 1e9) / 1e9;
  const sin = Math.round(Math.abs(Math.sin(angle)) * 1e9) / 1e9;
  const halfWidth = (width * cos + height * sin) / 2;
  const halfHeight = (width * sin + height * cos) / 2;
  grow(
    drawn,
    x + width / 2 - halfWidth,
    y + height / 2 - halfHeight,
    x + width / 2 + halfWidth,
    y + height / 2 + halfHeight,
  );
};

/** A paragraph as lines of text, each the [start, end) of its characters
 *  in `text`: a line break (`a:br`, a run "\n") ends a line; inside a line,
 *  words (runs of non-space) are packed greedily while their estimated
 *  width (`measure` of the characters [start, end), in points) fits
 *  `widthPt`. A word wider than the line stands alone on its line
 *  and is never cut. Spaces between words are kept as written inside a line,
 *  and dropped where a line wraps. `isUnwrapped` (`wrap="none"`): only the line
 *  breaks end a line. */
const wrap = (
  text: string,
  widthPt: number,
  measure: (start: number, end: number) => number,
  isUnwrapped: boolean,
): { start: number; end: number }[] => {
  const lines: { start: number; end: number }[] = [];
  const push = (start: number, rawEnd: number): void => {
    lines.push({ start, end: start + text.slice(start, rawEnd).trimEnd().length });
  };
  let base = 0;
  for (const hardLine of text.split("\n")) {
    const hardEnd = base + hardLine.length;
    if (isUnwrapped) {
      push(base, hardEnd);
      base = hardEnd + 1;
      continue;
    }
    let lineStart = base;
    let position = base;
    let hasWord = false;
    for (const piece of hardLine.match(/\s+|\S+/g) ?? []) {
      if (/^\s/.test(piece)) {
        position += piece.length;
        continue;
      }
      if (hasWord && measure(lineStart, position + piece.length) > widthPt) {
        push(lineStart, position);
        lineStart = position;
      }
      position += piece.length;
      hasWord = true;
    }
    push(lineStart, hardEnd);
    base = hardEnd + 1;
  }
  return lines;
};

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** The area drawn so far, in points; empty (infinite bounds) until
 *  something is drawn. */
interface Extent {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

const grow = (extent: Extent, x0: number, y0: number, x1: number, y1: number): void => {
  extent.x0 = Math.min(extent.x0, x0, x1);
  extent.y0 = Math.min(extent.y0, y0, y1);
  extent.x1 = Math.max(extent.x1, x0, x1);
  extent.y1 = Math.max(extent.y1, y0, y1);
};

const inside = (inner: Rect, outer: Rect): boolean =>
  holdsRect(RECT_RULES["text-in-shape"], rawFromXywh(outer), rawFromXywh(inner));

/** The colour attribute of a piece of text: none for black by default. */
const fillAttr = (color: Hex | undefined): string => (color === undefined ? "" : ` fill="${color}"`);

/** The corner radius of a round rectangle (points), from its `adj` (100 000ths of the shorter side, clamped to [0, 50 000]). */
const cornerRadius = (shape: PptxTextShape, outline: Rect): number =>
  (Math.min(outline.w, outline.h) * Math.min(Math.max(shape.adj ?? DEFAULT_ROUND_ADJ, 0), 50_000)) / 100_000;

/** A point turned by `degrees` degrees clockwise (y down, as SVG `rotate`) about (cx, cy). The cosine and sine are rounded to 1e-9 so that a quarter
 *  turn is exact (`Math.cos(Math.PI / 2)` is 6e-17), as `growByPicture` does. */
const turnedPoint = (x: number, y: number, cx: number, cy: number, degrees: number): [number, number] => {
  const exact = (value: number): number => Math.round(value * 1e9) / 1e9;
  const [cos, sin] = [exact(Math.cos((degrees * Math.PI) / 180)), exact(Math.sin((degrees * Math.PI) / 180))];
  const [dx, dy] = [x - cx, y - cy];
  return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos];
};

/** The `transform` attribute (with its leading space) turning a SmartArt shape's outline about its frame's centre; none for an angle of 0. */
const outlineTurn = (art: NonNullable<PptxTextShape["smartArt"]>): string =>
  art.turn === 0
    ? ""
    : ` transform="rotate(${formatNumber(art.turn / 60_000)} ${formatNumber(toPoints(art.frame.x + art.frame.cx / 2))} ${formatNumber(toPoints(art.frame.y + art.frame.cy / 2))})"`;

/** The six vertices (points, on the slide) of a `hexagon` preset, in its frame as written (the shape's turn comes as `transform`). Formulas of
 *  presetShapeDefinitions: `x1 = ss·a/100000` (`a` pinned to [0, 50000·w/ss]), `dy1 = (h/2)·vf/100000·sin 60°`, vertices at the left and right
 *  middles and at (x1, y1), (r − x1, y1), (r − x1, y2), (x1, y2). Their `ss`, `a`, `w`, `h`, `l`, `r` and `vc` are `shortSide`, `adjust`,
 *  `width`, `height`, `left`, `right` and `middleY` here. */
const hexagonPoints = (art: NonNullable<PptxTextShape["smartArt"]>, adj: number | undefined): [number, number][] => {
  const [left, top, width, height] = [
    toPoints(art.frame.x),
    toPoints(art.frame.y),
    toPoints(art.frame.cx),
    toPoints(art.frame.cy),
  ];
  const shortSide = Math.min(width, height);
  const adjust = shortSide === 0 ? 0 : Math.min(Math.max(adj ?? HEXAGON_ADJ, 0), (50_000 * width) / shortSide);
  const x1 = (shortSide * adjust) / 100_000;
  const dy1 = (((height / 2) * (art.vf ?? HEXAGON_VF)) / 100_000) * Math.sin(Math.PI / 3);
  const [middleY, right] = [top + height / 2, left + width];
  return [
    [left, middleY],
    [left + x1, middleY - dy1],
    [right - x1, middleY - dy1],
    [right, middleY],
    [right - x1, middleY + dy1],
    [left + x1, middleY + dy1],
  ];
};

/** A `chord` preset (ECMA-376 §20.1.9.18, presetShapeDefinitions): the ellipse of its frame cut by the straight line from the point at
 *  angle `adj1` to the point at angle `adj2` (60 000ths of a degree, clockwise from the right, y down), the arc between them going
 *  clockwise from `adj1`. Defaults `adj1` 2 700 000 (45 degrees), `adj2` 16 200 000 (270). A point at angle a is that of the ellipse's
 *  parameter atan2(w sin a, h cos a), the preset's own `at2` guide. The frame is as written (the shape's turn comes as `transform`).
 *  `undefined` when the arc is empty or a whole turn (the ellipse is then drawn). */
const chordPath = (art: NonNullable<PptxTextShape["smartArt"]>): string | undefined => {
  const [left, top, width, height] = [
    toPoints(art.frame.x),
    toPoints(art.frame.y),
    toPoints(art.frame.cx),
    toPoints(art.frame.cy),
  ];
  const [cx, cy] = [left + width / 2, top + height / 2];
  const pin = (value: number): number => Math.min(Math.max(value, 0), FULL_TURN - 1);
  const start = pin(art.adj1 ?? 2_700_000);
  const swing = pin(art.adj2 ?? 16_200_000) - start;
  const sweep = swing > 0 ? swing : swing + FULL_TURN;
  if (sweep >= FULL_TURN || width === 0 || height === 0) return undefined;
  const pointAt = (angle: number): [number, number] => {
    const radians = (angle / 60_000) * (Math.PI / 180);
    const param = Math.atan2((width / 2) * Math.sin(radians), (height / 2) * Math.cos(radians));
    return [cx + (width / 2) * Math.cos(param), cy + (height / 2) * Math.sin(param)];
  };
  const [x1, y1] = pointAt(start);
  const [x2, y2] = pointAt(start + sweep);
  return `M${formatNumber(x1)},${formatNumber(y1)}A${formatNumber(width / 2)} ${formatNumber(height / 2)} 0 ${sweep > FULL_TURN / 2 ? 1 : 0} 1 ${formatNumber(x2)},${formatNumber(y2)}Z`;
};

/** The element of a SmartArt shape's `hexagon` or `chord` preset; `undefined` for any other (and for a chord that is a whole turn). */
const presetElement = (
  shape: PptxTextShape,
  art: NonNullable<PptxTextShape["smartArt"]>,
  paint: string,
): string | undefined => {
  if (shape.geometry === "hexagon") {
    const points = hexagonPoints(art, shape.adj).map(([x, y]) => `${formatNumber(x)},${formatNumber(y)}`);
    return `<polygon points="${points.join(" ")}" ${paint}/>`;
  }
  const pathData = shape.geometry === "chord" ? chordPath(art) : undefined;
  return pathData === undefined ? undefined : `<path d="${pathData}" ${paint}/>`;
};

/** The element of a shape's outline or fill: the traced preset of a SmartArt shape, else an ellipse, else a rectangle (rounded for `roundRect`). */
const outlineElement = (shape: PptxTextShape, outline: Rect, radius: number, paint: string): string => {
  const traced = shape.smartArt === undefined ? undefined : presetElement(shape, shape.smartArt, paint);
  if (traced !== undefined) return traced;
  if (shape.geometry === "ellipse" || shape.geometry === "chord") {
    return `<ellipse cx="${formatNumber(outline.x + outline.w / 2)}" cy="${formatNumber(outline.y + outline.h / 2)}" rx="${formatNumber(outline.w / 2)}" ry="${formatNumber(outline.h / 2)}" ${paint}/>`;
  }
  const corner = shape.geometry === "roundRect" ? ` rx="${formatNumber(radius)}"` : "";
  return `<rect x="${formatNumber(outline.x)}" y="${formatNumber(outline.y)}" width="${formatNumber(outline.w)}" height="${formatNumber(outline.h)}"${corner} ${paint}/>`;
};

/** The paint attributes of an outline: a black stroke alone when no fill is drawn, else the fill, with the stroke when the line shows. */
const paintOf = (fill: Hex | undefined, stroke: boolean): string => {
  if (fill === undefined) return STROKE;
  return `fill="${fill}"${stroke ? ' stroke="#000" stroke-width="1"' : ""}`;
};

/** A shape's outline or fill, when it shows one, pushed onto `elements` and grown into `drawn`. Says which fill was drawn (`undefined`: none) and
 *  whether the shape holds visible text. */
const drawOutline = (
  shape: PptxTextShape,
  shown: { outline: Rect; bounds: Rect },
  radius: number,
  elements: string[],
  drawn: Extent,
): { fill: Hex | undefined; hasText: boolean } => {
  const { outline, bounds } = shown;
  // The fill is drawn only when it is exact and the colour of every visible
  // run is known too: a known fill under text of unknown colour could hide
  // it (white on white). Otherwise: an outline for a fill not drawn, black
  // text.
  const visibleRuns = shape.paragraphs.flatMap((paragraph) => paragraph.runs.filter((run) => run.text.trim() !== ""));
  const hasText = visibleRuns.length > 0;
  const fill =
    shape.fill !== "none" && shape.fill !== "unknown" && visibleRuns.every((run) => run.color !== undefined)
      ? shape.fill
      : undefined;
  const stroke = shape.hasVisibleLine || (fill === undefined && shape.fill !== "none");
  if (fill !== undefined || stroke) {
    const turn = shape.smartArt === undefined ? "" : outlineTurn(shape.smartArt);
    elements.push(outlineElement(shape, outline, radius, paintOf(fill, stroke) + turn));
    grow(drawn, bounds.x, bounds.y, bounds.x + bounds.w, bounds.y + bounds.h);
  }
  return { fill, hasText };
};

/** A turn of `degrees` degrees about (`cx`, `cy`), in points. */
interface Turn {
  degrees: number;
  cx: number;
  cy: number;
}

/** What a SmartArt shape's text adds to its lines: `sink` receives each line drawn; `turns` are the rotations the text is drawn under,
 *  in the order SVG applies them to a point: the text rectangle's own turn first, the shape's turn second. */
interface SmartArtText {
  sink: string[];
  turns: Turn[];
  page: Rect;
}

/** The text rectangle `textRect` and the `room` around which a line counts as out (`textRectangle`); `art`: a SmartArt shape's extras. */
interface TextRectangles {
  textRect: Rect;
  room: Rect;
  art?: SmartArtText;
}

/** The preset's own inset of the text rectangle on each side (ellipse, round rectangle). None for a SmartArt shape: its `dsp:txXfrm` IS
 *  the text rectangle, the preset's inset already taken. */
const presetInset = (shape: PptxTextShape, frame: Rect, radius: number): { x: number; y: number } => {
  if (shape.smartArt !== undefined) return { x: 0, y: 0 };
  if (shape.geometry === "ellipse") return { x: frame.w * ELLIPSE_TEXT_INSET, y: frame.h * ELLIPSE_TEXT_INSET };
  if (shape.geometry === "roundRect") return { x: radius * ROUND_TEXT_INSET, y: radius * ROUND_TEXT_INSET };
  return { x: 0, y: 0 };
};

/** The text rectangle `textRect` (the preset's own, less the written insets; a side not written: the default of `a:bodyPr`) and the `room` the text may
 *  fill before it counts as overflowing. */
const textRectangle = (shape: PptxTextShape, frame: Rect, radius: number): TextRectangles => {
  const { x: presetInsetX, y: presetInsetY } = presetInset(shape, frame, radius);
  const inset = {
    left: toPoints(shape.insets?.left ?? INSET_X_EMU),
    top: toPoints(shape.insets?.top ?? INSET_Y_EMU),
    right: toPoints(shape.insets?.right ?? INSET_X_EMU),
    bottom: toPoints(shape.insets?.bottom ?? INSET_Y_EMU),
  };
  const textRect: Rect = {
    x: frame.x + presetInsetX + inset.left,
    y: frame.y + presetInsetY + inset.top,
    w: frame.w - 2 * presetInsetX - inset.left - inset.right,
    h: frame.h - 2 * presetInsetY - inset.top - inset.bottom,
  };
  // What the text may fill before it is counted as overflowing: the
  // preset's text rectangle WITH its insets. An inset is padding: the text
  // is laid out inside it (wrapped at `textRect.w`, anchored in `textRect`) but may eat
  // into it and still sit inside its shape. A line counts as out when its
  // MIDDLE leaves that rectangle: more than half of it outside.
  // Starting value to calibrate (half a line), from PowerPoint's own export
  // of the witnesses (27/09/2026). Fits in PowerPoint: IT 3 « OTA… », three
  // 7 pt lines (25.2 pt) centred in 19.5 pt, a third of a line out at each
  // end; IT 2-6's legend labels, 9 pt at 7.2 pt below the top of a 17.2 pt
  // box, 0.8 pt (7 % of the line) out. Overflows in PowerPoint: Making-the-
  // mobile-web 18, 32 lines of 14 pt (538 pt) centred in 490 pt, the first
  // and last lines 1.4 lines out. Between the two populations: 0.34 and 1.4
  // of a line.
  // Bias, chosen: a text that spills into its padding, or past its box by
  // less than half a line, is not counted.
  const room: Rect = {
    x: frame.x + presetInsetX,
    y: frame.y + presetInsetY,
    w: frame.w - 2 * presetInsetX,
    h: frame.h - 2 * presetInsetY,
  };
  return { textRect, room };
};

/** The paragraphs of a text shape with each run's size resolved, and the size a level falls back to. */
interface SizedParagraphs {
  sized: { paragraph: PptxTextShape["paragraphs"][number]; sizes: { text: string; size: number | undefined }[] }[];
  fallback: (level: number) => number | undefined;
}

/** Each run's size: the one it writes; else, for a shape that is no
 *  placeholder, the one `p:defaultTextStyle` declares for the paragraph's
 *  level, or for level 1 when that level declares none; for a placeholder,
 *  the one it inherits at that level from its layout and master
 *  (`PptxTextShape.inheritedSizes`). A visible run
 *  left without a size leaves the whole text out (`undefined`): its size would be a guess. */
const sizeRuns = (
  shape: PptxTextShape,
  defaultSizes: Readonly<Record<number, number>>,
): SizedParagraphs | undefined => {
  const fallback = (level: number): number | undefined =>
    shape.placeholder === undefined ? (defaultSizes[level] ?? defaultSizes[0]) : shape.inheritedSizes?.[level];
  // `endSize` between the run's own size and the level's: the paragraph's end
  // mark is the one written size of a box whose runs carry none; without it
  // that box's whole text would be omitted (`size-not-declared`).
  const sized = shape.paragraphs.map((paragraph) => ({
    paragraph,
    sizes: paragraph.runs.map((run) => ({
      text: run.text,
      size: run.size ?? paragraph.endSize ?? fallback(paragraph.level ?? 0),
    })),
  }));
  if (sized.some(({ sizes }) => sizes.some((run) => run.size === undefined && run.text.trim() !== "")))
    return undefined;
  return { sized, fallback };
};

/** A line is written as pieces of one colour and size each (colour `undefined`: black). */
type Piece = { text: string; color: Hex | undefined; size: number };
interface TextLine {
  text: string;
  pieces: Piece[];
  size: number;
  height: number;
  width: number;
  align: "ctr" | "r" | undefined;
}

/** What lays out the lines of every paragraph of one shape alike. */
interface LayoutEnv {
  /** The width a line wraps at, in points. */
  wrapWidth: number;
  /** The width of the room a line may fill before it counts as overflowing. */
  roomWidth: number;
  isUnwrapped: boolean;
  /** The text keeps its colours; else it is black. */
  coloured: boolean;
  /** `a:normAutofit fontScale` (1: none). */
  scale: number;
  /** `lnSpcReduction`, taken off the line spacing's percentage. */
  reduction: number;
  /** Factor on every estimated advance (1 for a slide's text; `SMARTART_ADVANCE` for a SmartArt's). */
  advance: number;
}

/** The largest size among runs that have one; a loop, never `Math.max(...sizes)`: a paragraph can hold more runs
 *  than a call takes arguments (a RangeError past ~124k of them). */
const largestKnownSize = (sizes: readonly { size: number | undefined }[]): number | undefined => {
  let largest: number | undefined;
  for (const run of sizes) {
    if (run.size !== undefined && (largest === undefined || run.size > largest)) largest = run.size;
  }
  return largest;
};

/** The pieces of one line, `[start, end)` of the paragraph's text, and the largest size among its visible characters (0: none). */
const piecesOf = (
  joined: string,
  start: number,
  end: number,
  colorAt: readonly (Hex | undefined)[],
  sizeAt: readonly number[],
): { pieces: Piece[]; size: number } => {
  const pieces: Piece[] = [];
  let size = 0;
  for (let index = start; index < end; index++) {
    if (joined[index]!.trim() !== "") size = Math.max(size, sizeAt[index]!);
    const last = pieces.at(-1);
    if (last !== undefined && last.color === colorAt[index] && last.size === sizeAt[index]) last.text += joined[index];
    else pieces.push({ text: joined[index]!, color: colorAt[index], size: sizeAt[index]! });
  }
  return { pieces, size };
};

/** The lines of one paragraph, pushed onto `lines`; says whether one is wider than the room. `paragraphSize`: the height of a line with no visible text. */
const layoutParagraph = (
  paragraph: PptxTextShape["paragraphs"][number],
  sizes: readonly { text: string; size: number | undefined }[],
  paragraphSize: number,
  env: LayoutEnv,
  lines: TextLine[],
): boolean => {
  const { wrapWidth, roomWidth, isUnwrapped, coloured, scale, reduction } = env;
  const spacing = paragraph.lineSpacing;
  const joined = paragraph.runs.map((run) => run.text).join("");
  // Each character's colour and size, from the run it belongs to.
  const colorAt = paragraph.runs.flatMap((run) =>
    Array.from({ length: run.text.length }, () => (coloured ? run.color : undefined)),
  );
  const sizeAt = sizes.flatMap((run) =>
    Array.from({ length: run.text.length }, () => (run.size === undefined ? paragraphSize : run.size * scale)),
  );
  const measure = (start: number, end: number): number => {
    let width = 0;
    for (let index = start; index < end; index++) width += advanceEm(joined[index]!) * sizeAt[index]! * env.advance;
    return width;
  };
  let overflow = false;
  for (const { start, end } of wrap(joined, wrapWidth, measure, isUnwrapped)) {
    const { pieces, size: largest } = piecesOf(joined, start, end, colorAt, sizeAt);
    const size = largest === 0 ? paragraphSize : largest;
    const height =
      spacing !== undefined && "points" in spacing
        ? spacing.points
        : size * LINE_HEIGHT_EM * Math.max((spacing?.percent ?? 1) - reduction, 0);
    const width = measure(start, end);
    lines.push({ text: joined.slice(start, end), pieces, size, height, width, align: paragraph.align });
    if (width > roomWidth) overflow = true;
  }
  return overflow;
};

/** The lines of a text shape, wrapped at the text rectangle's width, and whether one is wider than the room. The text keeps its colours
 *  (`coloured`) or is black. */
const layoutLines = (
  shape: PptxTextShape,
  { sized, fallback }: SizedParagraphs,
  textRect: Rect,
  room: Rect,
  coloured: boolean,
): { lines: TextLine[]; overflow: boolean } => {
  // Each character at its run's size: a line is as high as the largest text
  // ON IT, not in its paragraph (at the paragraph's largest size throughout,
  // a line mixing 17 pt and 13 pt wraps too early and stands too high).
  // A line without visible text (empty paragraph, line breaks only)
  // is as high as its paragraph's largest size, else its level's default,
  // else the shape's first sized run: only a height, no text is drawn at it.
  // `a:normAutofit fontScale` shrinks every size; `lnSpcReduction` is taken
  // off the line spacing's percentage. A spacing in points is the distance
  // between baselines as written.
  const scale = shape.fontScale ?? 1;
  const env: LayoutEnv = {
    wrapWidth: Math.min(textRect.w, MAX_WRAP_WIDTH_PT),
    roomWidth: room.w,
    isUnwrapped: shape.isUnwrapped === true,
    coloured,
    scale,
    reduction: shape.lineSpaceReduction ?? 0,
    advance: shape.smartArt === undefined ? 1 : SMARTART_ADVANCE,
  };
  const shapeSize = sized.flatMap(({ sizes }) => sizes).find((run) => run.size !== undefined)!.size!;
  let overflow = false;
  const lines: TextLine[] = [];
  for (const { paragraph, sizes } of sized) {
    const paragraphSize = (largestKnownSize(sizes) ?? fallback(paragraph.level ?? 0) ?? shapeSize) * scale;
    if (layoutParagraph(paragraph, sizes, paragraphSize, env, lines)) overflow = true;
  }
  return { lines, overflow };
};

/** The four corners of a box turned by `turns` (each about its own centre, the first applied first). */
const turnedCorners = (box: Rect, turns: readonly Turn[]): [number, number][] =>
  (
    [
      [box.x, box.y],
      [box.x + box.w, box.y],
      [box.x, box.y + box.h],
      [box.x + box.w, box.y + box.h],
    ] as const
  ).map(([x, y]) =>
    turns.reduce<[number, number]>(([px, py], turn) => turnedPoint(px, py, turn.cx, turn.cy, turn.degrees), [x, y]),
  );

/** `grow`, for a box turned by `turns`. */
const growTurned = (drawn: Extent, box: Rect, turns: readonly Turn[]): void => {
  for (const [tx, ty] of turnedCorners(box, turns)) grow(drawn, tx, ty, tx, ty);
};

/** Whether the bounds of a box turned by `turns` meet the page. */
const crossesPage = (box: Rect, turns: readonly Turn[], page: Rect): boolean => {
  const corners = turnedCorners(box, turns);
  const xs = corners.map((corner) => corner[0]);
  const ys = corners.map((corner) => corner[1]);
  return (
    Math.max(...xs) > page.x &&
    Math.min(...xs) < page.x + page.w &&
    Math.max(...ys) > page.y &&
    Math.min(...ys) < page.y + page.h
  );
};

/** The `transform` attribute of a text turned by `turns` (none: nothing). SVG applies the list right to left, so the text rectangle's
 *  turn, first in `turns`, is written last. */
const transformAttr = (turns: readonly Turn[]): string => {
  if (turns.length === 0) return "";
  const rotations = [...turns]
    .reverse()
    .map((turn) => `rotate(${formatNumber(turn.degrees)} ${formatNumber(turn.cx)} ${formatNumber(turn.cy)})`);
  return ` transform="${rotations.join(" ")}"`;
};

/** Where a line is written, by its alignment: the `x` of its `<text>`, its `text-anchor` attribute (with its leading space; none for a
 *  line aligned left), and the left edge of the area it covers. */
const linePlacement = (line: TextLine, textRect: Rect): { x: number; anchor: string; left: number } => {
  if (line.align === "ctr") {
    const x = textRect.x + textRect.w / 2;
    return { x, anchor: ' text-anchor="middle"', left: x - line.width / 2 };
  }
  if (line.align === "r") {
    const x = textRect.x + textRect.w;
    return { x, anchor: ' text-anchor="end"', left: x - line.width };
  }
  return { x: textRect.x, anchor: "", left: textRect.x };
};

/** One line as its `<text>` element at `baseline`, and the area it covers (before the turns, which `growTurned` applies). `turns`, as SVG
 *  writes them, are applied right to left: the text rectangle's first. */
const lineElement = (
  line: TextLine,
  textRect: Rect,
  baseline: number,
  turns: readonly Turn[],
): { element: string; covers: Rect } => {
  const { x, anchor, left } = linePlacement(line, textRect);
  const turned = transformAttr(turns);
  const onlyPiece = line.pieces.length === 1 ? line.pieces[0]! : undefined;
  const sizeAttr = (piece: Piece): string =>
    piece.size === line.size ? "" : ` font-size="${formatNumber(piece.size)}"`;
  const body =
    onlyPiece !== undefined
      ? escapeAttr(line.text)
      : line.pieces
          .map((piece) => `<tspan${fillAttr(piece.color)}${sizeAttr(piece)}>${escapeAttr(piece.text)}</tspan>`)
          .join("");
  return {
    element: `<text x="${formatNumber(x)}" y="${formatNumber(baseline)}" font-family="sans-serif" font-size="${formatNumber(line.size)}"${anchor}${fillAttr(onlyPiece?.color)}${turned} xml:space="preserve">${body}</text>`,
    covers: { x: left, y: baseline - line.size, w: line.width, h: line.size * (1 + DESCENT_EM) },
  };
};

/** The top of a stack of lines `height` high in its text rectangle, by the shape's anchor: at the bottom (`b`), centred (`ctr`), else
 *  at the top. */
const stackTop = (anchor: PptxTextShape["anchor"], textRect: Rect, height: number): number => {
  if (anchor === "b") return textRect.y + textRect.h - height;
  if (anchor === "ctr") return textRect.y + (textRect.h - height) / 2;
  return textRect.y;
};

/** The lines as `<text>` elements, stacked from the anchor, grown into `drawn`; says whether a line's middle leaves the room. */
const emitLines = (
  shape: PptxTextShape,
  lines: readonly TextLine[],
  { textRect, room, art }: TextRectangles,
  elements: string[],
  drawn: Extent,
): boolean => {
  let overflow = false;
  const height = lines.reduce((sum, line) => sum + line.height, 0);
  let y = stackTop(shape.anchor, textRect, height);
  for (const line of lines) {
    // `a:spAutoFit`: the written box was grown to hold the text; a height
    // past it is our estimate's error, not the deck's.
    const middle = y + line.height / 2;
    if (shape.hasSpAutoFit !== true && (middle < room.y || middle > room.y + room.h)) overflow = true;
    // The baseline sits at the size below the line's top, the rest of the
    // line height below it (descenders).
    const baseline = y + line.size;
    y += line.height;
    if (line.text.trim() === "") continue;
    const { element, covers } = lineElement(line, textRect, baseline, art?.turns ?? []);
    // A SmartArt line turned wholly out of the page is cropped away: not drawn, its words stay undrawn.
    if (art !== undefined && !crossesPage(covers, art.turns, art.page)) continue;
    growTurned(drawn, covers, art?.turns ?? []);
    art?.sink.push(line.text);
    elements.push(element);
  }
  return overflow;
};

/** What drawing the shapes of one slide shares: the deck's default sizes, the boxes drawn filled so far (paint order), what is
 *  omitted, the area drawn so far, and the text lines drawn for SmartArt shapes. */
interface DrawEnv {
  /** The page the drawing is cropped to (points, origin 0): a SmartArt line wholly outside it is not drawn. */
  page: Rect;
  defaultSizes: Readonly<Record<number, number>>;
  filled: Rect[];
  omit: (reason: SlideSvgOmission) => void;
  drawn: Extent;
  smartArtText: string[];
}

/** Whether a text shape's geometry is one the drawer traces: `rect`, `roundRect`, `ellipse`, and for a SmartArt's shape the `hexagon` and the `chord` too. */
const tracesGeometry = (shape: PptxTextShape): boolean => {
  const geometry = shape.geometry ?? "";
  return DRAWN_GEOMETRIES.has(geometry) || (shape.smartArt !== undefined && SMARTART_GEOMETRIES.has(geometry));
};

/** A text shape's outline or fill (when it shows one) and its text, or why
 *  not. `env.filled`: the boxes drawn filled so far, in paint order; a box this
 *  shape fills is added to it. `env.drawn` grows by what it draws: its box when
 *  outlined or filled, each line of text as estimated (past the box too). */
const drawText = (shape: PptxTextShape, env: DrawEnv, elements: string[]): { overflow: boolean } => {
  const { omit } = env;
  const box = shape.box!;
  if (!tracesGeometry(shape)) {
    omit("geometry");
    return { overflow: false };
  }
  if (shape.rot !== undefined || shape.isTurnedByGroup) {
    omit("turned");
    return { overflow: false };
  }
  // `outline`: what the outline is drawn in (a SmartArt shape: its frame as written, under `rotate`); `bounds`: the box it covers once turned.
  const bounds: Rect = { x: toPoints(box.x), y: toPoints(box.y), w: toPoints(box.cx), h: toPoints(box.cy) };
  const frame = shape.smartArt?.frame;
  const outline: Rect =
    frame === undefined
      ? bounds
      : { x: toPoints(frame.x), y: toPoints(frame.y), w: toPoints(frame.cx), h: toPoints(frame.cy) };
  const radius = cornerRadius(shape, outline);
  const { fill, hasText } = drawOutline(shape, { outline, bounds }, radius, elements, env.drawn);
  // `coloured` below reads `filled` only when this shape has no fill of its
  // own, i.e. when nothing is pushed here: no copy of the list is needed.
  // The bounds of a shape turned by other than a quarter are no background for a text laid flat (`inside` compares unturned rectangles).
  if (fill !== undefined && (shape.smartArt === undefined || shape.smartArt.turn % QUARTER_TURN === 0))
    env.filled.push(bounds);
  if (!hasText) return { overflow: false };
  if (shape.isTextTurned || shape.flipV) {
    omit("turned-text");
    return { overflow: false };
  }
  return drawLines(shape, env, elements, { outline, radius, fill });
};

/** The rectangle a SmartArt shape's text sits in (`dsp:txXfrm`), in points; the shape's own box otherwise. */
const textFrameOf = (shape: PptxTextShape, outline: Rect): Rect => {
  const art = shape.smartArt;
  if (art === undefined) return outline;
  const written: Rect = {
    x: toPoints(art.textBox.x),
    y: toPoints(art.textBox.y),
    w: toPoints(art.textBox.cx),
    h: toPoints(art.textBox.cy),
  };
  // `a:bodyPr vert` of 90 or 270: the lines run along the written height, so the sides are swapped about the box's centre.
  if (art.textSwap === true)
    return {
      x: written.x + written.w / 2 - written.h / 2,
      y: written.y + written.h / 2 - written.w / 2,
      w: written.h,
      h: written.w,
    };
  return written;
};

/** A SmartArt shape's text is laid out flat in its rectangle, turned by `textTurn` about THAT rectangle's centre (insets included: they can
 *  differ side to side), then by the shape's turn about the FRAME's centre; one mechanism for every angle, a quarter turn is no special case. Two turns that cancel
 *  about one centre (a hexagon turned 90 and its text rectangle -90) write no transform. `undefined` for any other text. */
const smartArtText = (shape: PptxTextShape, env: DrawEnv): SmartArtText | undefined => {
  const art = shape.smartArt;
  if (art === undefined) return undefined;
  const inner: Turn = {
    degrees: art.textTurn / 60_000,
    cx: toPoints(art.textBox.x + art.textBox.cx / 2),
    cy: toPoints(art.textBox.y + art.textBox.cy / 2),
  };
  const outer: Turn = {
    degrees: art.turn / 60_000,
    cx: toPoints(art.frame.x + art.frame.cx / 2),
    cy: toPoints(art.frame.y + art.frame.cy / 2),
  };
  const sameCentre = Math.abs(inner.cx - outer.cx) < SAME_CENTRE_PT && Math.abs(inner.cy - outer.cy) < SAME_CENTRE_PT;
  const cancel = (art.turn + art.textTurn) % FULL_TURN === 0 && sameCentre;
  return {
    sink: env.smartArtText,
    page: env.page,
    turns: cancel ? [] : [inner, outer].filter((turn) => turn.degrees !== 0),
  };
};

/** The text of a shape whose outline is drawn: sized, laid out in its text rectangle and emitted, or left out when a size is not declared. */
const drawLines = (
  shape: PptxTextShape,
  env: DrawEnv,
  elements: string[],
  placed: { outline: Rect; radius: number; fill: Hex | undefined },
): { overflow: boolean } => {
  const { outline, radius, fill } = placed;
  const art = smartArtText(shape, env);
  const base = textRectangle(shape, textFrameOf(shape, outline), radius);
  const rects: TextRectangles = art === undefined ? base : { ...base, art };
  const { textRect, room } = rects;
  // The text keeps its colours on its own drawn fill, or, without a fill of
  // its own, when its text rectangle lies inside a box drawn filled BEFORE
  // it; else it is black. The text rectangle, not the box: a label's box may
  // run a few EMU past the bar it sits on. A fill of unknown
  // colour (not drawn) keeps the text black: what shows under it is not known.
  const coloured =
    fill !== undefined || (shape.fill === "none" && env.filled.some((filledBox) => inside(textRect, filledBox)));
  const runs = sizeRuns(shape, env.defaultSizes);
  if (runs === undefined) {
    env.omit("size-not-declared");
    return { overflow: false };
  }
  const laid = layoutLines(shape, runs, textRect, room, coloured);
  const outside = emitLines(shape, laid.lines, rects, elements, env.drawn);
  return { overflow: laid.overflow || outside };
};

/** A straight connector's line, or a bent one's polyline through the
 *  vertices the reader computed (`pptx-reader/geometry.ts#computeBentPath`), with
 *  their arrow heads; or why not. An invisible line (`hasVisibleLine` false)
 *  draws nothing and is no omission. */
const drawConnector = (
  shape: PptxConnectorShape,
  elements: string[],
  omit: (reason: SlideSvgOmission) => void,
  drawn: Extent,
): boolean => {
  const fate = connectorFate(shape);
  if (fate === "not-traceable") omit("connector-geometry");
  if (fate !== "drawn") return false;
  // Every declared head is drawn as a triangle: "triangle", "arrow" and
  // "stealth" all point the same way; the reader keeps no finer kind.
  const start = shape.startArrow ? ' marker-start="url(#s)"' : "";
  const end = shape.endArrow ? ' marker-end="url(#e)"' : "";
  const ends = STRAIGHT_CONNECTORS.has(shape.geometry ?? "") ? shape.ends : undefined;
  if (ends !== undefined) {
    elements.push(
      `<line x1="${formatNumber(toPoints(ends.x1))}" y1="${formatNumber(toPoints(ends.y1))}" x2="${formatNumber(toPoints(ends.x2))}" y2="${formatNumber(toPoints(ends.y2))}" stroke="#000" stroke-width="1"${start}${end}/>`,
    );
    grow(drawn, toPoints(ends.x1), toPoints(ends.y1), toPoints(ends.x2), toPoints(ends.y2));
  } else {
    const path = shape.path ?? [];
    for (const vertex of path)
      grow(drawn, toPoints(vertex.x), toPoints(vertex.y), toPoints(vertex.x), toPoints(vertex.y));
    // A vertex written twice in a row (an adj of 0 or 100 000 folds a
    // segment to nothing) is dropped: a head orients on its segment, and a
    // segment of no length has no direction.
    const points = path
      .map((vertex) => `${formatNumber(toPoints(vertex.x))},${formatNumber(toPoints(vertex.y))}`)
      .filter((point, index, all) => index === 0 || point !== all[index - 1]);
    elements.push(`<polyline points="${points.join(" ")}" fill="none" stroke="#000" stroke-width="1"${start}${end}/>`);
  }
  return shape.startArrow || shape.endArrow;
};

/** The two arrow heads, 6 × 6 stroke widths, tip on the line's end. */
const MARKERS =
  '<defs><marker id="s" viewBox="0 0 10 10" refX="0" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M10 0L0 5L10 10z"/></marker>' +
  '<marker id="e" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0 0L10 5L0 10z"/></marker></defs>';

/** A picture offered to the drawing: the slot of its shape (paint order), its `<image>` tag up to the base64 body, and its bytes. */
interface OfferedPicture {
  slot: number;
  shape: PptxPictureShape;
  head: string;
  bytes: Uint8Array;
}

/** What placing the pictures of one slide shares: what is omitted, the pictures left out, the ones offered so far, and the area drawn. */
interface PictureEnv {
  omit: (reason: SlideSvgOmission) => void;
  picturesNotDrawn: SlideSvg["picturesNotDrawn"];
  offered: OfferedPicture[];
  drawn: Extent;
}

/** A picture left out: counted under `omission`, and listed with the `reason` the caller reports. */
const leaveOutPicture = (
  shape: PptxPictureShape,
  omission: SlideSvgOmission,
  reason: PictureNotDrawn,
  env: PictureEnv,
): void => {
  env.omit(omission);
  env.picturesNotDrawn.push({ spTreeIndex: shape.spTreeIndex, reason });
};

/** A picture with a box offered at `slot`, its area grown into `env.drawn`; or why it is left out. Its base64 body is not encoded
 *  here: `embedPictures` encodes the ones the size cap keeps. */
const offerPicture = (shape: PptxPictureShape, slot: number, env: PictureEnv): void => {
  if (!("image" in shape)) {
    leaveOutPicture(shape, "picture-not-published", shape.notPublished, env);
    return;
  }
  const mimeType = drawableType(shape.image.bytes);
  if (mimeType === undefined) {
    leaveOutPicture(shape, "picture-not-published", "format", env);
    return;
  }
  if (shape.groupStretched) {
    leaveOutPicture(shape, "picture-group-stretched", "group-stretched", env);
    return;
  }
  // The reader gives a frame whenever it gives a box, but for a
  // picture turned inside a group stretched unevenly.
  const frame = shape.frame;
  if (frame === undefined) {
    leaveOutPicture(shape, "no-box", "no-box", env);
    return;
  }
  env.offered.push({
    slot,
    shape,
    head:
      `<image x="${formatNumber(toPoints(frame.x))}" y="${formatNumber(toPoints(frame.y))}" width="${formatNumber(toPoints(frame.cx))}" height="${formatNumber(toPoints(frame.cy))}"` +
      (shape.matrix === undefined
        ? pictureTransform(frame, shape.rot, shape.flipH === true, shape.flipV === true)
        : pictureMatrix(shape.matrix)) +
      ` preserveAspectRatio="none" href="data:${mimeType};base64,`,
    bytes: shape.image.bytes,
  });
  // Every picture offered counts, before the size cap decides: the
  // viewBox is part of that size. One the cap then drops leaves blank
  // space where it stood; the caller is told why (`picturesNotDrawn`).
  growByPicture(env.drawn, frame, shape.rot, shape.matrix);
};

/** The viewBox, in points: the drawn area plus `CROP_MARGIN_PT`, never past the slide (a slide shows nothing beyond its edges); the
 *  whole slide when nothing is drawn on it. */
const viewBoxOf = (drawn: Extent, slideWidth: number, slideHeight: number): Rect => {
  const cropLeft = Math.max(0, drawn.x0 - CROP_MARGIN_PT);
  const cropTop = Math.max(0, drawn.y0 - CROP_MARGIN_PT);
  const cropRight = Math.min(slideWidth, drawn.x1 + CROP_MARGIN_PT);
  const cropBottom = Math.min(slideHeight, drawn.y1 + CROP_MARGIN_PT);
  // Empty: nothing drawn (infinite bounds), or all of it off the slide.
  const empty = !(cropRight > cropLeft && cropBottom > cropTop);
  return empty
    ? { x: 0, y: 0, w: slideWidth, h: slideHeight }
    : { x: cropLeft, y: cropTop, w: cropRight - cropLeft, h: cropBottom - cropTop };
};

/** CSS pixels per point: the whole slide's scale (at most `MAX_WIDTH_PX` for its width, never enlarged), times `MAX_ENLARGEMENT` at
 *  most, and never past `MAX_WIDTH_PX` wide. A drawing that holds a SmartArt shape is shown at the slide's scale, not enlarged:
 *  enlarged twice, its text stood at twice the size it was saved at, against the page's other text. */
const displayScale = (shapes: readonly PptxShape[], viewWidth: number, slideWidth: number): number => {
  const enlargement = shapes.some((shape) => shape.kind === "text" && shape.smartArt !== undefined)
    ? 1
    : MAX_ENLARGEMENT;
  return Math.min(MAX_WIDTH_PX / viewWidth, (enlargement * Math.min(MAX_WIDTH_PX, slideWidth)) / slideWidth);
};

/** The opening `<svg>` tag, the arrow markers when a connector uses one, and the white background over the whole viewBox. */
const openingOf = (view: Rect, scale: number, arrows: boolean): string =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${formatNumber(view.w * scale)}" height="${formatNumber(view.h * scale)}" viewBox="${formatNumber(view.x)} ${formatNumber(view.y)} ${formatNumber(view.w)} ${formatNumber(view.h)}">` +
  (arrows ? MARKERS : "") +
  `<rect x="${formatNumber(view.x)}" y="${formatNumber(view.y)}" width="${formatNumber(view.w)}" height="${formatNumber(view.h)}" fill="#fff"/>`;

/** The offered pictures written into their slots, in shape-tree order, while the SVG stays within `MAX_SVG_BYTES`; `sizeWithoutPictures`:
 *  the bytes of everything else. One that would pass the cap is left out and counted, the next ones are still tried. */
const embedPictures = (parts: string[][], sizeWithoutPictures: number, env: PictureEnv): void => {
  let size = sizeWithoutPictures;
  for (const picture of env.offered) {
    // Base64 is ASCII, one byte per character, 4 of them per 3 bytes padded:
    // the tag's size is exact BEFORE encoding, so a picture the cap drops is
    // never encoded at all.
    const tagLength = picture.head.length + 4 * Math.ceil(picture.bytes.length / 3) + 3;
    if (size + tagLength > MAX_SVG_BYTES) {
      leaveOutPicture(picture.shape, "picture-over-cap", "over-cap", env);
      continue;
    }
    size += tagLength;
    parts[picture.slot]!.push(`${picture.head}${Buffer.from(picture.bytes).toString("base64")}"/>`);
  }
};

/** `defaultTextSizes`: the deck's `p:defaultTextStyle` sizes by level
 *  (`PptxReadResult.defaultTextSizes`), for the runs of a free shape that
 *  write none; without it such text is left out and counted. */
export const renderSlideSvg = (
  shapes: readonly PptxShape[],
  slide: Size,
  defaultTextSizes: Readonly<Record<number, number>> = {},
): SlideSvg => {
  const omitted: Partial<Record<SlideSvgOmission, number>> = {};
  const omit = (reason: SlideSvgOmission): void => {
    omitted[reason] = (omitted[reason] ?? 0) + 1;
  };
  // One slot per shape, in shape-tree order: the paint order. A picture's
  // slot is filled once the size of everything else is known.
  const parts: string[][] = [];
  let overflowingText = 0;
  let arrows = false;
  const slideWidth = toPoints(slide.cx);
  const slideHeight = toPoints(slide.cy);
  const drawn: Extent = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
  const textEnv: DrawEnv = {
    page: { x: 0, y: 0, w: slideWidth, h: slideHeight },
    defaultSizes: defaultTextSizes,
    filled: [],
    omit,
    drawn,
    smartArtText: [],
  };
  const pictureEnv: PictureEnv = { omit, picturesNotDrawn: [], offered: [], drawn };
  for (const shape of shapes) {
    const elements: string[] = [];
    parts.push(elements);
    if (shape.box === undefined) {
      omit("no-box");
      if (shape.kind === "picture")
        pictureEnv.picturesNotDrawn.push({ spTreeIndex: shape.spTreeIndex, reason: "no-box" });
      continue;
    }
    switch (shape.kind) {
      case "text":
        if (drawText(shape, textEnv, elements).overflow) overflowingText++;
        break;
      case "connector":
        if (drawConnector(shape, elements, omit, drawn)) arrows = true;
        break;
      case "table":
      case "graphic":
        omit("frame");
        break;
      case "picture":
        offerPicture(shape, parts.length - 1, pictureEnv);
        break;
    }
  }

  const view = viewBoxOf(drawn, slideWidth, slideHeight);
  const open = openingOf(view, displayScale(shapes, view.w, slideWidth), arrows);
  const close = "</svg>";
  const sizeWithoutPictures =
    utf8Length(open) +
    utf8Length(close) +
    parts.reduce((sum, part) => sum + part.reduce((partSum, element) => partSum + utf8Length(element), 0), 0);
  embedPictures(parts, sizeWithoutPictures, pictureEnv);
  // Shape-tree order: the cap is decided after the walk.
  pictureEnv.picturesNotDrawn.sort((first, second) => first.spTreeIndex - second.spTreeIndex);
  return {
    svg: open + parts.map((part) => part.join("")).join("") + close,
    omitted,
    overflowingText,
    picturesNotDrawn: pictureEnv.picturesNotDrawn,
    smartArtText: textEnv.smartArtText,
  };
};
