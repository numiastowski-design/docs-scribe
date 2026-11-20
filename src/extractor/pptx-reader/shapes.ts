/** The shapes of one slide, as its shape tree declares them: every visible
 *  `p:sp` with a text body, every visible `p:graphicFrame` (a table read cell
 *  by cell; a chart, SmartArt, OLE object or other frame named, not read),
 *  every visible `p:pic` (its image, or why it is not published), every
 *  visible `p:cxnSp` (the shapes it links, its arrow heads), groups walked
 *  into and flattened, the shapes an `mc:AlternateContent` offers in its
 *  first `mc:Choice` — its `mc:Fallback` only when that Choice yields
 *  nothing and the Fallback brings real content;
 *  each paragraph's runs as the strings written in the
 *  file (`<a:t>2</a:t>` stays the string "2"), the placeholder (`p:ph`) each
 *  text shape declares, and each shape's id, box on the slide and group;
 *  for the slide drawing (`extractor/slide-svg.ts`), a placeholder's
 *  inherited text sizes (`pptx-reader/text-size.ts`), each shape's flips,
 *  its fill (exact colour, none or unknown: `pptx-reader/color.ts`) and
 *  whether its line shows, a round rectangle's corner, where its
 *  text sits, whether it is turned, shrunk (`a:normAutofit`) or unwrapped
 *  (`wrap="none"`), its written insets, whether its box grows with its text
 *  (`a:spAutoFit`), a connector's two ends and a bent connector's vertices.
 *  Also decides, for a picture, the 1x1 pixel case, the off-slide rule and the structure counts.
 *  NOT here: which shape is the slide's title (`pptx-reader/index.ts`), what a
 *  run or a paragraph carries (`pptx-reader/text.ts`), a table's cells
 *  (`pptx-reader/table.ts`), a picture's file bytes (`pptx-reader/picture.ts`), a
 *  chart's data (`pptx-reader/chart.ts`), how a box is
 *  computed (`pptx-reader/geometry.ts`), what the boxes mean for the layout
 *  (`layout/slide-layout.ts`), nor
 *  which slide is read (`pptx-reader/index.ts`).
 *  Called by: `pptx-reader/index.ts`; `readBodyPr` by `extractor/smartart-drawing.ts`; its types by
 *  `pptx-reader/shape-parts.ts` and `docx-reader/shapes.ts`. */

import { findInternalTarget, type Relationship } from "../ooxml/package";
import type { Box } from "../../layout/slide-elements";
import { type ReaderLimits } from "doc-ir";
import { NO_CAP } from "doc-ir/internal";
import { ARROW_GEOMETRIES } from "../../layout/slide-layout";
import { placeholderOf } from "./placeholder";
import { connectorBody, turnedByGroupField, readAvLstGuides, readPrstGeom, shapeBase } from "./shape-parts";
import type { Size } from "../../layout/slide-elements";
import { isOnePixel } from "../image-reduction";
import { imageOnlyFallback, POWERPOINT_TEXT_TAGS, POWERPOINT_UNDERSTOOD, renderedBranch } from "../markup-compat";
import { findSmartArtDataPart } from "../smartart";
import { type ColorContext, type PptxFill, solidFillColor, resolveFillRef, findWrittenFill } from "./color";
import { isHiddenFlag } from "../../lib/ooxml-bool";
import { attributeOf, findChild, findChildren, childNodesOf, qualifiedNameOf, type XNode } from "../ooxml/xml";
import type { PptxChart } from "./chart";
import {
  type Frame,
  frameOf,
  type GroupLevel,
  readGroupLevel,
  isWhollyOutside,
  type LineEnds,
  placePicture,
  type Point,
  composeGroupTransforms,
  type Transform,
} from "./geometry";
import { type PptxMedia, type PptxPictureContent, readPictureWidthPt, readPicture } from "./picture";
import { type PptxTable, readTable } from "./table";
import { inheritedBullet } from "./bullet";
import { parsePercentageAsFraction, type PptxParagraph, readParagraph, type RunColors } from "./text";
import { inheritedTextColor, type TextStyleSources } from "./text-color";
import { inheritedTextSizes } from "./text-size";

interface ShapeBase {
  /** Rank of the shape in document order among every element of the shape
   *  tree (groups, pictures, frames and hidden shapes included, a group before
   *  its members). Stable when later steps start reading more element kinds.
   *  An `mc:AlternateContent` takes no rank of its own: the shapes of its
   *  chosen `mc:Choice` stand where it stands. */
  spTreeIndex: number;
  /** `p:cNvPr id` as written: what a connector's `a:stCxn`/`a:endCxn` names. */
  id?: string;
  /** The box the shape is SEEN in on the slide, in integer EMU: groups
   *  carried to slide coordinates, rotation turned into the box of the
   *  rotated frame (`pptx-reader/geometry.ts#computeVisibleBox`). Absent when the
   *  shape declares no frame (a placeholder inheriting its layout's). */
  box?: Box;
  /** `rot` as written (60 000ths of a degree, clockwise), when not 0. The box
   *  already accounts for it; kept for drawing the shape. */
  rot?: number;
  /** `flipH`/`flipV` of the shape's own frame, when written true. The box
   *  does not show them; a flipped picture is mirrored, a text shape flipped
   *  vertically shows its text upside down (the slide drawing omits both). */
  flipH?: true;
  flipV?: true;
  /** The `spTreeIndex` of the TOP-LEVEL group (`p:grpSp` directly in
   *  `p:spTree`) the shape belongs to, however deep; absent outside any group.
   *  A group is one element of the slide and its members are sorted inside it
   *  (MIKSUDOC « Spécification fonctionnelle » §9.3.2): a nested group decides nothing of its own,
   *  so only the outermost is recorded. The group's own box is not: its
   *  members' boxes are the part of it that shows. */
  group?: number;
}

/** A `p:sp` with a text body. */
export interface PptxTextShape extends ShapeBase {
  kind: "text";
  /** One entry per `a:p`, empty paragraphs included. */
  paragraphs: PptxParagraph[];
  /** The shape's `p:nvSpPr/p:nvPr/p:ph`, attributes as written; absent when
   *  the shape is no placeholder. */
  placeholder?: { type?: string; idx?: string };
  /** For a placeholder, the size in points a run writing none inherits, by
   *  paragraph level (0-based), for the levels its paragraphs use
   *  (`pptx-reader/text-size.ts`); absent for a free shape or when nothing
   *  in the chain writes one. For the slide drawing only. */
  inheritedSizes?: Readonly<Record<number, number>>;
  /** `a:prstGeom prst` as written ("rect", "upArrow"…): whether the shape is
   *  an arrow is decided by the layout, not here. Absent for a custom
   *  geometry (`a:custGeom`) or none. */
  geometry?: string;
  /** `a:prstGeom/a:avLst/a:gd name="adj"` written `val N`: the corner size of
   *  a "roundRect", in 100 000ths of the shorter side. Absent: not written
   *  (the preset default applies) or not a plain value. */
  adj?: number;
  /** The shape's fill, written on `p:spPr` or, when `p:spPr` writes none,
   *  taken from `p:style` (`a:fillRef idx` above 0, through the theme's fill
   *  styles): its exact colour (`pptx-reader/color.ts`), "none" (`a:noFill`,
   *  or nothing written nor referenced), or "unknown" (a modifier, a
   *  gradient, a pattern, a picture). A text box writes `a:noFill`: "none". */
  fill: PptxFill;
  /** The shape's line shows: `a:ln` does not write `a:noFill`, and writes a
   *  fill or takes a line from `p:style` (`a:lnRef idx` above 0). Its colour
   *  is not read: a line is drawn black. */
  hasVisibleLine: boolean;
  /** `a:bodyPr anchor`: the text sits at the middle ("ctr") or the bottom
   *  ("b") of the box; absent: top, or not written (an inherited anchor is
   *  not read). */
  anchor?: "ctr" | "b";
  /** `a:bodyPr` turns the text: `vert` other than "horz", or a `rot` other
   *  than 0. */
  isTextTurned?: true;
  /** `a:bodyPr/a:normAutofit fontScale` as a fraction (62 500 thousandths of
   *  a percent, or "62.5%", is 0.625): every size of the text is drawn times
   *  this. Absent: not written, 100 %, or not in (0, 100 %). */
  fontScale?: number;
  /** `a:bodyPr/a:normAutofit lnSpcReduction` as a fraction (20 000 is 0.2):
   *  the line height is reduced by this much. Absent: not written, 0, or not
   *  in (0, 100 %). */
  lineSpaceReduction?: number;
  /** `a:bodyPr wrap="none"`: lines break only where the text breaks them. */
  isUnwrapped?: true;
  /** `a:bodyPr lIns`, `tIns`, `rIns`, `bIns` in EMU, each only when written
   *  as a whole number of at least 0; absent when none is. A side not
   *  written takes the slide drawing's default. */
  insets?: { left?: number; top?: number; right?: number; bottom?: number };
  /** `a:bodyPr/a:spAutoFit`: the box grows with its text, so the written
   *  `a:ext` is the size already grown to hold it. */
  hasSpAutoFit?: true;
  /** An enclosing group turns or mirrors the shape (`geometry.ts#isTurnOrMirror`
   *  of the groups' map): the slide drawing leaves it out. */
  isTurnedByGroup?: true;
  /** A shape of a SmartArt's saved drawing (`dsp:sp`, `extractor/smartart-drawing.ts`), not of the slide's tree; present only for those.
   *  All boxes are in EMU on the slide (the drawing's own space carried to the frame's corner, scale 1). `frame`: the shape's frame BEFORE
   *  its turn `turn` (any angle, 60 000ths of a degree folded into [0, 360), 0 when none; `box` is the box AFTER it, as for any shape).
   *  `textBox`: the text rectangle (`dsp:txXfrm`, else the frame), laid out horizontally and then turned by `textTurn` (`dsp:txXfrm rot`
   *  plus the `vert` of `a:bodyPr`, 0 when none) about its own centre, then by `turn` about the frame's centre. `vf`: a hexagon's `vf`
   *  guide; `adj1`, `adj2`: a chord's start and end angles. `textSwap`: the lines are laid in `textBox` with its sides swapped about its centre (`a:bodyPr vert` of 90 or 270). Such a shape never has `rot`: the turn is `turn`, drawn as `rotate`. */
  smartArt?: {
    frame: Box;
    turn: number;
    textBox: Box;
    textTurn: number;
    textSwap?: true;
    vf?: number;
    adj1?: number;
    adj2?: number;
  };
}

/** A `p:cxnSp`: a line drawn between two shapes. Never has text. */
export interface PptxConnectorShape extends ShapeBase {
  kind: "connector";
  /** `a:stCxn id` and `a:endCxn id`: the `id` of the shapes it is glued to,
   *  when declared. */
  from?: string;
  to?: string;
  /** An arrow head is drawn at the line's start (`a:headEnd`) or end
   *  (`a:tailEnd`): its `type` is present and not "none". Only `a:ln` on the
   *  connector is read; a line end inherited from `p:style` is not. */
  startArrow: boolean;
  endArrow: boolean;
  /** `a:prstGeom prst` as written ("straightConnector1", "bentConnector3"…). */
  geometry?: string;
  /** Its line shows: `a:ln` does not write `a:noFill`, and writes a fill or
   *  takes a line from `p:style` (`a:lnRef idx` above 0). */
  hasVisibleLine: boolean;
  /** Where its line starts and ends on the slide, flips and groups applied
   *  (`pptx-reader/geometry.ts#computeLineEnds`); absent without a frame. */
  ends?: LineEnds;
  /** For `bentConnector2` to `bentConnector5` only: the polyline's vertices
   *  on the slide, start to end, from the preset's formulas and the written
   *  `adj1`…`adj3`, through the same flips, rotation and groups as `ends`
   *  (`pptx-reader/geometry.ts#computeBentPath`); absent without a frame. */
  path?: Point[];
}

/** A `p:graphicFrame` holding an `a:tbl`. */
interface PptxTableShape extends ShapeBase {
  kind: "table";
  table: PptxTable;
}

/** Any other `p:graphicFrame`, named by what its `a:graphicData` holds:
 *  `c:chart` → "chart", a `dgm:` element → "smartart", the OLE `uri` →
 *  "ole", anything else → "other". Nothing inside is read here. A video or
 *  a sound is a `p:pic`, never a graphic frame. */
export interface PptxGraphicShape extends ShapeBase {
  kind: "graphic";
  graphic: "chart" | "smartart" | "ole" | "other";
  /** For a chart only: the archive path of its chart part, when the
   *  `c:chart r:id` resolves to an internal relationship. */
  part?: string;
  /** For a chart only: what its part holds, when it has series
   *  (`pptx-reader/chart.ts`), or `"withheld"` when a bound a damaged or missing part keeps it out whole; read by
   *  `pptx-reader/index.ts`, after the shape walk. Absent: the chart has no series. */
  chart?: PptxChart | "withheld";
  /** For a SmartArt only: the archive path of its data-model part, when
   *  `dgm:relIds r:dm` resolves to an internal relationship. */
  dataPart?: string;
  /** For a SmartArt only: its words in data-model order, one piece per
   *  point (`extractor/smartart.ts`); read by `pptx-reader/index.ts`, after
   *  the shape walk. */
  words?: string[];
  /** For a SmartArt only: its saved drawing (`extractor/smartart-drawing.ts`), read after the shape walk by `pptx-reader/index.ts`:
   *  the shapes to add to the slide's, on the slide; "not-drawn" when a drawing exists but is kept out (past the shape caps, or its
   *  frame in a group that scales, turns or mirrors it). Absent: no drawing to draw (none saved, unlinked, damaged, empty). */
  drawing?: { shapes: PptxTextShape[] } | "not-drawn";
  /** An enclosing group scales, turns or mirrors the frame: the drawing, at scale 1 from the frame's corner, would be misplaced. */
  isScaledOrTurnedByGroup?: true;
}

/** A `p:pic`: its image, or the reason it is not published. No alt text. */
export type PptxPictureShape = ShapeBase & {
  kind: "picture";
  /** The declared display width in points (`p:spPr/a:xfrm/a:ext cx`), when
   *  written. Inside a group, in the group's coordinates: `box` has the width
   *  on the slide. */
  widthPt?: number;
  /** Where the slide drawing places the image, in EMU
   *  (`pptx-reader/geometry.ts#placePicture`). Without `matrix`: the
   *  declared frame (`a:off`, `a:ext`) carried to the slide BEFORE the
   *  picture's own `rot` and flips. With `matrix`: the frame as written, in
   *  its group's coordinates. Absent with `groupStretched`, or without a
   *  frame. */
  frame?: Box;
  /** A group turns, mirrors or stretches the picture unevenly: the whole map
   *  from `frame` to the slide, EMU, its own `rot` and flips included. */
  matrix?: Transform;
  /** The picture turns (other than 0° or 180°) inside a group stretched
   *  unevenly: neither `frame` nor `matrix` is given. */
  groupStretched?: true;
} & PptxPictureContent;

export type PptxShape = PptxTextShape | PptxTableShape | PptxGraphicShape | PptxPictureShape | PptxConnectorShape;

const OLE_URI = "http://schemas.openxmlformats.org/presentationml/2006/ole";

/** What a graphic frame holds, by the child tags of its `a:graphicData`. */
const readFrame = (
  frame: XNode,
  spTreeIndex: number,
  relationships: ReadonlyMap<string, Relationship>,
  slideNumber: number | undefined,
): PptxTableShape | PptxGraphicShape => {
  const data = findChild(findChild(frame, "a:graphic") ?? {}, "a:graphicData") ?? {};
  const tbl = findChild(data, "a:tbl");
  if (tbl) return { kind: "table", spTreeIndex, table: readTable(tbl, relationships, slideNumber) };
  const chart = findChild(data, "c:chart");
  if (chart) {
    const part = findInternalTarget(relationships, attributeOf(chart, "r:id"));
    return {
      kind: "graphic",
      spTreeIndex,
      graphic: "chart",
      ...(part !== undefined ? { part } : {}),
    };
  }
  if (childNodesOf(data).some((child) => qualifiedNameOf(child)?.startsWith("dgm:"))) {
    const relIds = findChild(data, "dgm:relIds");
    const dataPart = relIds === undefined ? undefined : findSmartArtDataPart(relationships, relIds);
    return { kind: "graphic", spTreeIndex, graphic: "smartart", ...(dataPart === undefined ? {} : { dataPart }) };
  }
  if (attributeOf(data, "uri") === OLE_URI) return { kind: "graphic", spTreeIndex, graphic: "ole" };
  return { kind: "graphic", spTreeIndex, graphic: "other" };
};

/** The non-visual properties element of each shape-tree element that has one. */
const NON_VISUAL = ["p:nvSpPr", "p:nvGrpSpPr", "p:nvPicPr", "p:nvGraphicFramePr", "p:nvCxnSpPr", "p:nvContentPartPr"];

/** The group's own properties, not members of its tree. */
const GROUP_PROPERTIES = new Set(["p:nvGrpSpPr", "p:grpSpPr", "p:extLst"]);

/** ECMA-376: `hidden="1"` (or "true", blanks and case tolerated) on `p:cNvPr` hides the shape. */
const isHidden = (element: XNode): boolean => {
  for (const tag of NON_VISUAL) {
    const nonVisual = findChild(element, tag);
    if (!nonVisual) continue;
    return isHiddenFlag(attributeOf(findChild(nonVisual, "p:cNvPr") ?? {}, "hidden"));
  }
  return false;
};

/** `p:style` gives the shape a fill or a line: its reference's `idx` is above 0. */
const styleGives = (element: XNode, ref: "a:fillRef" | "a:lnRef"): boolean => {
  const idx = Number(attributeOf(findChild(findChild(element, "p:style") ?? {}, ref) ?? {}, "idx"));
  return Number.isInteger(idx) && idx > 0;
};

/** The line `a:ln` writes shows unless its fill is `a:noFill`; when it
 *  writes no fill, the line `p:style` refers to. */
const lineShows = (element: XNode, spPr: XNode | undefined): boolean => {
  const fill = findWrittenFill(findChild(spPr ?? {}, "a:ln"));
  return fill === undefined ? styleGives(element, "a:lnRef") : qualifiedNameOf(fill) !== "a:noFill";
};

/** The fill `p:spPr` writes, else the one `p:style` refers to. */
const fillOf = (element: XNode, spPr: XNode | undefined, context: ColorContext | undefined): PptxFill => {
  const written = findWrittenFill(spPr);
  if (written === undefined) {
    if (!styleGives(element, "a:fillRef")) return "none";
    return resolveFillRef(findChild(findChild(element, "p:style") ?? {}, "a:fillRef")!, context);
  }
  switch (qualifiedNameOf(written)) {
    case "a:noFill":
      return "none";
    case "a:solidFill":
      return solidFillColor(written, context) ?? "unknown";
    default:
      return "unknown";
  }
};

/** How the runs of a text shape read their colour: through `context`, and
 *  a run writing no fill takes the colour it inherits
 *  (`pptx-reader/text-color.ts`). */
const runColors = (
  element: XNode,
  body: XNode,
  context: ColorContext | undefined,
  sources: TextStyleSources,
): RunColors => ({
  ...(context === undefined ? {} : { context }),
  inherited: inheritedTextColor(element, body, sources, context),
  inheritedBullet: inheritedBullet(element, body, sources),
});

/** ST_Coordinate32 of an inset: a whole number of EMU. A negative inset or
 *  a universal measure ("0.1in") is left unread: the default applies. */
const insetOf = (bodyPr: XNode, name: string): number | undefined => {
  const raw = attributeOf(bodyPr, name);
  return raw !== undefined && /^\d+$/.test(raw) ? Number(raw) : undefined;
};

/** The insets `a:bodyPr` writes (`lIns`, `tIns`, `rIns`, `bIns`); none when it writes none that reads. */
const insetsOf = (bodyPr: XNode): Pick<PptxTextShape, "insets"> => {
  const insets: NonNullable<PptxTextShape["insets"]> = {};
  for (const [side, name] of [
    ["left", "lIns"],
    ["top", "tIns"],
    ["right", "rIns"],
    ["bottom", "bIns"],
  ] as const) {
    const value = insetOf(bodyPr, name);
    if (value !== undefined) insets[side] = value;
  }
  return Object.keys(insets).length > 0 ? { insets } : {};
};

/** A fraction of `a:normAutofit` when it lies in (0, 1): what shrinks the text. */
const isShrinkFraction = (fraction: number | undefined): fraction is number =>
  fraction !== undefined && fraction > 0 && fraction < 1;

/** What `a:bodyPr/a:normAutofit` shrinks: the font scale and the line spacing reduction, each only in (0, 1). */
const autofitOf = (bodyPr: XNode): Pick<PptxTextShape, "fontScale" | "lineSpaceReduction"> => {
  const autofit = findChild(bodyPr, "a:normAutofit");
  const fontScale = parsePercentageAsFraction(autofit === undefined ? undefined : attributeOf(autofit, "fontScale"));
  const reduction = parsePercentageAsFraction(
    autofit === undefined ? undefined : attributeOf(autofit, "lnSpcReduction"),
  );
  return {
    ...(isShrinkFraction(fontScale) ? { fontScale } : {}),
    ...(isShrinkFraction(reduction) ? { lineSpaceReduction: reduction } : {}),
  };
};

/** What `a:bodyPr` says of where and how the text sits. */
export const readBodyPr = (
  body: XNode,
): Pick<
  PptxTextShape,
  "anchor" | "isTextTurned" | "fontScale" | "lineSpaceReduction" | "isUnwrapped" | "insets" | "hasSpAutoFit"
> => {
  const bodyPr = findChild(body, "a:bodyPr") ?? {};
  const anchor = attributeOf(bodyPr, "anchor");
  const vert = attributeOf(bodyPr, "vert");
  const rot = Number(attributeOf(bodyPr, "rot") ?? "0");
  const turned = (vert !== undefined && vert !== "horz") || rot !== 0;
  return {
    ...insetsOf(bodyPr),
    ...(findChild(bodyPr, "a:spAutoFit") ? { hasSpAutoFit: true as const } : {}),
    ...(anchor === "ctr" || anchor === "b" ? { anchor } : {}),
    ...(turned ? { isTextTurned: true as const } : {}),
    ...autofitOf(bodyPr),
    ...(attributeOf(bodyPr, "wrap") === "none" ? { isUnwrapped: true as const } : {}),
  };
};

/** The frame element of a shape-tree element: `p:xfrm` for a graphic frame,
 *  `p:spPr/a:xfrm` for the others. */
const xfrmOf = (element: XNode, tag: string): XNode | undefined =>
  tag === "p:graphicFrame" ? findChild(element, "p:xfrm") : findChild(findChild(element, "p:spPr") ?? {}, "a:xfrm");

/** The frame of a shape-tree element, whichever kind: a group writes its `a:xfrm` in `p:grpSpPr`. */
const frameOfElement = (element: XNode, tag: string): Frame | undefined =>
  tag === "p:grpSp"
    ? frameOf(findChild(findChild(element, "p:grpSpPr") ?? {}, "a:xfrm"))
    : frameOf(xfrmOf(element, tag));

/** Whether a top-level element lies WHOLLY outside the slide, so that it is not shown in a slideshow, print or export
 *  (rule « Rien d'invisible », 09/10/2026; the same rule as Word's `docx-reader/objects.ts#isOffPage`).
 *  Decided only on what is declared: a rotated element, an unknown frame (a placeholder inheriting its box from the
 *  layout) or an unknown slide size is kept. A group is judged by its own frame; the members of a group are never judged
 *  alone (their coordinates are in the group's child space). */
const isOffSlide = (frame: Frame | undefined, slide: Size | undefined): boolean => {
  if (!frame || !slide || frame.rot !== 0) return false;
  return isWhollyOutside(frame.x, frame.cx, slide.cx) || isWhollyOutside(frame.y, frame.cy, slide.cy);
};

/** `p:cNvPr id` of an element, from whichever non-visual properties it has. */
const idOf = (element: XNode): string | undefined => {
  for (const tag of NON_VISUAL) {
    const nonVisual = findChild(element, tag);
    if (nonVisual) return attributeOf(findChild(nonVisual, "p:cNvPr") ?? {}, "id");
  }
  return undefined;
};

/** The element count of an `mc:AlternateContent`: the larger of its rendered branch and the image-only Fallback that branch may give way to. */
const alternateContentCount = (element: XNode): number => {
  const branch = renderedBranch(element, POWERPOINT_UNDERSTOOD);
  // An unread Choice may give way to an image-only Fallback: the larger variant bounds the count.
  const fallback = branch ? imageOnlyFallback(element, branch, POWERPOINT_TEXT_TAGS) : undefined;
  return Math.max(branch ? countElements(branch) : 0, fallback ? countElements(fallback) : 0);
};

/** The element count of a shape tree, STRUCTURAL — no text, colour, media
 *  nor relationship is read, so that the element-cap withhold is decided
 *  before any media is inflated and a withheld slide spends nothing of the
 *  shared budget.
 *  Same ranking as `readShapes`: every element counts, a group counts
 *  itself then its members, an `mc:AlternateContent` counts ONCE — here
 *  as the LARGER of its two variants, an upper bound of the walk's own
 *  variant-kept count (chosen bias, written: a slide at the very edge of
 *  `MAX_SLIDE_ELEMENTS` whose alternatives are asymmetric may be withheld
 *  one variant early — the cap is a starting value anyway). */
export const countElements = (tree: XNode): number => {
  let count = 0;
  for (const element of childNodesOf(tree)) {
    const tag = qualifiedNameOf(element);
    if (tag === undefined || tag === "#text" || GROUP_PROPERTIES.has(tag)) continue;
    if (tag === "mc:AlternateContent") {
      count += alternateContentCount(element);
      continue;
    }
    count++;
    if (tag === "p:grpSp") count += countElements(element);
  }
  return count;
};

/** What a slide's walk reads against, the same for every element of it. */
export interface ReadEnv {
  /** The slide part's relationships, which link targets resolve against. */
  relationships: ReadonlyMap<string, Relationship>;
  /** The image parts the relationships name. */
  media: PptxMedia;
  context?: ColorContext | undefined;
  /** The parts a run's colour is inherited from (`pptx-reader/text-color.ts`); absent: none. */
  sources?: TextStyleSources | undefined;
  slideSize?: Size | undefined;
  /** What a `slidenum` field of this slide displays. */
  slideNumber?: number | undefined;
  /** The element cap the walk applies (`ReaderLimits.slideElements`); absent: NONE — `NO_CAP`, nothing withheld. */
  limits?: ReaderLimits | undefined;
}

/** The walk's running state: the shapes read, the next rank, the hidden elements left out below the cap (a hidden group for one: its members are not walked). */
interface WalkState {
  shapes: PptxShape[];
  rank: number;
  hiddenShapes: number;
}

/** Where in the shape tree an element stands: the enclosing groups (outermost first), the top-level group it belongs to, and the matrix the groups give. */
interface Scope {
  levels: readonly GroupLevel[];
  group: number | undefined;
  transform: Transform;
}

/** A `p:sp`: a text shape; a bodiless block arrow, an empty one; any other bodiless shape, nothing. */
const readTextShape = (
  env: ReadEnv,
  element: XNode,
  base: ShapeBase,
  transform: Transform,
): PptxTextShape | undefined => {
  const { relationships, context, sources = {} } = env;
  const body = findChild(element, "p:txBody");
  // A `p:sp` with NO text body is kept only when its preset geometry is a
  // BLOCK ARROW, as an empty text shape carrying that geometry (7b of the
  // arbitrages, 28/09/2026) — the layout
  // counts those (`slide-layout.ts#ARROW_GEOMETRIES`), so a slide of boxes
  // joined by arrow SHAPES reaches the diagram form like one joined by
  // connectors. Any other bodiless shape (a decorative rectangle…) stays
  // out: no arrow, no content.
  if (!body) {
    const spPr = findChild(element, "p:spPr");
    const geometry = readPrstGeom(spPr);
    if (geometry === undefined || !ARROW_GEOMETRIES.has(geometry)) return undefined;
    return {
      kind: "text",
      ...base,
      paragraphs: [],
      geometry,
      fill: fillOf(element, spPr, context),
      hasVisibleLine: lineShows(element, spPr),
      ...turnedByGroupField(transform),
    };
  }
  const placeholder = placeholderOf(element);
  // Built once per shape: its inherited colours are the same for every paragraph.
  const colors = {
    ...runColors(element, body, context, sources),
    ...(env.slideNumber === undefined ? {} : { slideNumber: env.slideNumber }),
  };
  const spPr = findChild(element, "p:spPr");
  const geometry = readPrstGeom(spPr);
  const adj = readAvLstGuides(spPr)["adj"];
  const paragraphs = findChildren(body, "a:p").map((paragraph) => readParagraph(paragraph, relationships, colors));
  const inheritedSizes = inheritedTextSizes(
    element,
    body,
    paragraphs.map((paragraph) => paragraph.level ?? 0),
    sources,
  );
  return {
    kind: "text",
    ...base,
    paragraphs,
    ...(placeholder === undefined ? {} : { placeholder }),
    ...(inheritedSizes === undefined ? {} : { inheritedSizes }),
    ...(geometry === undefined ? {} : { geometry }),
    ...(adj === undefined ? {} : { adj }),
    fill: fillOf(element, spPr, context),
    hasVisibleLine: lineShows(element, spPr),
    ...readBodyPr(body),
    ...turnedByGroupField(transform),
  };
};

/** A `p:cxnSp`: the shapes it links by id, its arrow heads, its line. */
const readConnector = (
  element: XNode,
  base: ShapeBase,
  frame: Frame | undefined,
  transform: Transform,
): PptxConnectorShape => {
  const connectorProperties = findChild(findChild(element, "p:nvCxnSpPr") ?? {}, "p:cNvCxnSpPr") ?? {};
  const from = attributeOf(findChild(connectorProperties, "a:stCxn") ?? {}, "id");
  const to = attributeOf(findChild(connectorProperties, "a:endCxn") ?? {}, "id");
  const spPr = findChild(element, "p:spPr");
  return {
    kind: "connector",
    ...base,
    ...(from === undefined ? {} : { from }),
    ...(to === undefined ? {} : { to }),
    ...connectorBody(spPr, frame, transform, lineShows(element, spPr)),
  };
};

/** A `p:pic`: its image (or why it is not published) and where the slide places it. */
const readPicShape = (
  env: ReadEnv,
  element: XNode,
  base: ShapeBase,
  frame: Frame | undefined,
  { levels, transform }: Scope,
): PptxPictureShape => {
  // The width on the slide: the declared one times the length the
  // groups give a unit along the picture's own x axis.
  const declared = readPictureWidthPt(element);
  const widthPt = declared === undefined ? undefined : declared * Math.hypot(transform.a, transform.b);
  return {
    kind: "picture",
    ...base,
    ...(widthPt === undefined ? {} : { widthPt }),
    ...picturePlacement(frame, levels),
    ...readPicture(element, env.relationships, env.media),
  };
};

/** Where the slide drawing places a picture (`geometry.ts#placePicture`): nothing without a frame, `groupStretched` for a turn inside
 *  a group stretched unevenly. */
const picturePlacement = (
  frame: Frame | undefined,
  levels: readonly GroupLevel[],
): Pick<PptxPictureShape, "frame" | "matrix" | "groupStretched"> => {
  if (frame === undefined) return {};
  const placed = placePicture(frame, levels);
  if (placed === "stretched-turn") return { groupStretched: true };
  return placed;
};

/** The shape a kept, visible element of the tree reads as; nothing for a kind this reader does not read. */
const readLeaf = (
  env: ReadEnv,
  element: XNode,
  tag: string,
  spTreeIndex: number,
  scope: Scope,
): PptxShape | undefined => {
  const { group, transform } = scope;
  const frame = frameOf(xfrmOf(element, tag));
  const id = idOf(element);
  const base = {
    ...shapeBase(spTreeIndex, frame, transform),
    ...(id === undefined ? {} : { id }),
    ...(group === undefined ? {} : { group }),
  };
  if (tag === "p:sp") return readTextShape(env, element, base, transform);
  if (tag === "p:cxnSp") return readConnector(element, base, frame, transform);
  if (tag === "p:graphicFrame") {
    const read = readFrame(element, spTreeIndex, env.relationships, env.slideNumber);
    const moved = transform.a === 1 && transform.d === 1 && transform.b === 0 && transform.c === 0;
    return {
      ...read,
      ...base,
      ...(read.kind === "graphic" && read.graphic === "smartart" && !moved
        ? { isScaledOrTurnedByGroup: true as const }
        : {}),
    };
  }
  if (tag === "p:pic") return readPicShape(env, element, base, frame, scope);
  return undefined;
};

/** An `mc:AlternateContent`: the branch PowerPoint renders is walked; when it yields nothing, its Fallback only if that is image-only. */
const walkAlternate = (env: ReadEnv, state: WalkState, element: XNode, { levels, group }: Scope): void => {
  const branch = renderedBranch(element, POWERPOINT_UNDERSTOOD);
  const startRank = state.rank;
  const startHidden = state.hiddenShapes;
  const before = state.shapes.length;
  if (branch) walkElements(env, state, childNodesOf(branch), levels, group);
  // M16 (02/10/2026): a Choice PowerPoint renders but this
  // reader cannot read leaves its Fallback to be read ONLY if that
  // carries no text — the same object, drawn as images. The variants
  // count ONCE toward `elementCount`: the Fallback's ranks replace the
  // empty Choice's (the larger count stands when the Fallback is empty too).
  const fallback =
    branch && state.shapes.length === before ? imageOnlyFallback(element, branch, POWERPOINT_TEXT_TAGS) : undefined;
  if (!fallback) return;
  const choiceRank = state.rank;
  state.rank = startRank;
  state.hiddenShapes = startHidden;
  walkElements(env, state, childNodesOf(fallback), levels, group);
  // Only 1×1 pixels (the documented junk, PLATFORM-STATUS §53): not content.
  const added = state.shapes.slice(before);
  const onlyPixels =
    added.length > 0 &&
    added.every((shape) => shape.kind === "picture" && "image" in shape && isOnePixel(shape.image.bytes));
  if (onlyPixels) state.shapes.length = before;
  if (state.shapes.length === before) state.rank = Math.max(choiceRank, state.rank);
};

/** One element of a shape tree other than an `mc:AlternateContent`: ranked, then counted past the cap, left out hidden or off the slide, walked into as a group, or read. */
const walkShapeElement = (env: ReadEnv, state: WalkState, element: XNode, tag: string, scope: Scope): void => {
  const { levels, group } = scope;
  const spTreeIndex = state.rank++;
  // Past the cap the element is COUNTED, never read: no text, media
  // nor layout paid (7a of the arbitrages, 28/09/2026 — the derivation
  // is on the constant). A group is still walked so the told total is
  // the slide's true one.
  if (spTreeIndex >= (env.limits?.slideElements ?? NO_CAP)) {
    if (tag === "p:grpSp") walkElements(env, state, childNodesOf(element), levels, group ?? spTreeIndex);
    return;
  }
  if (isHidden(element)) {
    state.hiddenShapes++;
    return;
  }
  // Rule « Rien d'invisible » (09/10/2026): top-level elements only, see `isOffSlide`.
  if (levels.length === 0 && isOffSlide(frameOfElement(element, tag), env.slideSize)) return;
  if (tag === "p:grpSp") {
    const level = readGroupLevel(findChild(element, "p:grpSpPr"));
    walkElements(
      env,
      state,
      childNodesOf(element),
      level === undefined ? levels : [...levels, level],
      group ?? spTreeIndex,
    );
    return;
  }
  const shape = readLeaf(env, element, tag, spTreeIndex, scope);
  if (shape !== undefined) state.shapes.push(shape);
};

/** `levels`: the enclosing groups, outermost first. */
const walkElements = (
  env: ReadEnv,
  state: WalkState,
  elements: readonly XNode[],
  levels: readonly GroupLevel[],
  group: number | undefined,
): void => {
  const scope: Scope = { levels, group, transform: composeGroupTransforms(levels) };
  for (const element of elements) {
    const tag = qualifiedNameOf(element);
    // Character data between elements (whitespace) is no shape.
    if (tag === undefined || tag === "#text" || GROUP_PROPERTIES.has(tag)) continue;
    if (tag === "mc:AlternateContent") walkAlternate(env, state, element, scope);
    else walkShapeElement(env, state, element, tag, scope);
  }
};

/** The visible shapes of a shape tree (`p:spTree` or `p:grpSp`), in
 *  document order, read against `env`.
 *  `mc:AlternateContent`: ONE branch is read, the one PowerPoint renders
 *  (`markup-compat.ts#renderedBranch`, ECMA-376 Part 3, AlternateContent):
 *  the first `mc:Choice` whose `Requires` names only namespaces PowerPoint
 *  understands, else the `mc:Fallback` (M15/M16, 02/10/2026: what Office
 *  DISPLAYS is published). A Choice PowerPoint understands but this reader
 *  cannot read gives way to its Fallback ONLY when that carries no text
 *  (`markup-compat.ts#imageOnlyFallback`: the same object, drawn as images;
 *  a Fallback with text is never read). What a branch holds is
 *  read by its canonical tags, as any other shape. */
export const readShapes = (
  tree: XNode,
  env: ReadEnv,
): { shapes: PptxShape[]; elementCount: number; hiddenShapes: number } => {
  const state: WalkState = { shapes: [], rank: 0, hiddenShapes: 0 };
  walkElements(env, state, childNodesOf(tree), [], undefined);
  return { shapes: state.shapes, elementCount: state.rank, hiddenShapes: state.hiddenShapes };
};
