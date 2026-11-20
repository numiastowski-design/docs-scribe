/** A SmartArt's SAVED DRAWING (`dsp:drawing`, [MS-ODRAWXML]: `word/diagrams/drawingN.xml`, `ppt/diagrams/drawingN.xml`, beside the
 *  data model) read into the shape model `extractor/slide-svg.ts` already draws, so a SmartArt is published as a picture.
 *  What is read, per `dsp:sp`: its frame (`a:xfrm`, with `rot`), its preset geometry and `adj`/`vf` guides, its fill and line (written
 *  on `dsp:spPr`, else the `dsp:style` references, through the part's theme colours), its text body, and `dsp:txXfrm` (the text
 *  rectangle, whose `rot` ADDS to the shape's: a hexagon turned 90 whose text rectangle is turned -90 shows its text horizontal).
 *  Any angle is kept as written: the drawer turns the outline about the frame's centre and the text about its rectangle's centre
 *  then the frame's; `a:bodyPr vert` adds 90 or 270 to the text.
 *  The drawing declares no extent (`dsp:grpSpPr` is empty on all 9 saved drawings measured): its space is the frame's, from the
 *  frame's corner, at scale 1 (`placeShapes`), never a ratio computed on the shapes' bounds.
 *  Also here: the model words a picture drawn this way does not show (`findUndrawnWords`).
 *  No saved drawing (other software, broken relationship, damaged part, no traced `dsp:sp`, over `ReaderLimits.slideElements`
 *  shapes): no picture, and the model's words are counted `smartart-not-drawn`; a damaged part never refuses the document. Office's
 *  layout engine (`layout*.xml`) is not reproduced.
 *  NOT here: how a shape is drawn (`extractor/slide-svg.ts`), which drawing a frame owns (`extractor/smartart.ts`: the data model's
 *  `dsp:dataModelExt relId`), the words of the model (`extractor/smartart.ts`), where the picture goes (`extractor/pptx.ts`,
 *  `docx-reader/drawing-objects.ts`).
 *  Called by: `pptx-reader/index.ts`, `docx-reader/drawing-objects.ts`, `extractor/pptx.ts`. */

import { type ReaderLimits } from "doc-ir";
import { NO_CAP } from "doc-ir/internal";
import { wordsOf } from "../lib/words";
import type { Box } from "../layout/slide-elements";
import { IDENTITY_MATRIX } from "doc-ir/internal";
import type { Relationship } from "./ooxml/package";
import { attributeOf, findChild, findChildren, findDescendants, qualifiedNameOf, type XNode } from "./ooxml/xml";
import {
  type ColorContext,
  elementChildrenOf,
  exactColorOf,
  type Hex,
  type PptxFill,
  readClrScheme,
  solidFillColor,
  resolveFillRef,
  findWrittenFill,
} from "./pptx-reader/color";
import { frameOf, type Frame, FULL_TURN, QUARTER_TURN, computeVisibleBox } from "./pptx-reader/geometry";
import { readAvLstGuides, readPrstGeom } from "./pptx-reader/shape-parts";
import { readBodyPr, type PptxTextShape } from "./pptx-reader/shapes";
import { type PptxParagraph, readParagraph, type RunColors } from "./pptx-reader/text";

/** An angle folded into [0, 360 degrees). */
const folded = (rot: number): number => ((rot % FULL_TURN) + FULL_TURN) % FULL_TURN;

/** The turn `a:bodyPr vert` adds to the text (`ST_TextVerticalType`): `vert` is 90 degrees clockwise,
 *  `vert270` 270. The others (`eaVert`, `mongolianVert`, `wordArtVert`, `wordArtVertRtl`) stack the glyphs: not a rotation, so no entry. */
const VERT_TURN: ReadonlyMap<string, number> = new Map([
  ["horz", 0],
  ["vert", QUARTER_TURN],
  ["vert270", 3 * QUARTER_TURN],
]);

const NO_RELATIONSHIPS: ReadonlyMap<string, Relationship> = new Map();

/** The shape's fill: `dsp:spPr` writes it, else `dsp:style/a:fillRef` names the theme's. */
const fillOf = (spPr: XNode | undefined, style: XNode | undefined, context: ColorContext | undefined): PptxFill => {
  const written = findWrittenFill(spPr);
  if (written === undefined) {
    const ref = findChild(style ?? {}, "a:fillRef");
    return ref === undefined ? "none" : resolveFillRef(ref, context);
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

/** The line shows unless `a:ln` writes `a:noFill`; with no fill written, when `dsp:style/a:lnRef` has an `idx` above 0. */
const lineShownOf = (spPr: XNode | undefined, style: XNode | undefined): boolean => {
  const written = findWrittenFill(findChild(spPr ?? {}, "a:ln"));
  if (written !== undefined) return qualifiedNameOf(written) !== "a:noFill";
  const idx = Number(attributeOf(findChild(style ?? {}, "a:lnRef") ?? {}, "idx"));
  return Number.isInteger(idx) && idx > 0;
};

/** The colour `dsp:style/a:fontRef` gives the text (a run writing none takes it); `undefined` when none or not exact. */
const fontColorOf = (style: XNode | undefined, context: ColorContext | undefined): Hex | undefined => {
  const colors = elementChildrenOf(findChild(style ?? {}, "a:fontRef") ?? {});
  return colors.length === 1 ? exactColorOf(colors[0]!, context) : undefined;
};

/** Whether `a:bodyPr upright` (xsd:boolean) is set. */
const isUpright = (upright: string | undefined): boolean => upright === "1" || upright === "true";

/** The text rectangle's frame, and the turn the text takes inside it: `dsp:txXfrm rot` plus the `vert` of `a:bodyPr`, both about the
 *  rectangle's centre ([MS-ODRAWXML] §2.5.40; the shape's own turn comes on top, about the frame's centre). `undefined` when the rule is
 *  not known and the words are left out: stacked glyphs, `upright`, a `bodyPr rot` (the standard does not say which of the two wins
 *  with `txXfrm rot`), a mirrored shape. 0 occurrences of those in 355 documents (measured 20/10/2026). */
const textPlacement = (
  sp: XNode,
  frame: Frame,
  body: XNode | undefined,
): { textFrame: Frame; textTurn: number | undefined; swapped: boolean } => {
  const writtenTextFrame = frameOf(findChild(sp, "dsp:txXfrm"));
  const textFrame = writtenTextFrame ?? frame;
  if (body === undefined) return { textFrame, textTurn: 0, swapped: false };
  const bodyPr = findChild(body, "a:bodyPr") ?? {};
  const upright = attributeOf(bodyPr, "upright");
  const vert = VERT_TURN.get(attributeOf(bodyPr, "vert") ?? "horz");
  const unknown =
    vert === undefined ||
    isUpright(upright) ||
    Number(attributeOf(bodyPr, "rot") ?? "0") !== 0 ||
    frame.flipH ||
    frame.flipV;
  // `vert` turns the glyphs but leaves the written width and height of the box as they are (a `txXfrm rot`, the hexagons', has them swapped
  // by Office already): with a quarter `vert` the lines run along the written HEIGHT, so the box is laid with its sides swapped.
  return {
    textFrame,
    textTurn: unknown ? undefined : folded((writtenTextFrame?.rot ?? 0) + vert),
    swapped: !unknown && vert !== 0 && vert % (2 * QUARTER_TURN) !== 0,
  };
};

const boxOf = (frame: Frame): Box => ({ x: frame.x, y: frame.y, cx: frame.cx, cy: frame.cy });

/** The guides a traced preset reads: a hexagon's `vf`, a chord's `adj1` and `adj2` (those written). */
const presetGuides = (guides: Readonly<Record<string, number>>): { vf?: number; adj1?: number; adj2?: number } =>
  Object.fromEntries(
    (["vf", "adj1", "adj2"] as const).filter((name) => guides[name] !== undefined).map((name) => [name, guides[name]]),
  );

/** One `dsp:sp`, in the drawing's own space (EMU, origin the frame's corner). */
const readShape = (sp: XNode, rank: number, context: ColorContext | undefined): PptxTextShape => {
  const spPr = findChild(sp, "dsp:spPr");
  const style = findChild(sp, "dsp:style");
  const body = findChild(sp, "dsp:txBody");
  const frame = frameOf(findChild(spPr ?? {}, "a:xfrm"));
  const geometry = readPrstGeom(spPr);
  const guides = readAvLstGuides(spPr);
  const fontColor = fontColorOf(style, context);
  const colors: RunColors = { ...(context === undefined ? {} : { context }), inherited: () => fontColor };
  const paragraphs: PptxParagraph[] =
    body === undefined
      ? []
      : findChildren(body, "a:p").map((paragraph) => readParagraph(paragraph, NO_RELATIONSHIPS, colors));
  // `readBodyPr` marks every `vert` as a turned text (right for a slide); `textPlacement` knows the ones that are only a turn.
  const bodyRead = readBodyPr(body ?? {});
  delete bodyRead.isTextTurned;
  const common = {
    kind: "text" as const,
    spTreeIndex: rank,
    paragraphs,
    ...(geometry === undefined ? {} : { geometry }),
    ...(guides["adj"] === undefined ? {} : { adj: guides["adj"] }),
    fill: fillOf(spPr, style, context),
    hasVisibleLine: lineShownOf(spPr, style),
    ...bodyRead,
  };
  if (frame === undefined) return common;
  const { textFrame, textTurn, swapped } = textPlacement(sp, frame, body);
  return {
    ...common,
    // The bounds of the turned frame (the picture grows by them); the outline itself is the frame as written, under `rotate`.
    box: computeVisibleBox(frame, IDENTITY_MATRIX),
    ...(frame.flipH ? { flipH: true as const } : {}),
    ...(frame.flipV ? { flipV: true as const } : {}),
    // A text whose rule is not known: the outline is kept, the words are left out (`turned-text`).
    ...(textTurn === undefined ? { isTextTurned: true as const } : {}),
    smartArt: {
      frame: boxOf(frame),
      turn: folded(frame.rot),
      textBox: boxOf(textFrame),
      textTurn: textTurn ?? 0,
      ...(swapped ? { textSwap: true as const } : {}),
      ...presetGuides(guides),
    },
  };
};

/** What reading a drawing gives: its shapes (document order), or "too-many" past `ReaderLimits.slideElements` shapes in ONE drawing. */
export type DrawingRead = { shapes: PptxTextShape[] } | "too-many";

/** The shapes of a `dsp:drawing` root, in the drawing's own space. `context`: the colours of the part that holds the frame. */
export const readSmartArtDrawing = (
  root: XNode,
  context: ColorContext | undefined,
  limits?: ReaderLimits,
): DrawingRead => {
  const shapeNodes = findDescendants(root, "dsp:sp");
  if (shapeNodes.length > (limits?.slideElements ?? NO_CAP)) return "too-many";
  // A saved drawing writes its colours with modifiers (`tint`, `lumMod`…): they are resolved here, for a SmartArt only
  // (`ColorContext.modifiers`). A slide's own shapes keep a modified colour unknown (MIKSUDOC « Spécification fonctionnelle » §9.3.3).
  const resolving: ColorContext = { ...(context ?? { clrMap: {}, clrScheme: {}, fillStyleLst: [] }), modifiers: true };
  return { shapes: shapeNodes.map((sp, rank) => readShape(sp, rank, resolving)) };
};

const movedBox = (box: Box, dx: number, dy: number): Box => ({ x: box.x + dx, y: box.y + dy, cx: box.cx, cy: box.cy });

/** The shapes of a drawing carried to the slide: the drawing's origin is the frame's corner, scale 1. */
export const placeShapes = (shapes: readonly PptxTextShape[], origin: { x: number; y: number }): PptxTextShape[] =>
  shapes.map((shape) => ({
    ...shape,
    ...(shape.box === undefined ? {} : { box: movedBox(shape.box, origin.x, origin.y) }),
    ...(shape.smartArt === undefined
      ? {}
      : {
          smartArt: {
            ...shape.smartArt,
            frame: movedBox(shape.smartArt.frame, origin.x, origin.y),
            textBox: movedBox(shape.smartArt.textBox, origin.x, origin.y),
          },
        }),
  }));

/** The colours a Word document's SmartArt text and fills resolve through: the document's theme (`word/theme/theme1.xml`) with the
 *  default mapping (`bg1` is `lt1`…). The document's own `w:clrSchemeMapping` is not read: a SmartArt drawing writes `lt1`/`accent1`
 *  directly (all 4 of smart-art.docx), the mapped names only matter to a colour written through them. */
export const readThemeColorContext = (theme: XNode | undefined): ColorContext | undefined => {
  const elementsOfTheme = findChild(theme ?? {}, "a:themeElements");
  if (elementsOfTheme === undefined) return undefined;
  const formatScheme = findChild(elementsOfTheme, "a:fmtScheme");
  return {
    clrMap: { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" },
    clrScheme: readClrScheme(findChild(elementsOfTheme, "a:clrScheme")),
    fillStyleLst: elementChildrenOf(findChild(formatScheme ?? {}, "a:fillStyleLst") ?? {}),
  };
};

/** The words of the data model that the picture does not carry: the model's words (each counted) less the drawn lines' words, one for
 *  one. A shape the drawer leaves out, a stale drawing (the model grew after the last save) and a drawing kept out leave their words
 *  here: they are counted `smartart-not-drawn` (`pipeline/declared-format-conservation.ts`); the drawn words are excused. Words are
 *  `lib/words.ts#wordsOf`, the rule the control shares. */
export const findUndrawnWords = (modelPieces: readonly string[], drawnLines: readonly string[]): string[] => {
  const drawn = new Map<string, number>();
  for (const word of drawnLines.flatMap(wordsOf)) drawn.set(word, (drawn.get(word) ?? 0) + 1);
  const left: string[] = [];
  for (const word of modelPieces.flatMap(wordsOf)) {
    const remaining = drawn.get(word) ?? 0;
    if (remaining > 0) drawn.set(word, remaining - 1);
    else left.push(word);
  }
  return left;
};
