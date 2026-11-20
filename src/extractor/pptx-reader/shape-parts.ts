/** The reading of a DrawingML shape that PowerPoint (`p:sp`/`p:cxnSp`) and
 *  Word (`wps:wsp`) share: both write the same `a:xfrm`, `a:prstGeom` and
 *  `a:ln`. One definition here, imported by `pptx-reader/shapes.ts` and
 *  `docx-reader/shapes.ts`, so the two readers cannot drift apart (a copy in
 *  the Word reader once never set `isTurnedByGroup`).
 *  NOT here: what differs between the two — the fill (`fillOf`: PowerPoint
 *  also reads `p:style/a:fillRef`, Word does not, on purpose), the element
 *  names, the text, and how a connector's line is told visible.
 *  Called by: `pptx-reader/shapes.ts`, `docx-reader/shapes.ts`,
 *  `extractor/smartart-drawing.ts`, `extractor/pptx-diagram-order.ts`. */

import { ARROW_GEOMETRIES } from "../../layout/slide-layout";
import { attributeOf, findChild, findChildren, type XNode } from "../ooxml/xml";
import {
  computeBentPath,
  type Frame,
  computeLineEnds,
  isTurnOrMirror,
  type Transform,
  computeVisibleBox,
} from "./geometry";
import type { PptxConnectorShape, PptxShape } from "./shapes";

export const readPrstGeom = (spPr: XNode | undefined): string | undefined =>
  attributeOf(findChild(spPr ?? {}, "a:prstGeom") ?? {}, "prst");

/** A line end draws an arrow head: `type` present and not "none". */
export const hasArrowHead = (ln: XNode | undefined, end: "a:headEnd" | "a:tailEnd"): boolean => {
  const type = attributeOf(findChild(ln ?? {}, end) ?? {}, "type");
  return type !== undefined && type !== "none";
};

/** The guides of the preset's `a:avLst` written `val N`, by name. A guide
 *  written as another formula is left out (its preset default applies). */
export const readAvLstGuides = (spPr: XNode | undefined): Record<string, number> => {
  const avLst = findChild(findChild(spPr ?? {}, "a:prstGeom") ?? {}, "a:avLst");
  const guides: Record<string, number> = {};
  for (const gd of findChildren(avLst ?? {}, "a:gd")) {
    const name = attributeOf(gd, "name");
    const fmlaMatch = /^val (-?\d+)$/.exec(attributeOf(gd, "fmla") ?? "");
    if (name !== undefined && fmlaMatch) guides[name] = Number(fmlaMatch[1]);
  }
  return guides;
};

/** What every shape carries: its place in the tree, its box on the page
 *  (groups applied) and its own turn and flips. */
export const shapeBase = (spTreeIndex: number, frame: Frame | undefined, transform: Transform) => ({
  spTreeIndex,
  ...(frame === undefined ? {} : { box: computeVisibleBox(frame, transform) }),
  ...(frame === undefined || frame.rot === 0 ? {} : { rot: frame.rot }),
  ...(frame?.flipH === true ? { flipH: true as const } : {}),
  ...(frame?.flipV === true ? { flipV: true as const } : {}),
});

/** A text shape inside a group that turns or mirrors it: its own box is not
 *  a faithful frame for the words (set on text shapes only). */
export const turnedByGroupField = (transform: Transform): { isTurnedByGroup?: true } =>
  isTurnOrMirror(transform) ? { isTurnedByGroup: true as const } : {};

/** A connector's own fields: arrow heads, geometry, line ends and the bent
 *  path. Whether its line shows (`hasVisibleLine`) is the caller's rule. */
export const connectorBody = (
  spPr: XNode | undefined,
  frame: Frame | undefined,
  transform: Transform,
  hasVisibleLine: boolean,
): Pick<PptxConnectorShape, "startArrow" | "endArrow" | "geometry" | "hasVisibleLine" | "ends" | "path"> => {
  const ln = findChild(spPr ?? {}, "a:ln");
  const geometry = readPrstGeom(spPr);
  const path =
    frame === undefined || geometry === undefined
      ? undefined
      : computeBentPath(frame, transform, geometry, readAvLstGuides(spPr));
  return {
    startArrow: hasArrowHead(ln, "a:headEnd"),
    endArrow: hasArrowHead(ln, "a:tailEnd"),
    ...(geometry === undefined ? {} : { geometry }),
    hasVisibleLine,
    ...(frame === undefined ? {} : { ends: computeLineEnds(frame, transform) }),
    ...(path === undefined ? {} : { path }),
  };
};

/** A text box that is never the label of a connector: turned (its own turn
 *  or its group's), or a block arrow carrying its own text. */
export const isNotLabel = (turned: boolean, geometry: string | undefined): boolean =>
  turned || (geometry !== undefined && ARROW_GEOMETRIES.has(geometry));

/** `isNotLabel` of a shape the pptx reader read. */
export const shapeIsNotLabel = (shape: PptxShape): boolean =>
  shape.rot !== undefined || (shape.kind === "text" && isNotLabel(shape.isTurnedByGroup === true, shape.geometry));
