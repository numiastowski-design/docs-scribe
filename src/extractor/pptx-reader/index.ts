/** A .pptx read by our own reader, as ECMA-376 declares it: its VISIBLE slides
 *  in `p:sldIdLst` order, each with its declared title and its other visible
 *  shapes (text, tables, charts and their data, other graphic frames,
 *  pictures, connectors), and the slide size. Speaker notes are never read:
 *  the page does not publish them (MIKSUDOC « Spécification fonctionnelle »). An
 *  intermediate model, not yet an IR document.
 *  NOT here: the package layer (`ooxml/package.ts`), the XML tree
 *  (`ooxml/xml.ts`), what a shape holds (`pptx-reader/shapes.ts`), a
 *  chart part (`pptx-reader/chart.ts`), nor what the page publishes
 *  (`extractor/pptx.ts`).
 *  Called by: `extractor/pptx.ts`; its types by `extractor/slide-svg.ts` and `extractor/pptx-diagram-order.ts`. */

import { ReaderError } from "doc-ir";
import type { Size } from "../../layout/slide-elements";
import { type ReaderLimits } from "doc-ir";
import { NO_CAP } from "doc-ir/internal";
import {
  type OpcPackage,
  IMAGE_MAX_TOTAL_BYTES,
  openPackage,
  readUnlessDamaged,
  PartDamagedError,
  type Relationship,
  findRelationshipOfKind,
} from "../ooxml/package";
import { isXsdFalse } from "../../lib/ooxml-bool";
import { attributeOf, findChild, findChildren, qualifiedNameOf, type XNode } from "../ooxml/xml";
import { readCachesWithoutModel, readSmartArt } from "../smartart";
import { placeShapes, readSmartArtDrawing } from "../smartart-drawing";
import { readChartOrWithheld } from "./chart";
import { type ColorContext, readClrMap, elementChildrenOf, readOverrideClrMapping, readClrScheme } from "./color";
import type { PptxMedia } from "./picture";
import { countElements, type PptxGraphicShape, type PptxShape, type PptxTextShape, readShapes } from "./shapes";
import { readTextFontSize } from "./text";
import type { TextStyleSources } from "./text-color";

export type { PptxPictureNotPublished } from "./picture";
export type { PptxConnectorShape, PptxGraphicShape, PptxPictureShape, PptxShape, PptxTextShape } from "./shapes";
export type { PptxTable } from "./table";
export type { PptxParagraph } from "./text";

export interface PptxSlide {
  /** 1-based rank among the VISIBLE slides, in `p:sldIdLst` order. */
  number: number;
  /** 0-based position in `p:sldIdLst`, hidden slides counted: how a slide is
   *  matched to its Google Slides page (`extractor/pptx.ts#pageIdsByListIndex`). */
  listIndex: number;
  /** Archive path of the slide part. */
  partPath: string;
  /** Every element of the shape tree, counted at `spTreeIndex` rank —
   *  groups, hidden shapes and the ones past `MAX_SLIDE_ELEMENTS`
   *  included. Above the cap, the tab tells it (`extractor/pptx.ts`).
   *  0 for a slide past the display cap (its shapes are not read). */
  elementCount: number;
  /** Hidden shapes (`p:cNvPr hidden`) left out of this slide, a hidden group for one; absent when none or when the slide's shapes are not read. */
  hiddenShapes?: number;
  /** The text of the shape the slide DECLARES as its title: `p:ph` of type
   *  `title` or `ctrTitle` with no `idx` (a `title` WITH an `idx` is a column
   *  header). Its non-blank paragraphs joined by "\n"
   *  (a line break inside one stays "\n"), the whole trimmed. Absent when no shape declares one, or when it is blank: the
   *  title is never guessed from the biggest text. */
  title?: string;
  /** The title is the previous VISIBLE slide's title, same text exactly
   *  (both trimmed): it is not published again (MIKSUDOC « Spécification fonctionnelle » §9.3.1). */
  isTitleRepeated: boolean;
  /** Every visible shape except the title shape, which is removed from
   *  the body even when blank. */
  shapes: PptxShape[];
}

export interface PptxReadResult {
  /** `p:sldSz cx`/`cy`, in EMU; absent when not written as positive integers
   *  within ECMA's ST_PositiveCoordinate bound. */
  slideSize?: Size;
  /** `p:defaultTextStyle`: the size, in points, its `a:lvlNpPr/a:defRPr sz`
   *  writes for each level (0-based: lvl1pPr is 0); a level writing none is
   *  absent. `a:defPPr` is not a level and is not read. Only the slide
   *  drawing uses it, for a shape that is no placeholder
   *  (`extractor/slide-svg.ts`). Absent when no level writes a size. */
  defaultTextSizes?: Readonly<Record<number, number>>;
  slides: PptxSlide[];
  /** How many slides `p:sldIdLst` lists, hidden ones included. */
  listedSlideCount: number;
  /** A theme (or theme override) part is declared but unreadable: default colours were used. Said as `styles-defaulted`. */
  isThemeDamaged?: boolean;
}

/** A shape that declares itself the slide title. When several do, the first in document order is taken; the others stay body. */
const isTitleShape = (shape: PptxShape): shape is PptxTextShape =>
  shape.kind === "text" &&
  shape.placeholder !== undefined &&
  shape.placeholder.idx === undefined &&
  (shape.placeholder.type === "title" || shape.placeholder.type === "ctrTitle");

const titleText = (shape: PptxTextShape): string | undefined => {
  const text = shape.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .filter((paragraphText) => paragraphText.trim() !== "")
    .join("\n")
    .trim();
  return text === "" ? undefined : text;
};

/** The image parts a slide's relationships name, loaded before its
 *  (synchronous) shape walk. `parts` keeps each part read once for the whole
 *  deck: a logo on every slide is one read — a DAMAGED one is one refusal,
 *  remembered, so the same picture on another slide says the same thing.
 *  `budget.left` is the deck-wide `IMAGE_MAX_TOTAL_BYTES` (arbitrage of
 *  28/09/2026, same T9 reasoning as the docx reader's): each read is capped
 *  at what remains, so the part passing it meets the CEILING refusal — a
 *  bomb, never degradable. Chart-part renderings stay under the package
 *  ceilings only: they are not slide media. */
const mediaOf = async (
  pkg: OpcPackage,
  relationships: ReadonlyMap<string, Relationship>,
  parts: Map<string, Uint8Array | "damaged" | undefined>,
  budget: { left: number },
): Promise<PptxMedia> => {
  const media = new Map<string, { bytes: Uint8Array; contentType: string } | "damaged">();
  for (const [id, relationship] of relationships) {
    if (relationship.kind !== "image" || relationship.external) continue;
    if (!parts.has(relationship.target)) {
      try {
        const bytes = await pkg.readBytes(relationship.target, budget.left);
        parts.set(relationship.target, bytes);
        budget.left -= bytes?.length ?? 0;
      } catch (error) {
        // A corrupt image part degrades to the picture's own marker
        // (arbitrage of 28/09/2026); any other failure is ours — the
        // budget's own refusal is NOT PartDamagedError and passes through.
        // Not `readUnlessDamaged`: « damaged » must stay apart from « absent » here.
        if (!(error instanceof PartDamagedError)) throw error;
        parts.set(relationship.target, "damaged");
      }
    }
    const bytes = parts.get(relationship.target);
    if (bytes === "damaged") media.set(id, "damaged");
    else if (bytes)
      media.set(id, { bytes, contentType: pkg.contentType(relationship.target) ?? "application/octet-stream" });
  }
  return media;
};

/** What is left of a slide's own budget of drawing shapes, and the limits the SmartArt drawings are read under. */
interface DrawingBudget {
  left: number;
  readonly limits: ReaderLimits | undefined;
}

/** The saved drawing of a SmartArt frame (`part`: the drawing the frame's data model names), carried to the frame's corner; "not-drawn"
 *  when it exists but is kept out: past `MAX_SLIDE_ELEMENTS` shapes, or past what is left of the slide's own budget of drawing
 *  shapes (`budget`: the SmartArt that crosses it is not drawn, the slide's element count never moves), or a frame a group scales, turns
 *  or mirrors. `undefined`: nothing to draw (no part, damaged, no shape), the words then stay in the diagram text alone. */
const drawingOf = async (
  pkg: OpcPackage,
  shape: PptxGraphicShape,
  part: string | undefined,
  context: ColorContext | undefined,
  budget: DrawingBudget,
): Promise<PptxGraphicShape["drawing"]> => {
  const root = part === undefined ? undefined : await readUnlessDamaged(() => pkg.readXml(part));
  if (root === undefined) return undefined;
  const read = readSmartArtDrawing(root, context, budget.limits);
  if (read === "too-many") return "not-drawn";
  if (read.shapes.length === 0) return undefined;
  if (shape.isScaledOrTurnedByGroup === true || shape.box === undefined || read.shapes.length > budget.left)
    return "not-drawn";
  budget.left -= read.shapes.length;
  return { shapes: placeShapes(read.shapes, shape.box) };
};

/** A SmartArt frame with its words and its saved drawing. `claimed`: the drawing part its data model names; `modelless`: no data model
 *  part could be read (a damaged one degrades to the frame's own marker, like a chart part). */
const readSmartArtFrame = async (
  pkg: OpcPackage,
  shape: PptxGraphicShape,
  slideRelationships: ReadonlyMap<string, Relationship>,
  context: ColorContext | undefined,
  budget: DrawingBudget,
): Promise<{ shape: PptxGraphicShape; claimed: string | undefined; modelless: boolean }> => {
  const read = await readUnlessDamaged(() => readSmartArt(pkg, slideRelationships, shape.dataPart));
  const drawing =
    read === undefined ? undefined : await drawingOf(pkg, shape, read.claimedDrawingPart, context, budget);
  const words = read !== undefined && read.words.length > 0 ? { words: read.words } : {};
  return {
    shape: { ...shape, ...words, ...(drawing === undefined ? {} : { drawing }) },
    claimed: read?.claimedDrawingPart,
    modelless: read === undefined,
  };
};

/** Each chart frame with its part's data (`pptx-reader/chart.ts`), each SmartArt frame with its words and its saved drawing. */
const withCharts = async (
  pkg: OpcPackage,
  shapes: PptxShape[],
  slideRelationships: ReadonlyMap<string, Relationship>,
  context: ColorContext | undefined,
  limits: ReaderLimits | undefined,
): Promise<PptxShape[]> => {
  const drawingBudget: DrawingBudget = { left: limits?.slideElements ?? NO_CAP, limits };
  const out: PptxShape[] = [];
  // The drawing caches the slide's data models name, and the frames with no model.
  const claimed = new Set<string>();
  const modelless: number[] = [];
  for (const shape of shapes) {
    if (shape.kind === "graphic" && shape.graphic === "smartart") {
      const frame = await readSmartArtFrame(pkg, shape, slideRelationships, context, drawingBudget);
      if (frame.claimed !== undefined) claimed.add(frame.claimed);
      if (frame.modelless) modelless.push(out.length);
      out.push(frame.shape);
    } else if (shape.kind === "graphic" && shape.graphic === "chart") {
      out.push(await withChart(pkg, shape, limits));
    } else {
      out.push(shape);
    }
  }
  if (modelless.length > 0) await addModellessWords(pkg, out, slideRelationships, claimed, modelless);
  return out;
};

/** A chart frame with its part's data. A frame whose relationship does not resolve has no part: said like a missing one. A damaged
 *  chart part is a chart kept out whole, like a bound (docx and xlsx say the same). */
const withChart = async (
  pkg: OpcPackage,
  shape: PptxGraphicShape,
  limits: ReaderLimits | undefined,
): Promise<PptxShape> => {
  if (shape.part === undefined) return { ...shape, chart: "withheld" };
  const chart = await readChartOrWithheld(pkg, shape.part, { limits });
  return chart === undefined ? shape : { ...shape, chart };
};

/** The SmartArts with no data model part (`modelless`: their positions in `out`) take the caches no model claimed
 *  (`extractor/smartart.ts`), in order. */
const addModellessWords = async (
  pkg: OpcPackage,
  out: PptxShape[],
  slideRelationships: ReadonlyMap<string, Relationship>,
  claimed: ReadonlySet<string>,
  modelless: readonly number[],
): Promise<void> => {
  const words = await readCachesWithoutModel(pkg, slideRelationships, claimed, modelless.length);
  modelless.forEach((at, modellessIndex) => {
    const shape = out[at];
    const mine = words[modellessIndex] ?? [];
    if (shape && mine.length > 0) out[at] = { ...shape, words: mine } as PptxShape;
  });
};

/** The first relationship of `kind` a part names, inside the package. */
const targetOf = (relationships: ReadonlyMap<string, Relationship>, kind: string): string | undefined =>
  findRelationshipOfKind(relationships, kind)?.target;

/** What a slide's `a:schemeClr` resolves through (`pptx-reader/color.ts`),
 *  and its layout and master (what its text inherits a colour from,
 *  `pptx-reader/text-color.ts`):
 *  its colour map, the slide's `p:clrMapOvr` first, then its layout's,
 *  then the master's `p:clrMap`; the master's theme, or the slide's
 *  `themeOverride` for what the override writes. The context is
 *  `undefined` when the slide names no layout, master or theme. `parts`
 *  keeps each part read once for the whole deck. */
const colorContextOf = async (
  pkg: OpcPackage,
  slide: XNode,
  slideRelationships: ReadonlyMap<string, Relationship>,
  parts: Map<string, XNode | undefined>,
  themeDamaged: { value: boolean },
): Promise<{ context?: ColorContext; layout?: XNode; master?: XNode }> => {
  const layoutPath = targetOf(slideRelationships, "slideLayout");
  const layout = await readCachedPart(pkg, parts, layoutPath);
  const masterPath =
    layoutPath === undefined ? undefined : targetOf(await pkg.relationships(layoutPath), "slideMaster");
  const master = await readCachedPart(pkg, parts, masterPath);
  const themePath = masterPath === undefined ? undefined : targetOf(await pkg.relationships(masterPath), "theme");
  const theme = findChild((await readCachedTheme(pkg, parts, themePath, themeDamaged)) ?? {}, "a:themeElements");
  const override = await readCachedTheme(pkg, parts, targetOf(slideRelationships, "themeOverride"), themeDamaged);
  const masterMap = findChild(master ?? {}, "p:clrMap");
  const parents = { ...(layout === undefined ? {} : { layout }), ...(master === undefined ? {} : { master }) };
  if (theme === undefined || masterMap === undefined) return parents;
  return { ...parents, context: colorContextFrom({ slide, layout, masterMap, theme, override }) };
};

/** A layout or master part, read once for the whole deck (`parts`); strict: a damaged one throws `PartDamagedError`, which no
 *  caller in this reader catches, so the deck is refused — it carries the slide's structure, not only its colours. A missing one
 *  is `undefined`. */
const readCachedPart = async (
  pkg: OpcPackage,
  parts: Map<string, XNode | undefined>,
  path: string | undefined,
): Promise<XNode | undefined> => {
  if (path === undefined) return undefined;
  if (!parts.has(path)) parts.set(path, await pkg.readXml(path));
  return parts.get(path);
};

/** A theme part (or theme override), read once for the whole deck (`parts`). A damaged one degrades to the same fallback as a
 *  missing one — default colours, cached so the next slide does not re-parse it (arbitrage of 28/09/2026) — and sets `themeDamaged`. */
const readCachedTheme = async (
  pkg: OpcPackage,
  parts: Map<string, XNode | undefined>,
  path: string | undefined,
  themeDamaged: { value: boolean },
): Promise<XNode | undefined> => {
  if (path === undefined) return undefined;
  if (!parts.has(path)) {
    const read = await readUnlessDamaged(() => pkg.readXml(path));
    if (read === undefined && pkg.hasPart(path)) themeDamaged.value = true;
    parts.set(path, read);
  }
  return parts.get(path);
};

/** The colour context of a slide whose master writes a `p:clrMap` and whose theme is read: the slide's `p:clrMapOvr` first, then its
 *  layout's, then the master's map; the override's colour and format schemes before the theme's. */
const colorContextFrom = (colorSources: {
  slide: XNode;
  layout: XNode | undefined;
  masterMap: XNode;
  theme: XNode;
  override: XNode | undefined;
}): ColorContext => {
  const { slide, layout, masterMap, theme, override } = colorSources;
  const clrScheme = findChild(override ?? {}, "a:clrScheme") ?? findChild(theme, "a:clrScheme");
  const fmtScheme = findChild(override ?? {}, "a:fmtScheme") ?? findChild(theme, "a:fmtScheme");
  return {
    clrMap: readOverrideClrMapping(slide) ?? readOverrideClrMapping(layout) ?? readClrMap(masterMap),
    clrScheme: readClrScheme(clrScheme),
    fillStyleLst: elementChildrenOf(findChild(fmtScheme ?? {}, "a:fillStyleLst") ?? {}),
  };
};

/** The one message for a damaged deck, never two copies to diverge. */
const damaged = (): ReaderError =>
  new ReaderError("EXTRACT_FAILED", "the file is damaged or is not a valid PowerPoint presentation");

export interface ReadPptxOptions {
  /** Shapes, media and charts are read only for the first this many VISIBLE
   *  slides; the slides past it are still read and counted (their XML
   *  fetched and validated, hidden ones skipped, `listedSlideCount` whole)
   *  but carry no shapes and no title. `extractor/pptx.ts` passes its
   *  display cap: a slide it drops never pays the inflation of its media
   *  parts nor the analysis of its chart parts. Absent: every slide. */
  detailCap?: number;
  /** The caps this reading applies (`lib/reader-limits.ts#ReaderLimits`): slide elements, the media byte allowance,
   *  the chart bounds. Absent: NONE — `NO_CAP`, nothing withheld (the app passes
   *  `core/document-caps.ts#PRODUCT_READER_LIMITS`). */
  limits?: ReaderLimits | undefined;
}

/** `p:sldSz` as a size, only when both sides are positive integers within ECMA's bound. */
const readSlideSize = (presentation: XNode): Size | undefined => {
  const sldSz = findChild(presentation, "p:sldSz");
  const cx = Number(attributeOf(sldSz ?? {}, "cx"));
  const cy = Number(attributeOf(sldSz ?? {}, "cy"));
  // ECMA-376 ST_PositiveCoordinate caps at 27 273 042 316 900 EMU; a size
  // past it ("1e300" is an integer to `Number.isInteger`) is treated as not
  // written — no geometry is invented, the slides come out in file order.
  const withinBounds = (side: number): boolean => Number.isInteger(side) && side > 0 && side <= 27_273_042_316_900;
  return sldSz && withinBounds(cx) && withinBounds(cy) ? { cx, cy } : undefined;
};

/** What every slide of one deck reads against: the package, the deck's own parts and the caches shared by its slides (each part read once). */
interface DeckReading {
  pkg: OpcPackage;
  defaultTextStyle: XNode | undefined;
  slideSize: Size | undefined;
  detailCap: number | undefined;
  limits: ReaderLimits | undefined;
  parts: Map<string, Uint8Array | "damaged" | undefined>;
  mediaBudget: { left: number };
  xmlParts: Map<string, XNode | undefined>;
  themeDamaged: { value: boolean };
  /** `p:presentation firstSlideNum` (default 1): the number of the first listed slide. */
  firstSlideNum: number;
}

/** The slide part a `p:sldId` names, parsed; a slide the list names but the package lacks (or that is no `p:sld`) is damage. */
const openSlidePart = async (
  pkg: OpcPackage,
  relationships: ReadonlyMap<string, Relationship>,
  sldId: XNode,
): Promise<{ target: string; slide: XNode }> => {
  const relId = attributeOf(sldId, "r:id");
  const relationship = relId === undefined ? undefined : relationships.get(relId);
  if (!relationship || relationship.external || relationship.kind !== "slide") throw damaged();
  const slide = await pkg.readXml(relationship.target);
  if (!slide || qualifiedNameOf(slide) !== "p:sld") throw damaged();
  return { target: relationship.target, slide };
};

/** A slide counted but not read: no shapes, no title. */
const unreadSlide = (
  slides: readonly PptxSlide[],
  listIndex: number,
  partPath: string,
  elementCount: number,
): PptxSlide => ({
  number: slides.length + 1,
  listIndex,
  partPath,
  elementCount,
  isTitleRepeated: false,
  shapes: [],
});

/** The shapes of a slide's tree with their charts and SmartArts read: its colour context, media, then the walk. */
const readSlideShapes = async (
  deck: DeckReading,
  slide: XNode,
  tree: XNode | undefined,
  slideRelationships: ReadonlyMap<string, Relationship>,
  slideNumber: number,
): Promise<{ read: { shapes: PptxShape[]; elementCount: number; hiddenShapes: number }; all: PptxShape[] }> => {
  const { pkg, defaultTextStyle } = deck;
  const { context, layout, master } = tree
    ? await colorContextOf(pkg, slide, slideRelationships, deck.xmlParts, deck.themeDamaged)
    : {};
  const sources: TextStyleSources = {
    ...(layout === undefined ? {} : { layout }),
    ...(master === undefined ? {} : { master }),
    ...(defaultTextStyle === undefined ? {} : { defaultTextStyle }),
  };
  const read = tree
    ? readShapes(tree, {
        relationships: slideRelationships,
        media: await mediaOf(pkg, slideRelationships, deck.parts, deck.mediaBudget),
        context,
        sources,
        slideSize: deck.slideSize,
        slideNumber,
        limits: deck.limits,
      })
    : { shapes: [], elementCount: 0, hiddenShapes: 0 };
  const all = await withCharts(pkg, read.shapes, slideRelationships, context, deck.limits);
  return { read, all };
};

/** One `p:sldId` read; `undefined` for a hidden slide, which produces nothing. */
const readSlide = async (
  deck: DeckReading,
  relationships: ReadonlyMap<string, Relationship>,
  slides: readonly PptxSlide[],
  listIndex: number,
  sldId: XNode,
): Promise<PptxSlide | undefined> => {
  const { pkg, detailCap } = deck;
  const { target, slide } = await openSlidePart(pkg, relationships, sldId);
  // ECMA-376: `show="0"` hides the slide; it produces nothing.
  if (isXsdFalse(attributeOf(slide, "show"))) return undefined;
  // Past the detail cap the slide is counted but its shapes are not read:
  // no media part is inflated, no chart part is opened (a broken image on
  // a slide the caller drops does not fail the deck).
  if (detailCap !== undefined && slides.length >= detailCap) return unreadSlide(slides, listIndex, target, 0);
  const tree = findChild(findChild(slide, "p:cSld") ?? {}, "p:spTree");
  // The element-cap withhold is DECIDED on a structural count BEFORE
  // anything else touches the slide, so that a withheld slide spends
  // nothing of the shared media budget. A withheld slide pays no colour context, no
  // media, no text, no chart — and reads NO TITLE either: the next
  // slide's `isTitleRepeated` suppression cannot apply after it (chosen
  // bias, written — the slide is not published, its declared title is
  // not read).
  const structuralCount = tree ? countElements(tree) : 0;
  if (structuralCount > (deck.limits?.slideElements ?? NO_CAP))
    return unreadSlide(slides, listIndex, target, structuralCount);
  const slideRelationships = await pkg.relationships(target);
  const { read, all } = await readSlideShapes(deck, slide, tree, slideRelationships, deck.firstSlideNum + listIndex);
  const titleShape = all.find(isTitleShape);
  const title = titleShape === undefined ? undefined : titleText(titleShape);
  const previous = slides.at(-1)?.title;
  return {
    number: slides.length + 1,
    listIndex,
    partPath: target,
    elementCount: read.elementCount,
    ...(read.hiddenShapes > 0 ? { hiddenShapes: read.hiddenShapes } : {}),
    ...(title === undefined ? {} : { title }),
    isTitleRepeated: title !== undefined && title === previous,
    // eslint-disable-next-line sonarjs/different-types-comparison -- identity against the title shape found in this same array (a `PptxTextShape` is a `PptxShape`).
    shapes: all.filter((shape) => shape !== titleShape),
  };
};

/** `p:defaultTextStyle`'s size for each of the nine levels that writes one. */
const defaultTextSizesOf = (defaultTextStyle: XNode | undefined): Record<number, number> => {
  const sizes: Record<number, number> = {};
  for (let level = 0; level < 9; level++) {
    const size = readTextFontSize(
      findChild(findChild(defaultTextStyle ?? {}, `a:lvl${level + 1}pPr`) ?? {}, "a:defRPr"),
    );
    if (size !== undefined) sizes[level] = size;
  }
  return sizes;
};

/** ECMA-376 Part 1 §19.3.1.26: `firstSlideNum` is an integer, 1 when absent; an unreadable value counts as absent. */
const firstSlideNumOf = (presentation: XNode): number => {
  const firstSlideNum = Number(attributeOf(presentation, "firstSlideNum"));
  return Number.isInteger(firstSlideNum) && firstSlideNum >= 0 && firstSlideNum < 10 ** 9 ? firstSlideNum : 1;
};

export const readPptx = async (bytes: Uint8Array, options: ReadPptxOptions = {}): Promise<PptxReadResult> => {
  const pkg = await openPackage(bytes, damaged);
  const presentation = await pkg.readXml(pkg.mainPart);
  if (!presentation || qualifiedNameOf(presentation) !== "p:presentation") throw damaged();
  const relationships = await pkg.relationships(pkg.mainPart);

  // The slide order is the order of `p:sldIdLst`, never the part names.
  // A slide the list names but the package lacks is damage, not a slide to
  // skip quietly.
  const slides: PptxSlide[] = [];
  const defaultTextStyle = findChild(presentation, "p:defaultTextStyle");
  const slideSize = readSlideSize(presentation);
  const deck: DeckReading = {
    pkg,
    defaultTextStyle,
    slideSize,
    detailCap: options.detailCap,
    limits: options.limits,
    parts: new Map(),
    mediaBudget: { left: options.limits?.payloadBytes ?? IMAGE_MAX_TOTAL_BYTES },
    xmlParts: new Map(),
    themeDamaged: { value: false },
    firstSlideNum: firstSlideNumOf(presentation),
  };
  const listed = findChildren(findChild(presentation, "p:sldIdLst") ?? {}, "p:sldId");
  for (const [listIndex, sldId] of listed.entries()) {
    const slide = await readSlide(deck, relationships, slides, listIndex, sldId);
    if (slide !== undefined) slides.push(slide);
  }
  const defaultTextSizes = defaultTextSizesOf(defaultTextStyle);
  return {
    ...(slideSize === undefined ? {} : { slideSize }),
    ...(Object.keys(defaultTextSizes).length === 0 ? {} : { defaultTextSizes }),
    slides,
    listedSlideCount: listed.length,
    ...(deck.themeDamaged.value ? { isThemeDamaged: true } : {}),
  };
};
