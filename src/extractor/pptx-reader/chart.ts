/** The data of one chart part (`c:chartSpace`), for a bar-, column- or
 *  line-shaped chart: a `c:cat` category axis and `c:ser` series whose
 *  `c:val` is a number reference. The table has one row per category, one
 *  column per series, preceded by one column per category level when the
 *  categories have several (outermost first); its title is the `c:title` of `c:chart`; the titles of the axes it shows (`c:catAx`, `c:valAx`, `c:dateAx`, `c:serAx`; not a `c:delete`d one): the first `c:catAx`/`c:dateAx` title heads the table's top-left cell when the categories take one column, the others ride along as `axisTitles` (lines under the table). `x` is the category axis, not the screen's horizontal (horizontal bars lay it vertically); `y` a value axis, `z` a `c:serAx` (3D depth); the value axes of a chart with a scatter or bubble group, or with no category or date axis, are told apart by `c:axPos` (`b`/`t` x, `l`/`r` y); its picture
 *  is the image the chart part's relationships name, when the file carries
 *  one (MIKSUDOC « Spécification technique » §5.4: the full data table always, the picture only
 *  when the archive holds a rendering). Values are the numbers AS THE CHART DISPLAYS them (`displayFormatCode`:
 *  `0.8086…` under « 0% » is "81%"; `<c:v>1840</c:v>` stays "1840"), not the stored doubles; category and name caches stay as written.
 *  A scatter or bubble series (`c:xVal`, `c:yVal`, `c:bubbleSize`) is read too: the X values take the left column when no series writes
 *  categories (the title of the horizontal value axis heads it), else a column of their own; series of the same X share one, series of other X
 *  each get their own (X, Y) pair; a bubble's sizes are one more column. A part with no series gives
 *  no data: the slide gets nothing, a `DocumentNotice` in `extractor/pptx.ts`
 *  says so on the dashboard, never on the page.
 *  Also turns that table into an IR block (`chartToIr`).
 *  A chart that carries NO value in any of its references (`c:strRef`, `c:numRef`, `c:multiLvlStrRef`: an address and no
 *  cache) is read from the sheet when the caller gives a `resolver` (only the workbook reader does; docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md):
 *  the points of every series reference come from `xlsx-chart-reference.ts`, enter the same table, bounds and drawing as cached ones, and a
 *  reference the resolver refuses keeps the file's own cache for the WHOLE chart.
 *  NOT here: the graphic frame that points at the part
 *  (`pptx-reader/shapes.ts`), the addresses and the cells (`extractor/xlsx-chart-reference.ts`).
 *  Called by: `pptx-reader/index.ts`, `extractor/xlsx.ts`, `docx-reader/drawing-objects.ts`
 *  (the read and `chartToIr`), and `extractor/pptx.ts` (`chartToIr`); its `PptxChart` type by `pptx-reader/shapes.ts`. */

import { type ReaderLimits } from "doc-ir";
import { MAX_CHART_CELLS, MAX_CHART_LEVELS, MAX_CHART_POINTS, MAX_CHART_SERIES } from "doc-ir/internal";
import type { ChartDrawing, ImageRef, IRBlock, TableData } from "doc-ir";
import type { ChartAxis, ChartDrawingType } from "doc-ir/internal";
import { readUnlessDamaged, type OpcPackage } from "../ooxml/package";
import {
  attributeOf,
  findChild,
  findChildren,
  findDescendants,
  childNodesOf,
  qualifiedNameOf,
  textContentOf,
  type XNode,
} from "../ooxml/xml";
import type { ChartCellPoint, ChartPoints, ChartResolver } from "../xlsx-chart-reference";
import { formatNumberCell } from "../xlsx-number-format";

export interface PptxChart {
  data: TableData;
  title?: string;
  /** The titles of the axes the chart shows (not the deleted ones) that the table's top-left cell does not carry, in document order. */
  axisTitles?: { axis: ChartAxis; text: string }[];
  /** What the chart macro needs to DRAW this chart (type and numbers), when a Forge component draws it faithfully; absent otherwise. */
  drawing?: ChartDrawing;
  /** The chart part's own rendering, when its relationships name an image
   *  inside the package. */
  image?: { bytes: Uint8Array; contentType: string };
}

/** The format code a series' values are DISPLAYED with: the series' data labels' `c:numFmt` when it is not source-linked
 *  (ECMA-376 Part 1 §21.2.2.121 `numFmt`, `sourceLinked="0"`; `c:dLbls` §21.2.2.49), else the `c:formatCode` of its `c:numCache`
 *  (§21.2.2.71). An axis `c:numFmt` formats ticks and is never read here. `undefined`: General. */
const displayFormatCode = (ser: XNode, numCache: XNode | undefined): string | undefined =>
  labelFormatCode(ser) ?? cacheFormatCode(numCache);

/** The code of a series' data labels when they do not follow the source (§21.2.2.121); `undefined` otherwise. */
const labelFormatCode = (ser: XNode): string | undefined => {
  const label = findChild(findChild(ser, "c:dLbls") ?? {}, "c:numFmt");
  const labelCode = label === undefined ? undefined : attributeOf(label, "formatCode");
  const linked = label === undefined ? undefined : attributeOf(label, "sourceLinked");
  return labelCode !== undefined && labelCode !== "" && linked !== "1" && linked !== "true" ? labelCode : undefined;
};

/** The `c:formatCode` of a number cache (§21.2.2.71): what an X value or a bubble size is shown with (data labels format the Y). `undefined`: General. */
const cacheFormatCode = (numCache: XNode | undefined): string | undefined => {
  const cached = textContentOf(findChild(numCache ?? {}, "c:formatCode") ?? {});
  return cached === "" ? undefined : cached;
};

/** A stored number as the chart shows it, through `formatNumberCell`; a cell that is not a plain number, or a code
 *  `formatNumberCell` does not interpret (it answers `undefined`), stays as stored (a value is never lost for lack of a format). */
const displayed = (stored: string, code: string | undefined): string => {
  const trimmed = stored.trim();
  if (trimmed === "") return stored;
  const storedNumber = Number(trimmed);
  if (!Number.isFinite(storedNumber)) return stored;
  // Known limit, shared by the cache and the sheet paths (kept alike on purpose): a date is written on the 1900 calendar even in a 1904 workbook.
  return formatNumberCell(storedNumber, code, false) ?? stored;
};

/** The bounds of a chart: points, category levels, series and cells: the limits given, else the product's. */
const boundsOf = (
  limits: ReaderLimits | undefined,
): { series: number; cells: number; points: number; levels: number } => ({
  series: limits?.chartSeries ?? MAX_CHART_SERIES,
  cells: limits?.chartCells ?? MAX_CHART_CELLS,
  points: limits?.chartPoints ?? MAX_CHART_POINTS,
  levels: limits?.chartLevels ?? MAX_CHART_LEVELS,
});

/** A cache's points by `idx`, `countNode` holding `c:ptCount` and `pointsNode`
 *  the `c:pt` entries: the same node for a plain cache, not for a
 *  `c:multiLvlStrCache`, whose points sit one level down in `c:lvl`.
 *  Sized from the points actually PRESENT, never from the declared count
 *  (a declared count of 2e9 would allocate an array that size); an `idx` at or
 *  past the declared count is dropped. Chosen bias, written: a trailing point
 *  NO cache of the series writes does not pad the table with an empty row —
 *  `tableOf` sizes the rows on the LONGEST cache, so a value whose category
 *  cell is blank (no `c:pt` written for it) keeps its row.
 *  `undefined`: a point really PRESENT (and under the declared count) sits
 *  at or past the point bound — the whole chart is then withheld to its
 *  marker (the whole-chart withhold, arbitrage of 28/09/2026: never a table
 *  that looks complete with rows missing).
 *  A LYING `c:ptCount` with no such point withholds nothing: refusing it
 *  would lose readable points over a broken attribute. */
const pointsFrom = (
  countNode: XNode | undefined,
  pointsNode: XNode | undefined,
  limits: ReaderLimits | undefined,
): string[] | undefined => {
  // `c:ptCount` is OPTIONAL in ECMA-376 (CT_NumData/CT_StrData, 0..1): absent
  // (or unreadable), the points PRESENT stand, under the point bound alone, so
  // that a conforming cache is never thrown away whole. A WRITTEN count — an
  // explicit 0 included — is a declaration: a point at or past it stays out
  // (anti-padding).
  const countText = attributeOf(findChild(countNode ?? {}, "c:ptCount") ?? {}, "val");
  const count = countText === undefined ? undefined : Number.parseInt(countText, 10);
  const declared = count !== undefined && Number.isInteger(count) && count >= 0 ? count : undefined;
  return denseFrom(cachedEntries(pointsNode), declared, limits);
};

/** The `c:pt` entries of a cache, in order; an unreadable `idx` (« abc ») is no point (the control reading ignores it too), and Infinity (400
 *  digits) is past the bound in `denseFrom`. */
function* cachedEntries(pointsNode: XNode | undefined): Generator<{ idx: number; value: string }> {
  for (const pointNode of findChildren(pointsNode ?? {}, "c:pt")) {
    const idx = Number.parseInt(attributeOf(pointNode, "idx") ?? "-1", 10);
    if (!Number.isNaN(idx)) yield { idx, value: textContentOf(findChild(pointNode, "c:v") ?? {}) };
  }
}

/** The dense array of points, cached or read from the sheet: see `pointsFrom` for the rules (the declared count, the bound, the length). */
const denseFrom = (
  entries: Iterable<{ idx: number; value: string }>,
  declared: number | undefined,
  limits: ReaderLimits | undefined,
): string[] | undefined => {
  const kept: { idx: number; value: string }[] = [];
  const pointCap = boundsOf(limits).points;
  let length = 0;
  for (const { idx, value } of entries) {
    // The withhold looks at the points PRESENT before any declared-count
    // filter: a ptCount lying LOW would otherwise drop a real point past the
    // bound in silence — the very loss the whole-chart withhold ends.
    if (idx >= pointCap) return undefined;
    if (idx < 0 || (declared !== undefined && idx >= declared)) continue;
    kept.push({ idx, value });
    if (idx + 1 > length) length = idx + 1;
  }
  const values = new Array<string>(length).fill("");
  for (const { idx, value } of kept) values[idx] = value;
  return values;
};
const cachePoints = (cache: XNode | undefined, limits: ReaderLimits | undefined): string[] | undefined =>
  pointsFrom(cache, cache, limits);

/** The points a chart reads from its sheet instead of a cache, by the reference node they stand for. */
type SheetPoints = ReadonlyMap<XNode, readonly ChartCellPoint[]>;

const REFERENCE_TAGS = ["c:strRef", "c:numRef", "c:multiLvlStrRef"] as const;
const LITERAL_TAGS = ["c:numLit", "c:strLit"] as const;
const SERIES_DATA_TAGS = ["c:tx", "c:cat", "c:val", "c:xVal", "c:yVal", "c:bubbleSize"] as const;

/** The points of a resolved reference, up to and including the first one at or past the point bound: that one is what withholds the chart
 *  (`denseFrom`), and a range reaching far past the bound is never read whole. */
const upToBound = (found: Iterable<ChartCellPoint>, limits: ReaderLimits | undefined): ChartCellPoint[] => {
  const pointCap = boundsOf(limits).points;
  const kept: ChartCellPoint[] = [];
  for (const point of found) {
    kept.push(point);
    if (point.index >= pointCap) break;
  }
  return kept;
};

/** A reference of the chart, parsed: what the resolver answered and whether it names the series (`c:tx`). */
interface PlannedReference {
  ref: XNode;
  found: ChartPoints;
  name: boolean;
}

/** The sheet's points for the series references of a chart; `undefined` when the chart keeps its cache: any reference of the chart carries a
 *  point (the file wrote values), it holds literal data, it has more series than the bound (its cache then withholds it in `tableOf` exactly
 *  as before the sheet was read), there is no series reference to read, or one reference is outside what the resolver reads (another sheet, an
 *  address outside the grammar, a multi-level category: a chart half from the sheet and half from nothing would draw a lie).
 *  `"withheld"`: the references can yield more cells than the bounds allow, and the chart is withheld whole, like a cache of that size
 *  (the whole-chart withhold of `pointsFrom`; the caller says so in its notice). That decision is made by ARITHMETIC on the parsed addresses, before any cell is read: one
 *  reference past the point bound (its extent, our dense length up to its last point, over MAX_CHART_POINTS), or all the references of the
 *  chart together over MAX_CHART_CELLS: an address counted once however many series name it (it is read once), a series name one cell each
 *  (only its first point is read). Points held for one chart: at most MAX_CHART_CELLS, decided before the first cell is read. */
const sheetPointsOf = (
  root: XNode,
  resolver: ChartResolver,
  limits: ReaderLimits | undefined,
): SheetPoints | "withheld" | undefined => {
  if (REFERENCE_TAGS.some((tag) => findDescendants(root, tag).some((ref) => findDescendants(ref, "c:pt").length > 0)))
    return undefined;
  // Literal data (`c:numLit`, `c:strLit`) is values the file wrote, which this reader does not read: the chart keeps its cache whole.
  if (LITERAL_TAGS.some((tag) => findDescendants(root, tag).length > 0)) return undefined;
  const series = findDescendants(root, "c:ser");
  if (series.length > boundsOf(limits).series) return undefined;
  const plan = planOf(series, resolver, limits);
  return plan === undefined || plan === "withheld" ? plan : pointsOfPlan(plan, limits);
};

/** Every series reference parsed, and the cells they can yield counted against the bounds, before any cell is read (`sheetPointsOf`).
 *  `undefined`: a multi-level category or a refused address (the chart keeps its cache); `"withheld"`: a bound passed. */
const planOf = (
  series: readonly XNode[],
  resolver: ChartResolver,
  limits: ReaderLimits | undefined,
): PlannedReference[] | "withheld" | undefined => {
  const bounds = boundsOf(limits);
  // A parse memo by address text: an address repeated by several series is parsed and read once.
  const answers = new Map<string, ChartPoints | undefined>();
  // The addresses already counted as ranges: read once (`pointsOfPlan`), so counted once. A series name is its own one-point copy: it counts
  // one cell every time.
  const counted = new Set<string>();
  const plan: PlannedReference[] = [];
  let cells = 0;
  for (const { ref, name } of series.flatMap(seriesReferences)) {
    if (qualifiedNameOf(ref) === "c:multiLvlStrRef") return undefined;
    const address = textContentOf(findChild(ref, "c:f") ?? {});
    if (!answers.has(address)) answers.set(address, resolver(address));
    const found = answers.get(address);
    if (found === undefined) return undefined;
    if (name) cells += 1;
    else if (!counted.has(address)) {
      counted.add(address);
      cells += found.extent;
    }
    if ((!name && found.extent > bounds.points) || cells > bounds.cells) return "withheld";
    plan.push({ ref, found, name });
  }
  return plan;
};

/** The points of a plan within the bounds: a name its first point, any other reference up to the bound, an address read once. */
const pointsOfPlan = (plan: readonly PlannedReference[], limits: ReaderLimits | undefined): SheetPoints | undefined => {
  const readOnce = new Map<ChartPoints, readonly ChartCellPoint[]>();
  const resolved = new Map<XNode, readonly ChartCellPoint[]>();
  for (const { ref, found, name } of plan) {
    if (name) resolved.set(ref, firstPoint(found));
    else {
      if (!readOnce.has(found)) readOnce.set(found, upToBound(found, limits));
      resolved.set(ref, readOnce.get(found)!);
    }
  }
  return resolved.size === 0 ? undefined : resolved;
};

/** The first point of a reference alone (a series name is its first cell, `seriesName`). */
const firstPoint = (found: Iterable<ChartCellPoint>): ChartCellPoint[] => {
  for (const point of found) return [point];
  return [];
};

/** The reference nodes a series reads (name, categories, values, X, Y, sizes), each with whether it is the series' name. */
const seriesReferences = (ser: XNode): { ref: XNode; name: boolean }[] =>
  SERIES_DATA_TAGS.flatMap((dataTag) =>
    REFERENCE_TAGS.flatMap((referenceTag): { ref: XNode; name: boolean }[] => {
      const ref = findChild(findChild(ser, dataTag) ?? {}, referenceTag);
      return ref === undefined ? [] : [{ ref, name: dataTag === "c:tx" }];
    }),
  );

/** The cells of a text reference (categories, names, string X values): the cache's, or the grid's own text. */
const textCells = (
  ref: XNode,
  cacheTag: "c:strCache" | "c:numCache",
  sheet: SheetPoints | undefined,
  limits: ReaderLimits | undefined,
): string[] | undefined => {
  const fromSheet = sheet?.get(ref);
  return fromSheet === undefined
    ? cachePoints(findChild(ref, cacheTag), limits)
    : denseFrom(
        fromSheet.map((point) => ({ idx: point.index, value: point.text })),
        undefined,
        limits,
      );
};

/** Whether the references under `node` carry a point, in the cache or on the sheet. */
const carriesPoints = (node: XNode | undefined, sheet: SheetPoints | undefined): boolean =>
  sheet === undefined
    ? findDescendants(node ?? {}, "c:pt").length > 0
    : (["c:strRef", "c:numRef"] as const).some((tag) => (sheet.get(findChild(node ?? {}, tag) ?? {}) ?? []).length > 0);

/** The numbers a sheet gives a number reference, as the chart displays them: `labelCode` (the data labels' code, when they do not follow the
 *  source) else the cell's own code. A cell that is not a number gives no point.
 *  TRIED, THEN DROPPED (16/11/2026): one code for the whole reference, the first numeric cell's, so that a series mixing
 *  formats is drawn on one scale. No real workbook without a cache mixes its formats, so the rule guarded nothing we have. */
const sheetNumbers = (
  read: readonly ChartCellPoint[],
  labelCode: string | undefined,
  limits: ReaderLimits | undefined,
): string[] | undefined =>
  denseFrom(
    read
      .filter((point) => point.stored !== undefined)
      .map((point) => ({ idx: point.index, value: displayed(point.stored!, labelCode ?? point.formatCode) })),
    undefined,
    limits,
  );

/** The levels of a `c:multiLvlStrCache`, in the order written (the leaf level first); empty when it writes none.
 *  `undefined`: a level past the level bound, or a level's cache past the point bound: the whole chart is then withheld. */
const multiLevelCachePoints = (
  multiLevelReference: XNode,
  limits: ReaderLimits | undefined,
): string[][] | undefined => {
  const cache = findChild(multiLevelReference, "c:multiLvlStrCache");
  const levels: string[][] = [];
  const levelCap = boundsOf(limits).levels;
  for (const lvl of findChildren(cache ?? {}, "c:lvl")) {
    if (levels.length >= levelCap) return undefined; // one level past the bound: the whole chart is withheld
    const levelPoints = pointsFrom(cache, lvl, limits);
    if (levelPoints === undefined) return undefined;
    levels.push(levelPoints);
  }
  return levels;
};

/** The categories of a series, as LEVELS, the leaf level first
 *  (`c:multiLvlStrRef`, ECMA-376 Part 1 §21.2.2.115: each `c:lvl` is one
 *  level, the first the innermost; an outer level writes a point only where
 *  its group starts, `idx` being the first leaf of the group). A plain
 *  `c:strRef`/`c:numRef` is one level. Every level is kept: an outer label
 *  is the document's own word, and dropping it was a lost word.
 *  `undefined` propagates a cache past the bound: the whole chart is then withheld. */
const referencedLevels = (
  parent: XNode | undefined,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): string[][] | undefined => {
  const multiLevelReference = findChild(parent ?? {}, "c:multiLvlStrRef");
  if (multiLevelReference) {
    const levels = multiLevelCachePoints(multiLevelReference, limits);
    if (levels === undefined || levels.length > 0) return levels;
  }
  const stringReference = findChild(parent ?? {}, "c:strRef");
  if (stringReference) {
    const referencePoints = textCells(stringReference, "c:strCache", sheet, limits);
    return referencePoints === undefined ? undefined : [referencePoints];
  }
  const numberReference = findChild(parent ?? {}, "c:numRef");
  if (numberReference) {
    const referencePoints = textCells(numberReference, "c:numCache", sheet, limits);
    return referencePoints === undefined ? undefined : [referencePoints];
  }
  return [[]];
};

/** Chosen bias, written: a series NAME cache past the bound only loses the
 *  name (an EMPTY header stands in — A4 of docs/CONCEPTION-CARTE-ET-REMARQUES.md,
 *  29/09/2026: no invented « Series N », the series existed and had
 *  no name) — the table shows one cell of it, nothing is amputated. */
const seriesName = (
  ser: XNode,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): string | undefined => {
  const tx = findChild(ser, "c:tx") ?? {};
  const ref = findChild(tx, "c:strRef");
  // A name typed in the chart (`c:tx/c:v`, no cell behind it) is the series' name as well.
  if (!ref) return findChild(tx, "c:v") === undefined ? undefined : textContentOf(findChild(tx, "c:v")!);
  return textCells(ref, "c:strCache", sheet, limits)?.[0];
};

interface TableColumn {
  name: string | undefined;
  values: string[];
  /** The numbers the values stand for, one per value (`null`: not a number). Series of categories only. */
  numbers?: (number | null)[];
}

const GROUPED_NUMBER = /^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
const PLAIN_NUMBER = /^-?\d+(?:\.\d+)?$/;
// One currency sign, before the number (after its minus) or after it, a space allowed between: « $5 », « -$5 », « 5 € ».
const LEADING_CURRENCY = /^(-?)\p{Sc}[   ]?(?!\p{Sc})(.*)$/u;
const TRAILING_CURRENCY = /^([^\p{Sc}]*?)[   ]?\p{Sc}$/u;

const withoutCurrency = (cell: string): string => {
  const leading = LEADING_CURRENCY.exec(cell);
  if (leading !== null) return leading[1]! + leading[2]!;
  return TRAILING_CURRENCY.exec(cell)?.[1] ?? cell;
};

/** The number a displayed cell stands for: the cell itself when it is a plain number, percentage or amount with one currency sign
 *  (« 81% » is 81, « 1,234 » is 1234, « $5 » is 5, accounting
 *  « (5) » is -5), else `null`: a gap in the drawing, as the table shows it. NEVER the stored value of a cell the chart displays as nothing
 *  (format `;;;`), as a word (`[<5]"faible";0`) or with a unit. */
export const parseDisplayedNumber = (cell: string): number | null => {
  const accounting = /^\((.*)\)$/.exec(cell);
  const inner = withoutCurrency(accounting === null ? cell : accounting[1]!);
  // Commas only as thousands separators, three digits at a time; anything else (units, words, a hidden or conditional display) is a gap.
  const body = inner.endsWith("%") ? inner.slice(0, -1) : inner;
  if (!GROUPED_NUMBER.test(body) && !PLAIN_NUMBER.test(body)) return null;
  const shown = Number(body.replaceAll(",", ""));
  return accounting === null ? shown : -Math.abs(shown);
};

const BAR_GROUPS: ReadonlySet<string> = new Set(["c:barChart", "c:bar3DChart"]);
const LINE_GROUPS: ReadonlySet<string> = new Set(["c:lineChart", "c:line3DChart"]);
const PIE_GROUPS: ReadonlySet<string> = new Set(["c:pieChart", "c:pie3DChart"]);

const barType = (group: XNode, grouping: string | undefined): ChartDrawingType | undefined => {
  if (grouping === "percentStacked") return undefined;
  const stacked = grouping === "stacked";
  if (attributeOf(findChild(group, "c:barDir") ?? {}, "val") === "bar") return stacked ? "hStackBar" : "hBar";
  return stacked ? "stackBar" : "bar";
};

/** The component a chart is drawn with, from its ONE plot group: `c:barChart` / `c:bar3DChart` by `c:barDir`
 *  (ECMA-376 Part 1 §21.2.2.17) and `c:grouping` (§21.2.2.76: `stacked` stacks, `percentStacked` has no component), `c:lineChart` / `c:line3DChart` when
 *  `standard`, `c:pieChart` / `c:pie3DChart` (not `c:ofPieChart`: a pie of pie is not a pie), `c:doughnutChart`. 3D groups are drawn flat (a stated limit). Several groups (a combined chart), or any other group (scatter,
 *  bubble, area, radar, stock, surface): `undefined`, the table stands alone. */
const drawingTypeOf = (root: XNode): ChartDrawingType | undefined => {
  const groups = childNodesOf(findChild(findChild(root, "c:chart") ?? {}, "c:plotArea") ?? {}).filter((child) =>
    /^c:\w+Chart$/.test(qualifiedNameOf(child) ?? ""),
  );
  if (groups.length !== 1) return undefined;
  const group = groups[0]!;
  const tag = qualifiedNameOf(group) ?? "";
  const grouping = attributeOf(findChild(group, "c:grouping") ?? {}, "val");
  if (BAR_GROUPS.has(tag)) return barType(group, grouping);
  if (LINE_GROUPS.has(tag)) return grouping === undefined || grouping === "standard" ? "line" : undefined;
  if (PIE_GROUPS.has(tag)) return "pie";
  return tag === "c:doughnutChart" ? "donut" : undefined;
};

/** The `c:val` values of a series that is not a scatter or bubble one, as the chart displays them: from its cache, or from the sheet. `undefined`: a point at or
 *  past the bound. */
const valueCells = (
  ser: XNode,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): string[] | undefined => {
  const numRef = findChild(findChild(ser, "c:val") ?? {}, "c:numRef");
  const fromSheet = numRef === undefined ? undefined : sheet?.get(numRef);
  if (fromSheet !== undefined) return sheetNumbers(fromSheet, labelFormatCode(ser), limits);
  const numCache = findChild(numRef ?? {}, "c:numCache");
  const stored = cachePoints(numCache, limits);
  if (stored === undefined) return undefined;
  const code = displayFormatCode(ser, numCache);
  return stored.map((value) => displayed(value, code));
};

/** A scatter or bubble series: it writes `c:xVal`, `c:yVal` or `c:bubbleSize` where the others write `c:cat` and `c:val`. */
const isXySeries = (ser: XNode): boolean =>
  ["c:xVal", "c:yVal", "c:bubbleSize"].some((tag) => findChild(ser, tag) !== undefined);

/** The values of one `c:xVal`, `c:yVal` or `c:bubbleSize`: a number reference shown with `codeOf` (the format of its cache, ECMA-376 Part 1 §21.2.2.71),
 *  or a string reference as written. Literal data (`c:numLit`, `c:strLit`) is not read (written limit, as for `c:val`). `undefined`: a cache past the bound. */
const xyValues = (
  parent: XNode | undefined,
  codeOf: (numCache: XNode | undefined) => string | undefined,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
  labelCode: string | undefined,
): string[] | undefined => {
  const stringReference = findChild(parent ?? {}, "c:strRef");
  if (stringReference) return textCells(stringReference, "c:strCache", sheet, limits);
  const numRef = findChild(parent ?? {}, "c:numRef");
  const fromSheet = numRef === undefined ? undefined : sheet?.get(numRef);
  if (fromSheet !== undefined) return sheetNumbers(fromSheet, labelCode, limits);
  const numCache = findChild(numRef ?? {}, "c:numCache");
  const stored = cachePoints(numCache, limits);
  if (stored === undefined) return undefined;
  const code = codeOf(numCache);
  return stored.map((value) => displayed(value, code));
};

/** The same cells, as displayed, row by row (a missing tail is blank). */
const sameValues = (left: readonly string[], right: readonly string[]): boolean => {
  for (let index = 0; index < Math.max(left.length, right.length); index++)
    if ((left[index] ?? "") !== (right[index] ?? "")) return false;
  return true;
};
const isBlank = (values: readonly string[]): boolean => values.every((value) => value === "");

/** The drawing of a chart whose rows are the indices `kept` of its columns: one category column, series of categories only (no scatter column),
 *  and ONE series for a pie or a donut. `undefined`: the table stands alone. */
const drawingOf = (
  type: ChartDrawingType | undefined,
  labelColumnCount: number,
  columns: readonly TableColumn[],
  kept: readonly number[],
): ChartDrawing | undefined => {
  if (type === undefined || labelColumnCount !== 1 || columns.length === 0) return undefined;
  if ((type === "pie" || type === "donut") && columns.length !== 1) return undefined;
  const seriesNumbers = columns.map((column) => column.numbers);
  if (seriesNumbers.some((numbers) => numbers === undefined)) return undefined;
  return { type, values: seriesNumbers.map((numbers) => kept.map((index) => numbers![index] ?? null)) };
};

/** Whether the category axis runs from its last category to its first: the `c:orientation` of the `c:scaling` of the plot area's first
 *  `c:catAx` or `c:dateAx` says `maxMin` (ECMA-376 Part 1 §21.2.2.130, §21.2.3.30; absent: `minMax`). A value axis is not looked at: its direction moves no category. */
const categoryAxisReversed = (root: XNode): boolean => {
  const plotArea = findChild(findChild(root, "c:chart") ?? {}, "c:plotArea") ?? {};
  const axis = findChild(plotArea, "c:catAx") ?? findChild(plotArea, "c:dateAx");
  return attributeOf(findChild(findChild(axis ?? {}, "c:scaling") ?? {}, "c:orientation") ?? {}, "val") === "maxMin";
};

/** The drawing with the direction of its category axis. A pie and a donut have no axis: never reversed. */
const withCategoryDirection = (drawing: ChartDrawing | undefined, root: XNode): ChartDrawing | undefined => {
  if (drawing === undefined || drawing.type === "pie" || drawing.type === "donut" || !categoryAxisReversed(root))
    return drawing;
  return { ...drawing, categoriesReversed: true };
};

/** The indices at which any label or value column holds a non-empty cell. */
const occupiedIndices = (labelColumns: readonly string[][], columns: readonly TableColumn[]): Set<number> => {
  const occupied = new Set<number>();
  for (const column of [...labelColumns, ...columns.map((valueColumn) => valueColumn.values)]) {
    for (let index = 0; index < column.length; index++) if (column[index] !== "") occupied.add(index);
  }
  return occupied;
};

/** The table rows at the occupied indices, in index order, and the index each kept row came from. */
const occupiedRows = (
  occupied: ReadonlySet<number>,
  labelColumns: readonly string[][],
  columns: readonly TableColumn[],
): { rows: string[][]; keptIndices: number[] } => {
  const rows: string[][] = [];
  const keptIndices: number[] = [];
  // Rows come from the SORTED occupied indices only, never 0..rowCount.
  for (const index of [...occupied].sort((left, right) => left - right)) {
    const row = [
      ...labelColumns.map((labelColumn) => labelColumn[index] ?? ""),
      ...columns.map((column) => column.values[index] ?? ""),
    ];
    if (row.some((cell) => cell !== "")) {
      rows.push(row);
      keptIndices.push(index);
    }
  }
  return { rows, keptIndices };
};

/** A chart's table as read, before its titles. `labelColumns`: how many leading columns hold category levels (or a scatter's X values).
 *  `labelFromX`: the left column holds the X values of a scatter (no series writes categories), so the title of the X axis heads it. */
interface ChartTable {
  data: TableData;
  labelColumns: number;
  labelFromX: boolean;
  drawing?: ChartDrawing;
}

/** The table of a chart while its series are read, one series after the other. */
interface TableInProgress {
  /** Category levels, the leaf level first (a plain chart has one). */
  levels: string[][];
  labelFromX: boolean;
  readonly columns: TableColumn[];
  /** The X columns already published beside the left column: a scatter series whose X (as displayed) is one of them shares it. Chosen
   *  rule, written: sharing keeps one row per point; series of other X each get their own (X, Y) pair, so no point shares a row with a
   *  point of another X and no number is lost. An X column that is entirely blank (a cache the file never filled) is not published. */
  readonly publishedX: string[][];
  /** The length of the longest array built so far: the running bound of `tableOf` multiplies by it. */
  longest: number;
}

const hasNoLevels = (table: TableInProgress): boolean => table.levels.every((level) => level.length === 0);

/** The X values of a scatter series are no column of their own when they repeat the left column or an X column already published. */
const isNewXColumn = (xValues: readonly string[], table: TableInProgress): boolean =>
  !(table.levels.length === 1 && sameValues(xValues, table.levels[0]!)) &&
  !table.publishedX.some((published) => sameValues(xValues, published));

/** The columns a scatter or bubble series adds: its X values (the left column when no series writes categories, else a column unless
 *  shared), its Y values under its name, its bubble sizes. `categorised`: some series writes categories. `"withheld"`: a cache past the bound. */
const xySeriesColumns = (
  ser: XNode,
  table: TableInProgress,
  categorised: boolean,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): TableColumn[] | "withheld" => {
  const xValues = xyValues(findChild(ser, "c:xVal"), cacheFormatCode, limits, sheet, undefined);
  const yValues = xyValues(
    findChild(ser, "c:yVal"),
    (cache) => displayFormatCode(ser, cache),
    limits,
    sheet,
    labelFormatCode(ser),
  );
  const sizes = xyValues(findChild(ser, "c:bubbleSize"), cacheFormatCode, limits, sheet, undefined);
  if (xValues === undefined || yValues === undefined || sizes === undefined) return "withheld";
  const added: TableColumn[] = [];
  if (!isBlank(xValues)) {
    if (!categorised && hasNoLevels(table)) {
      table.levels = [xValues];
      table.labelFromX = true;
    } else if (isNewXColumn(xValues, table)) {
      added.push({ name: undefined, values: xValues });
      table.publishedX.push(xValues);
    }
  }
  added.push({ name: seriesName(ser, limits, sheet), values: yValues });
  if (!isBlank(sizes)) added.push({ name: undefined, values: sizes });
  for (const column of [...table.levels, ...added.map((addedColumn) => addedColumn.values)])
    table.longest = Math.max(table.longest, column.length);
  return added;
};

/** The column a series adds when it is not a scatter or bubble one (bar, line, pie, area, radar, stock, surface…: it reads `c:cat` and
 *  `c:val`); the first such series to find no levels yet gives the table its categories.
 *  `"withheld"`: a cache past the bound. */
const categorySeriesColumns = (
  ser: XNode,
  table: TableInProgress,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): TableColumn[] | "withheld" => {
  if (hasNoLevels(table)) {
    const categoryLevels = referencedLevels(findChild(ser, "c:cat"), limits, sheet);
    if (categoryLevels === undefined) return "withheld"; // a cache past the bound: the whole chart is withheld
    table.levels = categoryLevels;
    for (const level of table.levels) table.longest = Math.max(table.longest, level.length);
  }
  const values = valueCells(ser, limits, sheet);
  if (values === undefined) return "withheld";
  table.longest = Math.max(table.longest, values.length);
  return [{ name: seriesName(ser, limits, sheet), values, numbers: values.map(parseDisplayedNumber) }];
};

/** The rows of a table whose columns are all read, and its drawing. Only OCCUPIED rows come out: the 0..maxIdx span would let ONE point
 *  at idx 9 999 fabricate 10 000 near-empty rows. A written value past the last written category still keeps its row, label cell empty
 *  (never a number lost); a row whose label and every value are blank carries nothing and is not fabricated (chosen bias, written). */
const finishedTable = (
  root: XNode,
  table: TableInProgress,
  bounds: ReturnType<typeof boundsOf>,
): ChartTable | "withheld" => {
  const { columns } = table;
  // One leading column per category level, the OUTERMOST first: an outer
  // label stands on the row where its group starts, blank after (never
  // repeated, never joined to the leaf: no word is invented).
  const labelColumns = [...table.levels].reverse();
  // Withheld before any row is built: the bound is on what the table WOULD
  // hold — OCCUPIED rows × columns, so a sparse cache is not penalised.
  const occupied = occupiedIndices(labelColumns, columns);
  if (occupied.size * (labelColumns.length + columns.length) > bounds.cells) return "withheld";
  const { rows, keptIndices } = occupiedRows(occupied, labelColumns, columns);
  const drawing = withCategoryDirection(
    drawingOf(drawingTypeOf(root), labelColumns.length, columns, keptIndices),
    root,
  );
  return {
    data: { headers: [...labelColumns.map(() => ""), ...columns.map((column) => column.name ?? "")], rows },
    labelColumns: labelColumns.length,
    labelFromX: table.labelFromX,
    ...(drawing === undefined ? {} : { drawing }),
  };
};

/** `"withheld"`: the chart HAS series but a bound keeps it out whole; `undefined`: it has none. */
const tableOf = (
  root: XNode,
  limits: ReaderLimits | undefined,
  sheet: SheetPoints | undefined,
): ChartTable | "withheld" | undefined => {
  const series = findDescendants(root, "c:ser");
  if (series.length === 0) return undefined;
  // The categories win the left column over a scatter's X values, wherever the series stand: decided from the points written.
  const categorised = series.some((ser) => !isXySeries(ser) && carriesPoints(findChild(ser, "c:cat"), sheet));
  const table: TableInProgress = { levels: [[]], labelFromX: false, columns: [], publishedX: [], longest: 0 };
  const bounds = boundsOf(limits);
  let seen = 0;
  for (const ser of series) {
    // Checked before any array of this series is built: the whole chart is withheld.
    if (seen >= bounds.series) return "withheld";
    seen++;
    const added = isXySeries(ser)
      ? xySeriesColumns(ser, table, categorised, limits, sheet)
      : categorySeriesColumns(ser, table, limits, sheet);
    if (added === "withheld") return "withheld";
    // Running bound (chosen bias, written): the dense arrays are sized by
    // their last idx, so (levels + columns incl. this series') × the longest
    // array is checked before each push — a sparse many-series chart is
    // withheld on array length, which keeps the build itself bounded.
    if ((table.levels.length + table.columns.length + added.length) * table.longest > bounds.cells) return "withheld";
    table.columns.push(...added);
  }
  return finishedTable(root, table, bounds);
};

/** The lines of a rich title: one per paragraph (`a:p`) and per line break (`a:br`); the runs of one line are joined as written. */
const richLines = (title: XNode): string[] => {
  const lines: string[] = [];
  for (const paragraph of findDescendants(title, "a:p")) {
    let line = "";
    for (const piece of childNodesOf(paragraph)) {
      if (qualifiedNameOf(piece) === "a:br") {
        lines.push(line);
        line = "";
      } else
        line += findDescendants(piece, "a:t")
          .map((textNode) => textContentOf(textNode))
          .join("");
    }
    lines.push(line);
  }
  return lines;
};

/** The text a `c:title` shows: its rich text, lines separated by one space (never welded: « Total » over « revenue » is « Total revenue »), or,
 *  for a title linked to a cell (`c:tx/c:strRef`), the first point of its `c:strCache` (ECMA-376 Part 1 §21.2.2.178, §21.2.2.176). */
const titleText = (title: XNode | undefined, limits: ReaderLimits | undefined): string | undefined => {
  if (title === undefined) return undefined;
  const rich = richLines(title)
    .map((line) => line.trim())
    .filter((line) => line !== "")
    .join(" ");
  const linked =
    rich === ""
      ? cachePoints(findChild(findChild(findChild(title, "c:tx") ?? {}, "c:strRef") ?? {}, "c:strCache"), limits)?.[0]
      : undefined;
  const text = rich !== "" ? rich : (linked ?? "").trim();
  return text === "" ? undefined : text;
};

/** The chart's OWN title: the `c:title` child of `c:chart` (an axis title is never taken for it). */
const titleOf = (root: XNode, limits: ReaderLimits | undefined): string | undefined =>
  titleText(findChild(findChild(root, "c:chart") ?? {}, "c:title"), limits);

/** An axis the chart deletes (`c:delete`, an optional CT_Boolean: no `val` means true) shows no title. */
const axisIsDeleted = (axis: XNode): boolean => {
  const deleteElement = findChild(axis, "c:delete");
  if (deleteElement === undefined) return false;
  const val = attributeOf(deleteElement, "val");
  return val === undefined || val === "1" || val === "true";
};

interface AxisTitle {
  axis: ChartAxis;
  text: string;
  /** A category or date axis: it may head the table's top-left cell when the categories take it. */
  corner: boolean;
  /** A value axis (`c:valAx`). */
  valued: boolean;
}

const axisRole = (tag: string | undefined, position: string | undefined, byPosition: boolean): ChartAxis => {
  if (tag === "c:serAx") return "z";
  if (tag !== "c:valAx") return "x";
  if (byPosition) return position === "b" || position === "t" ? "x" : "y";
  return "y";
};

/** `x`/`y` are the roles of the chart's DATA, not screen directions: `x` is the category axis even when horizontal bars lay it vertically. */
const axisTitlesOf = (root: XNode, limits: ReaderLimits | undefined): AxisTitle[] => {
  const areas = findDescendants(root, "c:plotArea");
  const axes = areas.flatMap((area) =>
    childNodesOf(area).filter((child) =>
      ["c:catAx", "c:dateAx", "c:serAx", "c:valAx"].includes(qualifiedNameOf(child) ?? ""),
    ),
  );
  // A value axis is classed by its position (`c:axPos`: bottom or top is x, left or right is y) when the chart has no category or date axis
  // (a scatter or bubble alone) AND when it has a scatter or bubble group, whose value axes may be horizontal beside a category axis (a
  // combined chart). Otherwise (bars, lines) a value axis is y whatever its position: horizontal bars lay their value axis at the bottom.
  const categorical = axes.some((axis) => qualifiedNameOf(axis) === "c:catAx" || qualifiedNameOf(axis) === "c:dateAx");
  const xy = areas.some((area) =>
    childNodesOf(area).some(
      (child) => qualifiedNameOf(child) === "c:scatterChart" || qualifiedNameOf(child) === "c:bubbleChart",
    ),
  );
  const byPosition = !categorical || xy;
  const titles: AxisTitle[] = [];
  for (const axis of axes) {
    if (axisIsDeleted(axis)) continue;
    const text = titleText(findChild(axis, "c:title"), limits);
    if (text === undefined) continue;
    const tag = qualifiedNameOf(axis);
    titles.push({
      axis: axisRole(tag, attributeOf(findChild(axis, "c:axPos") ?? {}, "val"), byPosition),
      text,
      corner: tag === "c:catAx" || tag === "c:dateAx",
      valued: tag === "c:valAx",
    });
  }
  return titles;
};

/** The table and the title fields of a chart read. The first category-axis title heads the table's top-left cell when the categories take
 *  ONE column, and the first horizontal value-axis title when that column holds a scatter's X values; every other axis title is a line under the table. */
const withTitles = (
  root: XNode,
  table: ChartTable,
  limits: ReaderLimits | undefined,
): { data: TableData; titles: Pick<PptxChart, "title" | "axisTitles"> } => {
  const title = titleOf(root, limits);
  const axisTitles = axisTitlesOf(root, limits);
  const heading = table.labelFromX
    ? axisTitles.find((axisTitle) => axisTitle.valued && axisTitle.axis === "x")
    : axisTitles.find((axisTitle) => axisTitle.corner);
  const headed = heading !== undefined && table.labelColumns === 1 && table.data.headers !== undefined;
  const data = headed ? { ...table.data, headers: [heading.text, ...table.data.headers!.slice(1)] } : table.data;
  const lines = axisTitles
    .filter((axisTitle) => !headed || axisTitle !== heading)
    .map(({ axis, text }) => ({ axis, text }));
  return {
    data,
    titles: { ...(title === undefined ? {} : { title }), ...(lines.length === 0 ? {} : { axisTitles: lines }) },
  };
};

/** A chart read: its table and titles, `"withheld"` when it is kept out whole, `undefined` when it has no series. */
type ChartOutcome = PptxChart | "withheld" | undefined;

/** The chart part's own rendering: the first image relation whose part EXISTS, so that a relation whose part is absent does not end
 *  the search. The picture is only a fallback: a damaged image part costs that picture, never the chart's data. */
const firstPicture = async (pkg: OpcPackage, part: string): Promise<PptxChart["image"]> => {
  for (const relationship of (await pkg.relationships(part)).values()) {
    if (relationship.kind !== "image" || relationship.external) continue;
    const bytes = await readUnlessDamaged(() => pkg.readBytes(relationship.target));
    if (bytes) return { bytes, contentType: pkg.contentType(relationship.target) ?? "application/octet-stream" };
  }
  return undefined;
};

/** The chart at `part`, `undefined` when it has no series (nothing to say), `"withheld"` when the part is MISSING (said like a damaged
 *  part, the same notice in every format) or when the chart has series and a bound (points, levels, cells, series) keeps it out whole: its title and
 *  labels are then not on the page, which the caller reports. */
export const readChartOutcome = async (
  pkg: OpcPackage,
  part: string,
  options: { picture?: boolean; limits?: ReaderLimits | undefined; resolver?: ChartResolver | undefined } = {},
): Promise<ChartOutcome> => {
  const root = await pkg.readXml(part);
  if (!root) return "withheld";
  const sheet = options.resolver === undefined ? undefined : sheetPointsOf(root, options.resolver, options.limits);
  if (sheet === "withheld") return "withheld";
  const table = tableOf(root, options.limits, sheet);
  if (table === undefined || table === "withheld") return table;
  const { data, titles } = withTitles(root, table, options.limits);
  // A caller that publishes no picture (docx, xlsx) never reads its bytes.
  const drawing = table.drawing === undefined ? {} : { drawing: table.drawing };
  if (options.picture === false) return { data, ...titles, ...drawing };
  const image = await firstPicture(pkg, part);
  return { data, ...titles, ...drawing, ...(image === undefined ? {} : { image }) };
};

/** `readChartOutcome`, with a chart part that is DAMAGED (a corrupt stream, an XML that does not parse) read as `"withheld"`, like a missing one:
 *  the chart is kept out whole and the rest of the document stands. Any other failure still throws. */
export const readChartOrWithheld = async (
  pkg: OpcPackage,
  part: string,
  options: { picture?: boolean; limits?: ReaderLimits | undefined; resolver?: ChartResolver | undefined } = {},
): Promise<ChartOutcome> => {
  // The wrapper makes `undefined` mean "damaged"; the outcome's own `undefined` (no series) says nothing.
  const read = await readUnlessDamaged(async () => ({ outcome: await readChartOutcome(pkg, part, options) }));
  return read === undefined ? "withheld" : read.outcome;
};

/** The IR block of a chart read, and whether a bound kept it out whole (`withheld`: its title and labels are not on the
 *  page, which the caller reports in its own notice). `collectImage` stores the chart's picture, when the caller keeps
 *  it as a fallback (pptx); without it the block carries no picture. */
export const chartToIr = (
  read: ChartOutcome,
  options: { collectImage?: (image: NonNullable<PptxChart["image"]>) => ImageRef } = {},
): { block?: IRBlock; withheld: boolean } => {
  if (read === "withheld") return { withheld: true };
  if (read === undefined) return { withheld: false };
  const image = read.image && options.collectImage ? options.collectImage(read.image) : undefined;
  return {
    withheld: false,
    block: {
      kind: "chart",
      data: read.data,
      ...(read.drawing === undefined ? {} : { drawing: read.drawing }),
      ...(read.title === undefined ? {} : { title: read.title }),
      ...(read.axisTitles === undefined ? {} : { axisTitles: read.axisTitles }),
      ...(image === undefined ? {} : { image }),
    },
  };
};
