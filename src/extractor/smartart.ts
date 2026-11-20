/** The words of a SmartArt, read from its DATA MODEL.
 *
 *  Source: ECMA-376 Part 1 §21.4 (DrawingML Diagrams): `dgm:dataModel`
 *  (§21.4.2.10) holds `dgm:ptLst` (the points, the text in each `dgm:pt`'s
 *  `dgm:t`) and `dgm:cxnLst` (§21.4.3.3) of `dgm:cxn` (§21.4.3.2), a
 *  connection from `srcId` to `destId`. A connection with no `type` is
 *  `parOf` (the schema's default): the parent-child tree of the diagram's
 *  text, the children of a parent in `srcOrd` order.
 *
 *  Chosen order, in this order of preference:
 *  1. the TREE of `parOf` connections, from the `doc` point (or the points
 *     with no parent), children by `srcOrd`, each point once;
 *  2. the document order of the points when the relations are absent or do
 *     not form a tree (a point with two parents, a cycle, a missing end).
 *  Points the tree does not reach (`parTrans`, `sibTrans`, `pres`) come last,
 *  in document order, when they carry text: no word is ever dropped.
 *
 *  Why the data model and not the drawing cache: the data model is what the
 *  author typed and what the schema defines; the cache (`dsp:drawing`, a
 *  Microsoft extension, [MS-ODRAWXML]) is what Office last RENDERED, with
 *  the layout's own placement. It is read only when the data model is
 *  missing or holds no text (a producer that wrote the cache alone).
 *
 *  No character is added: one piece per point (its paragraphs joined by a
 *  space), the caller separates the pieces.
 *
 *  Each piece carries its LEVEL in the tree, counted from the shallowest text
 *  actually published (the `doc` point carries none, so its children are 0).
 *  The page publishes a SmartArt's words as a nested bullet list, nested on
 *  these levels; a model that is not a tree gives every piece level 0, so the
 *  list stays flat and invents no structure.
 *  NOT here: where the words go (`extractor/pptx.ts`, `docx-reader/body.ts`),
 *  nor the list itself (`renderer/storage.ts#renderOutline`).
 *  Called by: `pptx-reader/index.ts`, `pptx-reader/shapes.ts`, `docx-reader/drawing-objects.ts`. */

import { collapseBlanks } from "doc-ir/internal";
import { byCodeUnit } from "doc-ir/internal";
import { findInternalTarget, readUnlessDamaged, type OpcPackage, type Relationship } from "./ooxml/package";
import {
  attributeOf,
  findChild,
  findChildren,
  findDescendants,
  childNodesOf,
  qualifiedNameOf,
  textContentOf,
  type XNode,
} from "./ooxml/xml";

/** One text body (`dgm:t` or `dsp:txBody`) as one line: its paragraphs'
 *  runs (`a:t`, a field's included), a line break (`a:br`) as a space,
 *  paragraphs joined by a space. Ends trimmed, nothing else touched. */
const lineOf = (body: XNode): string => {
  const paragraphText = (paragraph: XNode): string =>
    childNodesOf(paragraph)
      .map((childNode) => {
        const tag = qualifiedNameOf(childNode);
        if (tag === "a:br") return " ";
        if (tag === "a:r" || tag === "a:fld") return findDescendants(childNode, "a:t").map(textContentOf).join("");
        return "";
      })
      .join("");
  return collapseBlanks(findChildren(body, "a:p").map(paragraphText).join(" "));
};

/** Starting value to calibrate: the most points, and the most connections,
 *  of a data model whose TREE is walked. The SmartArts of the repo's
 *  fixtures hold tens of points; a model of 40 000 points cost 14 s and a
 *  chain of 20 000 overflowed the stack. Past it the document order of the
 *  points stands: every word is published, only the tree's order is given up. */
const MAX_SMARTART_POINTS = 5_000;

interface Point {
  id: string | undefined;
  text: string;
}

/** One `parOf` connection (`dgm:cxn`): its two ends as ranks in `dgm:ptLst` (`undefined`: no point carries that id), its `srcOrd`, and
 *  its rank among the `parOf` connections of `dgm:cxnLst`. */
interface Edge {
  from: number | undefined;
  to: number | undefined;
  srcOrd: number;
  order: number;
}

/** One text of a data model and its depth in the `parOf` tree, counted from
 *  the shallowest text actually published: the `doc` point carries no text,
 *  so its children are level 0. A model whose relations are not a tree gives
 *  every text level 0 — a flat list invents no structure. */
export interface DataModelPiece {
  text: string;
  level: number;
}

/** The `parOf` connections of a data model, in document order; a `dgm:cxn` that writes no `type` is one. */
const parentChildEdges = (root: XNode, points: readonly Point[]): Edge[] => {
  const byId = new Map<string, number>();
  points.forEach((point, index) => {
    if (point.id !== undefined) byId.set(point.id, index);
  });
  return findChildren(findChild(root, "dgm:cxnLst") ?? {}, "dgm:cxn")
    .filter((connection) => (attributeOf(connection, "type") ?? "parOf") === "parOf")
    .map((connection, order) => ({
      from: byId.get(attributeOf(connection, "srcId") ?? ""),
      to: byId.get(attributeOf(connection, "destId") ?? ""),
      srcOrd: Number(attributeOf(connection, "srcOrd") ?? "0"),
      order,
    }));
};

/** The parent of each point that has one. `undefined` when the connections are not a tree: an end that is no point, or a point with
 *  two parents. */
const parentsOf = (edges: readonly Edge[]): Map<number, number> | undefined => {
  const parentOf = new Map<number, number>();
  for (const edge of edges) {
    if (edge.from === undefined || edge.to === undefined || parentOf.has(edge.to)) return undefined;
    parentOf.set(edge.to, edge.from);
  }
  return parentOf;
};

/** The connections to each parent's children, in `srcOrd` order, then in document order. Called once `parentsOf` has found every end. */
const childEdgesByParent = (edges: readonly Edge[]): Map<number, Edge[]> => {
  const childEdgesOf = new Map<number, Edge[]>();
  for (const edge of edges) {
    const siblings = childEdgesOf.get(edge.from!);
    if (siblings === undefined) childEdgesOf.set(edge.from!, [edge]);
    else siblings.push(edge);
  }
  for (const siblings of childEdgesOf.values())
    siblings.sort((first, second) => first.srcOrd - second.srcOrd || first.order - second.order);
  return childEdgesOf;
};

/** The texts the tree reaches from `roots`, in pre-order, each at its depth, and the points visited. `undefined` when a point is reached
 *  twice: the connections hold a cycle. */
const walkTree = (
  points: readonly Point[],
  roots: readonly number[],
  childEdgesOf: ReadonlyMap<number, Edge[]>,
): { walked: DataModelPiece[]; seen: Set<number> } | undefined => {
  const seen = new Set<number>();
  const walked: DataModelPiece[] = [];
  let cyclic = false;
  // An explicit stack: a deep chain must not overflow the call stack.
  const stack: { children: Edge[]; next: number }[] = [];
  const visit = (pointIndex: number): void => {
    if (seen.has(pointIndex)) {
      cyclic = true;
      return;
    }
    seen.add(pointIndex);
    const text = points[pointIndex]?.text ?? "";
    // The depth is the height of the stack at the moment of the visit.
    if (text !== "") walked.push({ text, level: stack.length });
    stack.push({ children: childEdgesOf.get(pointIndex) ?? [], next: 0 });
  };
  for (const root of roots) {
    visit(root);
    while (stack.length > 0 && !cyclic) {
      const frame = stack[stack.length - 1]!;
      if (frame.next >= frame.children.length) stack.pop();
      else visit(frame.children[frame.next++]!.to!);
    }
  }
  return cyclic ? undefined : { walked, seen };
};

/** The pieces of text of a data model, in the order described above. */
export const readDataModelPieces = (root: XNode): DataModelPiece[] => {
  const pointNodes = findChildren(findChild(root, "dgm:ptLst") ?? {}, "dgm:pt");
  const points: Point[] = pointNodes.map((pointNode) => ({
    id: attributeOf(pointNode, "modelId"),
    text: lineOf(findChild(pointNode, "dgm:t") ?? {}),
  }));
  const inDocumentOrder = (): DataModelPiece[] =>
    points.filter((point) => point.text !== "").map((point) => ({ text: point.text, level: 0 }));

  const edges = parentChildEdges(root, points);
  if (edges.length === 0 || points.length > MAX_SMARTART_POINTS || edges.length > MAX_SMARTART_POINTS)
    return inDocumentOrder();
  const parentOf = parentsOf(edges);
  if (parentOf === undefined) return inDocumentOrder();
  // The walk starts from the `doc` point, else from every point with no parent.
  const docIndex = pointNodes.findIndex((pointNode) => attributeOf(pointNode, "type") === "doc");
  const roots = docIndex >= 0 ? [docIndex] : points.map((_, index) => index).filter((index) => !parentOf.has(index));
  // A cycle the walk enters gives the document order; one it never reaches leaves its points unreached (below).
  const tree = walkTree(points, roots, childEdgesByParent(edges));
  if (tree === undefined) return inDocumentOrder();
  // Counted from the shallowest text, not from the roots: a `doc` point carries no text, so without this its children
  // would be level 1 and the published list would start one notch in. `reduce` and not `Math.min(...)`: the spread of a
  // model of thousands of points is one more way to overflow the stack, which this walk exists to avoid.
  const base = tree.walked.reduce((least, piece) => Math.min(least, piece.level), Number.MAX_SAFE_INTEGER);
  const pieces = tree.walked.map((piece) => ({ text: piece.text, level: piece.level - base }));
  // A point the tree does not reach has no depth: it goes at level 0, never nested under a text it is not a child of.
  points.forEach((point, index) => {
    if (!tree.seen.has(index) && point.text !== "") pieces.push({ text: point.text, level: 0 });
  });
  return pieces;
};

/** The drawing cache's pieces: one per shape (`dsp:sp`) with text, in
 *  document order. */
export const readDrawingPieces = (root: XNode): string[] =>
  findDescendants(root, "dsp:sp")
    .map((sp) => lineOf(findChild(sp, "dsp:txBody") ?? {}))
    .filter((text) => text !== "");

/** The data-model part a `dgm:relIds` element names (`r:dm`), when internal. */
export const findSmartArtDataPart = (
  relationships: ReadonlyMap<string, Relationship>,
  relIds: XNode,
): string | undefined => findInternalTarget(relationships, attributeOf(relIds, "r:dm"));

export interface SmartArtRead {
  words: string[];
  /** The level of each word's piece in the model's tree, one per `words` entry, 0 when the model is not a tree or the
   *  words come from the drawing cache (which has no tree). The published list nests on it. */
  levels: number[];
  /** The drawing-cache part the data model names (`dsp:dataModelExt relId`,
   *  [MS-ODRAWXML]), whether or not its text was used: the cache is THAT
   *  frame's, and no model-less frame may be given it. */
  claimedDrawingPart?: string;
}

/** A SmartArt's words, from its data-model part. `relationships`: those
 *  of the part that holds the frame (slide or document), where the data
 *  model's own extension names the drawing cache. `undefined`: no data-model
 *  part (the caller may look for a cache nobody claimed,
 *  `readCachesWithoutModel`). */
export const readSmartArt = async (
  pkg: OpcPackage,
  relationships: ReadonlyMap<string, Relationship>,
  dataPart: string | undefined,
): Promise<SmartArtRead | undefined> => {
  const data = dataPart === undefined ? undefined : await pkg.readXml(dataPart);
  if (!data) return undefined;
  const extension = findDescendants(data, "dsp:dataModelExt")[0];
  const relId = extension ? attributeOf(extension, "relId") : undefined;
  const cachePart = findInternalTarget(relationships, relId);
  const fromModel = readDataModelPieces(data);
  if (fromModel.length > 0 || cachePart === undefined)
    return {
      words: fromModel.map((piece) => piece.text),
      levels: fromModel.map((piece) => piece.level),
      ...(cachePart === undefined ? {} : { claimedDrawingPart: cachePart }),
    };
  // A model with no text: the cache is the only source, and it holds no tree — a flat list.
  const cache = await pkg.readXml(cachePart);
  const words = cache ? readDrawingPieces(cache) : [];
  return { words, levels: words.map(() => 0), claimedDrawingPart: cachePart };
};

/** The number in a part name (`drawing12.xml` gives 12); parts without one sort last. */
export const partNumberOf = (path: string): number =>
  Number(/(\d{1,15})\.xml$/.exec(path)?.[1] ?? Number.MAX_SAFE_INTEGER);

/** The words of the SmartArts that have NO data model, `frames` of them in
 *  reading order, from the drawing caches no model claimed: the part's
 *  `diagramDrawing` relationships (type
 *  `http://schemas.microsoft.com/office/2007/relationships/diagramDrawing`),
 *  in the order of their part names' numbers. A frame names only its data
 *  model (`dgm:relIds` has no cache attribute), so a cache cannot be tied to
 *  its frame by the file: as many caches as frames → one each, in order;
 *  otherwise ALL of their words go to the first frame (chosen bias: an
 *  extra place is better than a lost word, never a word that is not the
 *  file's own). Two parts with the same number are ordered by their names'
 *  UTF-16 code units: the same order on every machine. */
export const readCachesWithoutModel = async (
  pkg: OpcPackage,
  relationships: ReadonlyMap<string, Relationship>,
  claimed: ReadonlySet<string>,
  frames: number,
): Promise<string[][]> => {
  if (frames === 0) return [];
  const targets = [
    ...new Set(
      [...relationships.values()]
        .filter((relationship) => relationship.kind === "diagramDrawing" && !relationship.external)
        .map((relationship) => relationship.target),
    ),
  ]
    .filter((target) => !claimed.has(target))
    .sort((first, second) => partNumberOf(first) - partNumberOf(second) || byCodeUnit(first, second));
  const read: string[][] = [];
  for (const target of targets) {
    // A cache whose XML cannot be read gives no words; a ceiling refusal or any other error still propagates.
    const root = await readUnlessDamaged(() => pkg.readXml(target));
    read.push(root ? readDrawingPieces(root) : []);
  }
  if (read.length === frames) return read;
  return Array.from({ length: frames }, (_, frameIndex) => (frameIndex === 0 ? read.flat() : []));
};
