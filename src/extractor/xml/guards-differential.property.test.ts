/**
 * A fuzz of the DEPTH guard on REAL parts of the versioned fixtures, wrapped in enough elements to land around the depth
 * boundary; its name keeps « differential » from the time each case was compared with another parser, and is not
 * renamed, by project rule.
 *
 * Rule held, EXACT (`tree.ts#open`: a start tag is refused when more than MAX_XML_DEPTH elements are open, an
 * empty-element tag never is; `guards.test.ts` holds the same table on hand-written parts). A part is refused exactly
 * when its deepest NON-self-closing element sits at level MAX_XML_DEPTH + 2 or deeper, and then with « elements nested
 * past the depth limit »; it is read otherwise, and the tree read is as deep as the text written. With `levels` the
 * wrapper's elements, `seed.depth` the seed's deepest level and `seed.openDepth` its deepest element that is not an
 * empty-element tag:
 * - innermost `emptyPair` (a branch ending in `<y></y>`): refused exactly when levels + seed.depth >= MAX_XML_DEPTH + 2;
 * - innermost `seed` and `emptyTag` (a branch ending in `<y/>`): refused exactly when levels + seed.openDepth >= MAX_XML_DEPTH + 2.
 * The wrapper is written so that both sides of the boundary (start tag versus empty-element tag) are reached.
 *
 * Seeds: for each kind of part (its name with digits dropped), the smallest one under 8 KiB that our reader reads.
 * fast-check only chooses the seed, the distance to the boundary and the innermost wrapper.
 */
import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { MAX_XML_DEPTH } from "../ooxml/xml";
import { tokenize } from "./tokenizer";
import { buildTree } from "./tree";
import { fixtureXmlParts, type XNode } from "./xml-fixtures.test.support";
import { readerAccepts } from "./xml-shadow.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(60_000, process.env) });

const SEED_MAX_CHARS = 8 * 1024;

interface Seed {
  readonly text: string;
  /** The deepest level an element of the seed reaches, an empty element counting as a level. */
  readonly depth: number;
  /** The deepest level of an element that is not an empty-element tag (0 when there is none). */
  readonly openDepth: number;
}

let seeds: Seed[] = [];

const depthsOf = (text: string): { depth: number; openDepth: number } => {
  let open = 0;
  let depth = 0;
  let openDepth = 0;
  tokenize(text, {
    open(_name, _attributes, selfClosing) {
      depth = Math.max(depth, open + 1);
      if (!selfClosing) {
        open++;
        openDepth = Math.max(openDepth, open);
      }
    },
    close() {
      open--;
    },
    text() {},
  });
  return { depth, openDepth };
};

/** The element nesting of a tree, an empty element counting as a level; walked without recursion. */
const treeDepth = (nodes: readonly XNode[]): number => {
  let deepest = 0;
  const pending: [XNode, number][] = nodes.map((node) => [node, 1]);
  while (pending.length > 0) {
    const [node, depth] = pending.pop()!;
    const tag = Object.keys(node).find((key) => key !== ":@")!;
    if (tag === "#text") continue;
    deepest = Math.max(deepest, depth);
    for (const kid of node[tag] as XNode[]) pending.push([kid, depth + 1]);
  }
  return deepest;
};

beforeAll(async () => {
  const best = new Map<string, number>();
  const kept = await fixtureXmlParts((part) => {
    if (part.text.length > SEED_MAX_CHARS) return false;
    const kind = part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, "");
    const known = best.get(kind);
    if (known !== undefined && known <= part.text.length) return false;
    if (!readerAccepts(part.text)) return false;
    best.set(kind, part.text.length);
    return true;
  });
  const smallest = new Map<string, string>();
  for (const part of kept) smallest.set(part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, ""), part.text);
  seeds = [...smallest.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, text]) => ({ text, ...depthsOf(text) }))
    .filter((seed) => seed.depth > 0);
}, 180_000);

/** The innermost level: the seed alone, or beside it a branch as deep as the seed ending in an empty pair or an
 *  empty-element tag, so that both kinds of tag meet the boundary. */
type Innermost = "seed" | "emptyPair" | "emptyTag";

/** `seed` wrapped in `levels` elements, the declaration kept first. */
const wrapped = (seed: Seed, levels: number, innermost: Innermost): string => {
  const declaration = /^\uFEFF?<\?xml[^>]*>/.exec(seed.text)?.[0] ?? "";
  const body = seed.text.slice(declaration.length);
  const leaf = innermost === "emptyPair" ? "<y></y>" : "<y/>";
  const branch = innermost === "seed" ? "" : "<y>".repeat(seed.depth - 1) + leaf + "</y>".repeat(seed.depth - 1);
  return declaration + "<x>".repeat(levels) + body + branch + "</x>".repeat(levels);
};

/** Our reading of a part: the tree, or our refusal's message; a ReaderError that is not one of ours (`xml: …`) and any
 *  other throw are crashes and propagate. */
const reading = (text: string): { tree: XNode[] } | { refusal: string } => {
  try {
    return { tree: buildTree(text, "document") };
  } catch (error) {
    if (error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: "))
      return { refusal: error.message };
    throw error;
  }
};

describe("xml guards — fuzz of the depth guard on real parts", () => {
  it("found a seed for each common kind of part", () => {
    expect(seeds.length).toBeGreaterThan(20);
  });

  it("around the boundary, ours refuses exactly when a start tag sits past the limit, with the depth literal, and reads the depth written", () => {
    fc.assert(
      fc.property(
        fc.nat(),
        fc.integer({ min: -3, max: 3 }),
        fc.constantFrom<Innermost>("seed", "emptyPair", "emptyTag"),
        (which, delta, innermost) => {
          const seed = seeds[which % seeds.length]!;
          const levels = Math.max(1, MAX_XML_DEPTH - seed.depth + delta);
          const written = levels + seed.depth;
          const deepestStartTag = levels + (innermost === "emptyPair" ? seed.depth : seed.openDepth);
          const outcome = reading(wrapped(seed, levels, innermost));
          if (deepestStartTag >= MAX_XML_DEPTH + 2)
            expect(outcome, `${innermost}: a start tag at level ${deepestStartTag}`).toEqual({
              refusal: "xml: elements nested past the depth limit",
            });
          else {
            expect(
              "tree" in outcome,
              `${innermost}: read at ${written} levels, deepest start tag ${deepestStartTag}`,
            ).toBe(true);
            expect(treeDepth((outcome as { tree: XNode[] }).tree), `${innermost}: depth of the tree read`).toBe(
              written,
            );
          }
        },
      ),
      { numRuns: fuzzRuns(300, process.env) },
    );
  });
});
