import fc from "fast-check";
import { describe, expect, it } from "vitest";

import { type DiagramConnector, type DiagramElement, orderDiagram } from "./diagram-graph";
import { fuzzRuns } from "../lib/docs-fuzz-runs.test.support";

/**
 * Priority 1 of the diagram text: whatever the shapes and connectors, the
 * order we hand back holds every element exactly once — the multiset of
 * words before is the multiset after (the fixtures-based twin is `extractor/diagram-text.order.fixtures.test.ts`).
 */

const boxArb = fc.record({
  x: fc.integer({ min: 0, max: 1000 }),
  y: fc.integer({ min: 0, max: 1000 }),
  cx: fc.integer({ min: 0, max: 300 }),
  cy: fc.integer({ min: 0, max: 300 }),
});

const elementArb = (ids: readonly string[]): fc.Arbitrary<DiagramElement> =>
  fc.record({
    id: fc.option(fc.constantFrom(...ids), { nil: undefined }),
    box: fc.option(boxArb, { nil: undefined }),
    node: fc.boolean(),
  }) as fc.Arbitrary<DiagramElement>;

const connectorArb = (ids: readonly string[]): fc.Arbitrary<DiagramConnector> =>
  fc.record({
    from: fc.option(fc.constantFrom(...ids, "missing"), { nil: undefined }),
    to: fc.option(fc.constantFrom(...ids, "missing"), { nil: undefined }),
    ends: fc.option(
      fc.record({
        x1: fc.integer({ min: 0, max: 1300 }),
        y1: fc.integer({ min: 0, max: 1300 }),
        x2: fc.integer({ min: 0, max: 1300 }),
        y2: fc.integer({ min: 0, max: 1300 }),
      }),
      { nil: undefined },
    ),
    startArrow: fc.boolean(),
    endArrow: fc.boolean(),
    visible: fc.boolean(),
  }) as fc.Arbitrary<DiagramConnector>;

describe("orderDiagram (property)", () => {
  it("paragraphs then rest is a permutation of the elements: nothing lost, nothing repeated", () => {
    const ids = ["1", "2", "3", "4", "5", "6"];
    fc.assert(
      fc.property(
        fc.array(elementArb(ids), { maxLength: 12 }),
        fc.array(connectorArb(ids), { maxLength: 14 }),
        fc.array(
          fc.record({
            id: fc.option(fc.constantFrom(...ids), { nil: undefined }),
            box: fc.option(boxArb, { nil: undefined }),
          }),
          { maxLength: 4 },
        ) as fc.Arbitrary<{ id?: string; box?: { x: number; y: number; cx: number; cy: number } }[]>,
        (elements, connectors, obstacles) => {
          const r = orderDiagram({
            elements,
            obstacles,
            connectors,
            tolerances: { endTouch: 40, labelReach: 60, ambiguity: 1.5 },
          });
          const all = [...r.paragraphs.flat(), ...r.rest];
          expect([...all].sort((a, b) => a - b)).toEqual(elements.map((_, i) => i));
        },
      ),
      { numRuns: fuzzRuns(400, process.env) },
    );
  });

  it("random trees of declared connectors are all reordered, each element once", () => {
    fc.assert(
      fc.property(
        fc
          .integer({ min: 2, max: 10 })
          .chain((n) =>
            fc.tuple(
              fc.constant(n),
              fc.array(fc.nat(), { minLength: n - 1, maxLength: n - 1 }),
              fc.shuffledSubarray([...Array(n).keys()], { minLength: n, maxLength: n }),
            ),
          ),
        ([n, parents, perm]) => {
          // Element i sits at rank perm[i]; node k's parent is an earlier node of the permutation: a random tree.
          const elements: DiagramElement[] = Array.from({ length: n }, (_, i) => ({
            id: String(i),
            box: { x: i * 200, y: 0, cx: 100, cy: 40 },
            node: true,
          }));
          const connectors: DiagramConnector[] = perm.slice(1).map((child, k) => ({
            from: String(perm[parents[k]! % (k + 1)]),
            to: String(child),
            startArrow: false,
            endArrow: true,
            visible: true,
          }));
          const r = orderDiagram({
            elements,
            obstacles: [],
            connectors,
            tolerances: { endTouch: 40, labelReach: 60, ambiguity: 1.5 },
          });
          expect(r.rest).toEqual([]);
          expect([...r.paragraphs.flat()].sort((a, b) => a - b)).toEqual(elements.map((_, i) => i));
        },
      ),
      { numRuns: fuzzRuns(200, process.env) },
    );
  });
});
