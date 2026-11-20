import { describe, expect, it } from "vitest";

import { type DiagramConnector, type DiagramElement, orderDiagram, type DiagramTolerances } from "./diagram-graph";

// Units are EMU in the callers; here boxes are small round numbers.
const TOL: DiagramTolerances = { endTouch: 10, labelReach: 50, ambiguity: 1.5 };

const box = (x: number, y: number, cx = 100, cy = 40) => ({ x, y, cx, cy });
const node = (id: string, x: number, y: number, cx = 100, cy = 40): DiagramElement => ({
  id,
  box: box(x, y, cx, cy),
  node: true,
});
const declared = (from: string, to: string, arrow: "end" | "start" | "none" = "end"): DiagramConnector => ({
  from,
  to,
  startArrow: arrow === "start",
  endArrow: arrow === "end",
  visible: true,
});
const line = (
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  arrow: "end" | "start" | "none" = "end",
): DiagramConnector => ({
  ends: { x1, y1, x2, y2 },
  startArrow: arrow === "start",
  endArrow: arrow === "end",
  visible: true,
});

describe("orderDiagram: declared connectors", () => {
  it("one relation: source then target, the rest after", () => {
    const els = [node("1", 0, 0), node("9", 0, 500), node("2", 300, 0)];
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [declared("1", "2")], tolerances: TOL });
    expect(r.paragraphs).toEqual([[0, 2]]);
    expect(r.rest).toEqual([1]);
  });

  it("a chain stays one paragraph; two relations are two paragraphs", () => {
    const els = [node("a", 0, 0), node("c", 0, 200), node("b", 300, 0), node("d", 300, 200), node("e", 600, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("c", "d"), declared("a", "b"), declared("b", "e")],
      tolerances: TOL,
    });
    // Components in reading order of their first element: {a, b, e} then {c, d}.
    expect(r.paragraphs).toEqual([
      [0, 2, 4],
      [1, 3],
    ]);
    expect(r.rest).toEqual([]);
  });

  it("a root with three children: the first child continues the paragraph, the others open one each", () => {
    const els = [node("d", 200, 0), node("s", 0, 200), node("e", 200, 200), node("t", 400, 200)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("d", "t"), declared("d", "s"), declared("d", "e")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1], [2], [3]]);
  });

  it("the arrow head decides the direction: a head at the start reverses", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 0, 200), node("d", 300, 200)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b", "start"), declared("c", "d", "end")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([
      [1, 0],
      [2, 3],
    ]);
  });

  // 02/10/2026: a plain line (no head) has no direction (stCxn -> endCxn is arbitrary there).
  it("a connector with no arrow head says nothing about the order (its direction is arbitrary)", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [declared("a", "b", "none")], tolerances: TOL });
    expect(r.paragraphs).toEqual([]);
    expect(r.rest).toEqual([0, 1]);
    expect(r.abstentions["no-arrowhead"]).toBe(1);
  });

  it("a plain line inside a component does not break its tree: the arrowed ones alone are read", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 600, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b", "end"), declared("b", "c", "none"), declared("a", "c", "none")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1]]);
    expect(r.rest).toEqual([2]);
  });

  it("the end of a plain line is not a label of a nearby arrow", () => {
    const els = [
      node("a", 0, 0),
      node("b", 400, 0),
      { id: "m", box: box(190, -30, 120, 20), node: true } as DiagramElement,
      node("z", 190, 300),
    ];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [
        { ...line(100, 20, 400, 20) },
        { from: "m", to: "z", startArrow: false, endArrow: false, visible: true },
      ],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1]]);
    expect(r.rest).toEqual([2, 3]);
  });

  // 02/10/2026: a component with no cycle that is no tree is read in topological order, ties by reading order.
  it("a merge (two sources, one target) is no tree but has no cycle: topological order, the sources in reading order, one paragraph", () => {
    const els = [node("a", 0, 0), node("b", 0, 100), node("c", 300, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("b", "c"), declared("a", "c")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1, 2]]);
    expect(r.rest).toEqual([]);
  });

  it("every source comes before its target, whatever the reading order says", () => {
    const els = [node("t", 0, 0), node("u", 0, 100), node("s", 0, 200), node("x", 0, 300)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("s", "t"), declared("u", "t"), declared("s", "u")],
      tolerances: TOL,
    });
    // s (2) first, then u (1), then t (0); x has no relation.
    expect(r.paragraphs).toEqual([[2, 1, 0]]);
    expect(r.rest).toEqual([3]);
  });

  it("a cycle keeps today's order and says why; a double connector is no cycle", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    const cyc = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b"), declared("b", "a")],
      tolerances: TOL,
    });
    expect(cyc.paragraphs).toEqual([]);
    expect(cyc.rest).toEqual([0, 1]);
    expect(cyc.abstentions["has-cycle"]).toBe(1);
    expect(
      orderDiagram({
        elements: els,
        obstacles: [],
        connectors: [declared("a", "b"), declared("a", "b")],
        tolerances: TOL,
      }).paragraphs,
    ).toEqual([[0, 1]]);
  });

  it("a tree component is reordered while a component with a cycle keeps today's order", () => {
    const els = [node("a", 0, 0), node("m1", 0, 300), node("b", 300, 0), node("m2", 300, 300), node("m3", 600, 300)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b"), declared("m1", "m2"), declared("m2", "m3"), declared("m3", "m1")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 2]]);
    expect(r.rest).toEqual([1, 3, 4]);
  });

  it("a declared end on a shape without text, or on an unknown id, abstains", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    const obstacles = [{ id: "7", box: box(0, 300) }];
    const r = orderDiagram({
      elements: els,
      obstacles,
      connectors: [declared("a", "7"), declared("a", "zz")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([]);
    expect(r.abstentions["end-on-shape-without-text"]).toBe(1);
    expect(r.abstentions["declared-unknown"]).toBe(1);
  });

  it("an invisible connector says nothing", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [{ ...declared("a", "b"), visible: false }],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([]);
    expect(r.abstentions["invisible"]).toBe(1);
  });
});

// 02/10/2026 (High_level_IT_Landscape: hubs with three arrows): a component that is not a tree keeps today's
// order, and its arrows are given as relations, one per target.
describe("orderDiagram: the relations of a component that is not a tree", () => {
  it("a component with a cycle gives its relations and keeps today's order", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 600, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b"), declared("b", "c"), declared("c", "a")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([]);
    expect(r.links).toHaveLength(3);
  });

  it("a merge gives one relation: the target and its sources", () => {
    const els = [node("a", 0, 0), node("b", 0, 100), node("c", 300, 0)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("b", "c"), declared("a", "c")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1, 2]]);
    expect(r.rest).toEqual([]);
    expect(r.links).toEqual([{ target: 2, sources: [0, 1] }]);
  });

  it("a head at the start reverses the relation; a double connector counts once; plain lines are no relation", () => {
    const els = [node("a", 0, 0), node("b", 0, 100), node("c", 300, 0), node("d", 300, 100)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [
        declared("a", "c", "start"),
        declared("b", "a"),
        declared("b", "a"),
        declared("c", "d", "none"),
        declared("d", "a", "none"),
      ],
      tolerances: TOL,
    });
    expect(r.links).toEqual([{ target: 0, sources: [1, 2] }]);
    expect(r.paragraphs).toEqual([[1, 2, 0]]);
  });

  it("a cycle gives a relation each way; a tree gives none (it is ordered instead)", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    expect(
      orderDiagram({
        elements: els,
        obstacles: [],
        connectors: [declared("a", "b"), declared("b", "a")],
        tolerances: TOL,
      }).links,
    ).toEqual([
      { target: 0, sources: [1] },
      { target: 1, sources: [0] },
    ]);
    expect(
      orderDiagram({ elements: els, obstacles: [], connectors: [declared("a", "b")], tolerances: TOL }).links,
    ).toEqual([]);
  });

  it("a tree component beside a hub keeps its order and gives no relation of its own; the hub is read topologically", () => {
    const els = [node("a", 0, 0), node("m1", 0, 300), node("b", 300, 0), node("m2", 300, 300), node("m3", 600, 300)];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b"), declared("m1", "m3"), declared("m2", "m3")],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([
      [0, 2],
      [1, 3, 4],
    ]);
    expect(r.links).toEqual([{ target: 4, sources: [1, 3] }]);
  });
});

describe("orderDiagram: geometry only", () => {
  it("both ends within the tolerance of exactly one shape each", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    // a's right edge is x = 100, b's left edge is x = 300; ends 4 units off.
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [line(104, 20, 296, 20)], tolerances: TOL });
    expect(r.paragraphs).toEqual([[0, 1]]);
  });

  it("an end beyond the tolerance, or inside a shape far from its edge, abstains", () => {
    const els = [node("a", 0, 0, 100, 400), node("b", 300, 0)];
    expect(
      orderDiagram({ elements: els, obstacles: [], connectors: [line(130, 20, 296, 20)], tolerances: TOL }).paragraphs,
    ).toEqual([]);
    const near = orderDiagram({ elements: els, obstacles: [], connectors: [line(50, 200, 296, 20)], tolerances: TOL });
    expect(near.paragraphs).toEqual([]);
    expect(near.abstentions["end-touches-nothing"]).toBe(1);
  });

  it("an end touching two shapes abstains; a frame around the touched shape is dropped (innermost)", () => {
    const two = [node("a", 0, 0), node("a2", 100, 0), node("b", 300, 0)];
    const r = orderDiagram({ elements: two, obstacles: [], connectors: [line(100, 20, 296, 20)], tolerances: TOL });
    expect(r.paragraphs).toEqual([]);
    expect(r.abstentions["end-touches-several"]).toBe(1);
    // A frame whose edge lies within the tolerance of the end, containing b.
    const framed = [node("a", 0, 0), node("b", 300, 0), node("frame", 295, -10, 300, 80)];
    const ok = orderDiagram({ elements: framed, obstacles: [], connectors: [line(104, 20, 297, 20)], tolerances: TOL });
    expect(ok.paragraphs).toEqual([[0, 1]]);
  });

  // Audit 07/10, F20 (migration to `layout/rect.ts`): the migration showed that no test held the STRICT side of « holds ». A shape
  // is the frame of another only if it is strictly larger on one axis; two shapes of the SAME box are not frames of each other.
  it("two shapes of the very same box are not frames of each other: an end touching both abstains as ambiguous", () => {
    const twins = [node("a", 0, 0), node("b", 300, 0), node("b2", 300, 0)];
    const r = orderDiagram({ elements: twins, obstacles: [], connectors: [line(104, 20, 296, 20)], tolerances: TOL });
    expect(r.paragraphs).toEqual([]);
    // Failure value (a non-strict « holds »): both twins hold each other, both are dropped, and the reason is « end-touches-nothing ».
    expect(r.abstentions["end-touches-several"]).toBe(1);
    expect(r.abstentions["end-touches-nothing"]).toBeUndefined();
  });

  it("a declared shape that is the frame around another touched shape is still the declared one (USA deck: 10 false disagreements)", () => {
    const els = [node("a", 0, 0), node("frame", 295, -10, 300, 80), node("inner", 300, 0)];
    const c: DiagramConnector = {
      from: "a",
      to: "frame",
      ends: { x1: 104, y1: 20, x2: 297, y2: 20 },
      startArrow: false,
      endArrow: true,
      visible: true,
    };
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [c], tolerances: TOL });
    expect(r.paragraphs).toEqual([[0, 1]]);
    expect(r.rest).toEqual([2]);
  });

  it("the declared reading and the geometry must agree, else abstain", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 600, 0)];
    const disagree: DiagramConnector = {
      from: "a",
      to: "c",
      ends: { x1: 104, y1: 20, x2: 296, y2: 20 },
      startArrow: false,
      endArrow: true,
      visible: true,
    };
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [disagree], tolerances: TOL });
    expect(r.paragraphs).toEqual([]);
    expect(r.abstentions["declared-and-geometry-disagree"]).toBe(1);
  });
});

describe("orderDiagram: labels", () => {
  const base = [node("a", 0, 0), node("b", 400, 0)];
  const arrow = line(100, 20, 400, 20);

  it("a label next to one connector's midpoint goes between source and target", () => {
    const label: DiagramElement = { box: box(190, -30, 120, 20), node: true };
    const r = orderDiagram({
      elements: [base[0]!, label, base[1]!],
      obstacles: [],
      connectors: [arrow],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 1, 2]]);
  });

  it("a label equally near two parallel connectors is NOT attached (stays in the rest)", () => {
    const els = [
      node("a", 0, 0),
      node("b", 400, 0),
      node("c", 0, 100),
      node("d", 400, 100),
      { box: box(190, 55, 120, 30), node: true } as DiagramElement,
    ];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [line(100, 20, 400, 20), line(100, 120, 400, 120)],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([
      [0, 1],
      [2, 3],
    ]);
    expect(r.rest).toEqual([4]);
    expect(r.abstentions["label-ambiguous"]).toBe(1);
  });

  it("a shape flagged notLabel (turned, or a block arrow with its own text) is never a label, even next to the connector's midpoint (05-Groupes-tournes2, « texte dans le groupe »)", () => {
    const label: DiagramElement = { box: box(190, -30, 120, 20), node: true, notLabel: true };
    const r = orderDiagram({
      elements: [base[0]!, label, base[1]!],
      obstacles: [],
      connectors: [arrow],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 2]]);
    expect(r.rest).toEqual([1]);
  });

  it("a text far from every connector is not a label", () => {
    const far: DiagramElement = { box: box(190, 600, 120, 20), node: true };
    const r = orderDiagram({
      elements: [base[0]!, far, base[1]!],
      obstacles: [],
      connectors: [arrow],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 2]]);
    expect(r.rest).toEqual([1]);
  });

  it("a label whose nearest connector is in a component that is not a tree is not attached (it stays in the rest)", () => {
    const els = [
      node("a", 0, 0),
      node("b", 400, 0),
      node("c", 400, 100),
      { box: box(190, -30, 120, 20), node: true } as DiagramElement,
    ];
    const r = orderDiagram({
      elements: els,
      obstacles: [],
      connectors: [declared("a", "b"), declared("a", "c"), declared("c", "b"), { ...line(100, 20, 400, 20) }],
      tolerances: TOL,
    });
    expect(r.paragraphs).toEqual([[0, 2, 1]]);
    expect(r.rest).toEqual([3]);
  });
});

describe("orderDiagram: never loses nor repeats an element", () => {
  it("elements that are not nodes (tables, pictures) and elements without a box stay in the rest, in order", () => {
    const els: DiagramElement[] = [
      { node: false, box: box(0, 500) },
      node("a", 0, 0),
      { node: true },
      node("b", 300, 0),
    ];
    const r = orderDiagram({ elements: els, obstacles: [], connectors: [declared("a", "b")], tolerances: TOL });
    expect(r.paragraphs).toEqual([[1, 3]]);
    expect(r.rest).toEqual([0, 2]);
  });
});

// Review of 01/10/2026 (security): the emission walk was recursive, once per
// node along a chain, and filtered every edge per node — a 5 000-element
// chain overflowed the stack (the document failed on every attempt), 12 000
// took 11.6 s. The walk is iterative over a source → edges map.
describe("orderDiagram: long chains", () => {
  const chain = (n: number) => {
    const elements: DiagramElement[] = Array.from({ length: n }, (_, i) => ({
      id: `n${i}`,
      box: box(i * 200, 0),
      node: true,
    }));
    const connectors = Array.from({ length: n - 1 }, (_, i) => declared(`n${i}`, `n${i + 1}`));
    return { elements, connectors };
  };
  for (const n of [5_000, 12_000]) {
    it(`a chain of ${n} elements neither throws nor takes seconds, and stays one paragraph in order`, () => {
      const { elements, connectors } = chain(n);
      const t0 = Date.now();
      const r = orderDiagram({ elements, obstacles: [], connectors, tolerances: TOL });
      // Failure value: RangeError (stack overflow) at 5 000; 11.6 s at 12 000.
      expect(r.paragraphs).toHaveLength(1);
      expect(r.paragraphs[0]).toEqual(Array.from({ length: n }, (_, i) => i));
      expect(r.rest).toEqual([]);
      expect(Date.now() - t0).toBeLessThan(3_000);
    }, 30_000);
  }

  it("a star of 12 000 leaves keeps the first child in the parent's paragraph and opens one paragraph per other child", () => {
    const n = 12_000;
    const elements: DiagramElement[] = Array.from({ length: n }, (_, i) => ({
      id: `n${i}`,
      box: box(i * 200, 0),
      node: true,
    }));
    const connectors = Array.from({ length: n - 1 }, (_, i) => declared("n0", `n${i + 1}`));
    const t0 = Date.now();
    const r = orderDiagram({ elements, obstacles: [], connectors, tolerances: TOL });
    expect(r.paragraphs).toHaveLength(n - 1);
    expect(r.paragraphs[0]).toEqual([0, 1]);
    expect(r.paragraphs[n - 2]).toEqual([n - 1]);
    expect(Date.now() - t0).toBeLessThan(3_000);
  }, 30_000);
});
