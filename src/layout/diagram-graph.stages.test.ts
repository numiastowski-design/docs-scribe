import { describe, expect, it } from "vitest";

import {
  type DiagramConnector,
  type DiagramElement,
  type DiagramInput,
  type DiagramTolerances,
  orderDiagram,
} from "./diagram-graph";

// Characterisation of the stages `orderDiagram` is cut into (ends, relations, components, classes, labels, emission): each
// counter and each output field the other test files leave unpinned. Written green on the code before the cut.
const TOL: DiagramTolerances = { endTouch: 10, labelReach: 50, ambiguity: 1.5 };

const node = (id: string, x: number, y: number, cx = 100, cy = 40): DiagramElement => ({
  id,
  box: { x, y, cx, cy },
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
const order = (elements: DiagramElement[], connectors: DiagramConnector[], obstacles: DiagramInput["obstacles"] = []) =>
  orderDiagram({ elements, obstacles, connectors, tolerances: TOL });

describe("orderDiagram stages: every abstention is counted once, at its own stage", () => {
  it("a connector with no ends and no glue says no-ends; a loop is a loop", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    expect(order(els, [{ startArrow: true, endArrow: true, visible: true }]).abstentions).toEqual({ "no-ends": 1 });
    expect(order(els, [{ from: "a", startArrow: true, endArrow: true, visible: true }]).abstentions).toEqual({
      "no-ends": 1,
    });
    const loop = order(els, [declared("a", "a")]);
    expect(loop.abstentions).toEqual({ loop: 1 });
    expect(loop.relations).toEqual({ resolved: 0, total: 0, declared: 0, geometry: 0, mixed: 0 });
    expect(loop.paragraphs).toEqual([]);
    expect(loop.rest).toEqual([0, 1]);
  });

  it("an unknown id, two shapes of one id, and a start that fails hide the end's reading (the start is judged first)", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("b", 300, 100)];
    expect(order(els, [declared("a", "zz")]).abstentions).toEqual({ "declared-unknown": 1 });
    expect(order(els, [declared("a", "b")]).abstentions).toEqual({ "declared-unknown": 1 });
    expect(order(els, [declared("zz", "q")]).abstentions).toEqual({ "declared-unknown": 1 });
    expect(order(els, [line(1000, 1000, 100, 20)]).abstentions).toEqual({ "end-touches-nothing": 1 });
  });

  it("a declared end that is an obstacle, or a node without text, abstains as a shape without text", () => {
    const els = [node("a", 0, 0), { ...node("t", 300, 0), node: false }];
    const obstacles = [{ id: "o", box: { x: 0, y: 200, cx: 100, cy: 40 } }];
    expect(order(els, [declared("a", "o")], obstacles).abstentions).toEqual({ "end-on-shape-without-text": 1 });
    expect(order(els, [declared("t", "a")], obstacles).abstentions).toEqual({ "end-on-shape-without-text": 1 });
    expect(order(els, [declared("a", "t")], obstacles).abstentions).toEqual({ "end-on-shape-without-text": 1 });
  });

  it("an obstacle with a box is touched by a line end like any shape, and it is no text", () => {
    const els = [node("a", 0, 0)];
    const obstacles = [{ id: "o", box: { x: 300, y: 0, cx: 100, cy: 40 } }];
    expect(order(els, [line(100, 20, 300, 20)], obstacles).abstentions).toEqual({ "end-on-shape-without-text": 1 });
  });

  it("the first failing stage wins: invisible, then ends, then loop, then the arrow head", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    expect(order(els, [{ ...declared("a", "a"), visible: false }]).abstentions).toEqual({ invisible: 1 });
    expect(order(els, [declared("zz", "zz")]).abstentions).toEqual({ "declared-unknown": 1 });
    expect(order(els, [declared("a", "a", "none")]).abstentions).toEqual({ loop: 1 });
    expect(order(els, [declared("a", "b", "none")]).abstentions).toEqual({ "no-arrowhead": 1 });
  });

  it("the keys of the abstention map follow the order the stages meet them: connectors, then cycles, then labels", () => {
    const els = [
      node("A", 0, 0),
      node("B", 300, 0),
      node("C", 0, 200),
      node("D", 300, 200),
      node("E", 0, 400),
      node("F", 300, 400),
      node("L", 150, 700),
    ];
    const r = order(els, [
      { ...declared("A", "B"), visible: false },
      declared("A", "A"),
      declared("C", "D", "none"),
      declared("A", "B"),
      declared("B", "A"),
      { ...declared("E", "F"), ...line(100, 420, 300, 420) },
    ]);
    expect(JSON.stringify(r.abstentions)).toBe(
      '{"invisible":1,"loop":1,"no-arrowhead":1,"has-cycle":1,"label-too-far":1}',
    );
    expect(r.relations).toEqual({ resolved: 3, total: 1, declared: 1, geometry: 0, mixed: 0 });
    expect(r.paragraphs).toEqual([[4, 5]]);
    expect(r.rest).toEqual([0, 1, 2, 3, 6]);
    expect(r.links).toEqual([
      { target: 0, sources: [1] },
      { target: 1, sources: [0] },
    ]);
  });
});

describe("orderDiagram stages: how a relation's ends were found", () => {
  it("declared, geometry and mixed are counted apart, and only for the relations kept", () => {
    const els = [
      node("a", 0, 0),
      node("b", 300, 0),
      node("c", 0, 200),
      node("d", 300, 200),
      node("e", 0, 400),
      node("f", 300, 400),
      node("g", 0, 600),
      node("h", 300, 600),
    ];
    const r = order(els, [
      declared("a", "b"),
      line(100, 220, 300, 220),
      { from: "e", ends: { x1: 100, y1: 420, x2: 300, y2: 420 }, startArrow: false, endArrow: true, visible: true },
      { to: "h", ends: { x1: 100, y1: 620, x2: 300, y2: 620 }, startArrow: false, endArrow: true, visible: true },
    ]);
    expect(r.relations).toEqual({ resolved: 4, total: 4, declared: 1, geometry: 1, mixed: 2 });
    expect(r.edges).toEqual([
      { source: 0, target: 1, labels: [] },
      { source: 2, target: 3, labels: [] },
      { source: 4, target: 5, labels: [] },
      { source: 6, target: 7, labels: [] },
    ]);
  });

  it("a relation of a component that is no tree is resolved but not kept", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 600, 0)];
    const r = order(els, [declared("a", "c"), declared("b", "c")]);
    expect(r.relations).toEqual({ resolved: 2, total: 0, declared: 0, geometry: 0, mixed: 0 });
    expect(r.edges).toEqual([]);
    expect(r.paragraphs).toEqual([[0, 1, 2]]);
  });
});

describe("orderDiagram stages: labels", () => {
  it("two labels on one connector go between its ends, in element order, and both are counted", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("l1", 150, 50), node("l2", 180, 55)];
    const r = order(els, [line(100, 20, 300, 20)]);
    expect(r.paragraphs).toEqual([[0, 2, 3, 1]]);
    expect(r.edges).toEqual([{ source: 0, target: 1, labels: [2, 3] }]);
    expect(r.labelsAttached).toBe(2);
    expect(r.rest).toEqual([]);
    expect(r.abstentions).toEqual({});
  });

  it("the labels of a further child's connector open its paragraph; those of the first continue the parent's", () => {
    const els = [node("d", 200, 0), node("s", 0, 200), node("e", 400, 200), node("L1", 100, 100), node("L2", 300, 100)];
    const r = order(els, [line(250, 40, 50, 200), line(250, 40, 450, 200)]);
    expect(r.paragraphs).toEqual([
      [0, 3, 1],
      [4, 2],
    ]);
    expect(r.edges).toEqual([
      { source: 0, target: 1, labels: [3] },
      { source: 0, target: 2, labels: [4] },
    ]);
    expect(r.labelsAttached).toBe(2);
  });

  it("a text box that holds another shape is a frame: counted, never attached", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("F", -50, -50, 600, 300)];
    const r = order(els, [declared("a", "b")]);
    expect(r.abstentions).toEqual({ "label-is-a-frame": 1 });
    expect(r.labelsAttached).toBe(0);
    expect(r.paragraphs).toEqual([[0, 1]]);
    expect(r.rest).toEqual([2]);
  });

  it("a text far from the nearest connector's midpoint is counted as too far", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("l", 150, 100)];
    const r = order(els, [line(100, 20, 300, 20)]);
    expect(r.abstentions).toEqual({ "label-too-far": 1 });
    expect(r.rest).toEqual([2]);
  });

  it("no label is looked for when no relation is kept, and a text with no connector at all stays silent", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("l", 150, 50)];
    expect(order(els, [declared("a", "b", "none")]).abstentions).toEqual({ "no-arrowhead": 1 });
    expect(order(els, []).abstentions).toEqual({});
    const two = [node("a", 0, 0), node("b", 300, 0), node("c", 600, 0), node("l", 150, 50)];
    expect(order(two, [declared("a", "b"), declared("b", "a")]).abstentions).toEqual({ "has-cycle": 1 });
  });

  it("a label whose nearest connector sits in a component that is not kept is counted, and stays in the rest", () => {
    const els = [
      node("a", 0, 0),
      node("b", 300, 0),
      node("c", 150, 200),
      node("x", 0, 600),
      node("y", 300, 600),
      node("l", 150, 50),
    ];
    const r = order(els, [
      { ...declared("a", "b"), ...line(100, 20, 300, 20) },
      { ...declared("b", "c"), ...line(350, 40, 250, 200) },
      { ...declared("c", "a"), ...line(150, 200, 50, 40) },
      declared("x", "y"),
    ]);
    expect(JSON.stringify(r.abstentions)).toBe('{"has-cycle":1,"label-on-unplaced-connector":1}');
    expect(r.labelsAttached).toBe(0);
    expect(r.paragraphs).toEqual([[3, 4]]);
    expect(r.rest).toEqual([0, 1, 2, 5]);
  });

  it("the nearest connector is judged against the others by the ambiguity ratio, then by its component", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 0, 400), node("d", 300, 400), node("l", 150, 50)];
    const kept = order(els, [{ ...declared("a", "b"), ...line(100, 20, 300, 20) }, declared("c", "d")]);
    expect(kept.edges[0]).toEqual({ source: 0, target: 1, labels: [4] });
    expect(kept.abstentions).toEqual({});
    const unplaced = order(els, [
      { ...declared("a", "b"), ...line(100, 20, 300, 20) },
      { ...declared("b", "a"), ...line(300, 20, 100, 20) },
      { ...declared("c", "d"), ...line(100, 420, 300, 420) },
    ]);
    expect(unplaced.labelsAttached).toBe(0);
  });
});

describe("orderDiagram stages: boundaries and ties", () => {
  it("a line end exactly at the touch tolerance touches the outline", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    expect(order(els, [line(110, 20, 300, 20)]).paragraphs).toEqual([[0, 1]]);
    expect(order(els, [line(111, 20, 300, 20)]).abstentions).toEqual({ "end-touches-nothing": 1 });
  });

  it("a connector with a head at both ends runs from its start to its end", () => {
    const els = [node("a", 0, 0), node("b", 300, 0)];
    const both: DiagramConnector = { from: "a", to: "b", startArrow: true, endArrow: true, visible: true };
    expect(order(els, [both]).edges).toEqual([{ source: 0, target: 1, labels: [] }]);
    expect(order(els, [{ ...both, startArrow: false }]).edges).toEqual([{ source: 0, target: 1, labels: [] }]);
    expect(order(els, [{ ...both, endArrow: false }]).edges).toEqual([{ source: 1, target: 0, labels: [] }]);
  });

  it("components follow the reading order of their LOWEST element, whatever the order of the connectors", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 0, 200), node("d", 300, 200)];
    const r = order(els, [declared("d", "a"), declared("b", "c")]);
    expect(r.paragraphs).toEqual([
      [3, 0],
      [1, 2],
    ]);
  });

  it("a label exactly at the reach is attached, one a unit farther is not", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("l", 150, 70)];
    expect(order(els, [line(100, 20, 300, 20)]).labelsAttached).toBe(1);
    expect(order([...els.slice(0, 2), node("l", 150, 71)], [line(100, 20, 300, 20)]).abstentions).toEqual({
      "label-too-far": 1,
    });
  });

  it("a label that holds the midpoints of two connectors is ambiguous (distance 0 each)", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 150, -200), node("d", 150, 200), node("l", 150, 0)];
    const r = order(els, [line(100, 20, 300, 20), line(200, -160, 200, 200)]);
    expect(r.relations.total).toBe(2);
    expect(r.abstentions).toEqual({ "label-ambiguous": 1 });
    expect(r.labelsAttached).toBe(0);
  });

  it("the other connector must be at least `ambiguity` times farther: exactly 1.5 times is enough", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 0, 90), node("d", 300, 90), node("l", 150, 40)];
    const r = order(els, [line(100, 20, 300, 20), line(100, 110, 300, 110)]);
    expect(r.edges).toEqual([
      { source: 0, target: 1, labels: [4] },
      { source: 2, target: 3, labels: [] },
    ]);
    expect(r.abstentions).toEqual({});
  });

  it("an invisible connector is no candidate for a label", () => {
    const els = [node("a", 0, 0), node("b", 300, 0), node("c", 0, 400), node("d", 300, 400), node("l", 150, 440)];
    const r = order(els, [line(100, 20, 300, 20), { ...line(100, 420, 300, 420), visible: false }]);
    // c, d and l are all texts that end no relation, and every one is far from the one visible connector.
    expect(r.abstentions).toEqual({ invisible: 1, "label-too-far": 3 });
  });
});
