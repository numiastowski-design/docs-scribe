import { describe, expect, it } from "vitest";

import { MAX_SLIDE_ELEMENTS, PAGE_SIZED_LIMITS } from "../page-caps.test.support";
import { parseXml } from "../ooxml/xml";
import { readShapes } from "./shapes";

// Characterisation of the shape walk's own bookkeeping (ranks, the element cap, hidden elements, groups), called on a tree directly:
// `readPptx` withholds a slide past the cap before it walks, so only a direct call reaches the cap branch.
const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const sp = (id: number, hidden = false): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"${hidden ? ' hidden="1"' : ""}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>t${id}</a:t></a:r></a:p></p:txBody></p:sp>`;
const grp = (id: number, members: string, hidden = false): string =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="g${id}"${hidden ? ' hidden="1"' : ""}/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${members}</p:grpSp>`;
// The element cap is the PRODUCT's, asked for here as the app asks for it: the library caps nothing by itself (09/11/2026).
const walk = (inner: string) =>
  readShapes(parseXml(`<p:spTree ${NS}>${inner}</p:spTree>`)!, {
    relationships: new Map(),
    media: new Map(),
    limits: PAGE_SIZED_LIMITS,
  });

describe("readShapes: the element cap", () => {
  it("the element at the cap is counted, never read; a group past it is still walked so the told total is true", () => {
    const inCap = Array.from({ length: MAX_SLIDE_ELEMENTS }, (_, i) => sp(i + 10)).join("");
    const result = walk(inCap + sp(1) + grp(2, sp(3) + grp(4, sp(5))) + sp(6, true));
    // 500 read; then one box, a group, its box, a group, its box (4), and a hidden box: counted, not read, and the hidden one is
    // not counted as hidden (it is not even looked at).
    expect(result.shapes).toHaveLength(MAX_SLIDE_ELEMENTS);
    expect(result.elementCount).toBe(MAX_SLIDE_ELEMENTS + 1 + 4 + 1);
    expect(result.hiddenShapes).toBe(0);
  });

  it("the last element under the cap is read", () => {
    const result = walk(Array.from({ length: MAX_SLIDE_ELEMENTS }, (_, i) => sp(i + 10)).join(""));
    expect([result.shapes.length, result.elementCount, result.shapes.at(-1)!.spTreeIndex]).toEqual([
      MAX_SLIDE_ELEMENTS,
      MAX_SLIDE_ELEMENTS,
      MAX_SLIDE_ELEMENTS - 1,
    ]);
  });
});

describe("readShapes: ranks, hidden elements and groups", () => {
  it("character data and a group's own properties take no rank", () => {
    const result = walk(
      `stray${sp(2)}more<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/><p:extLst/>${sp(3)}`,
    );
    expect(result.shapes.map((s) => s.spTreeIndex)).toEqual([0, 1]);
    expect(result.elementCount).toBe(2);
  });

  it("a hidden element is ranked and counted; a hidden group takes its members with it, which are not ranked", () => {
    const result = walk(sp(2, true) + sp(3) + grp(4, sp(5) + sp(6), true) + sp(7));
    expect(result.shapes.map((s) => [s.id, s.spTreeIndex])).toEqual([
      ["3", 1],
      ["7", 3],
    ]);
    expect(result.hiddenShapes).toBe(2);
    expect(result.elementCount).toBe(4);
  });

  it("a member of a group carries the rank of the OUTERMOST group, however deep; a group ranks before its members", () => {
    const result = walk(sp(2) + grp(3, sp(4) + grp(5, sp(6) + grp(7, sp(8)))) + sp(9));
    expect(result.shapes.map((s) => [s.id, s.spTreeIndex, s.group])).toEqual([
      ["2", 0, undefined],
      ["4", 2, 1],
      ["6", 4, 1],
      ["8", 6, 1],
      ["9", 7, undefined],
    ]);
    expect(result.elementCount).toBe(8);
  });

  it("an element this reader does not read still takes its rank", () => {
    const result = walk(
      `<p:contentPart r:id="x"/><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="2" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>${sp(3)}`,
    );
    expect(result.shapes.map((s) => [s.kind, s.spTreeIndex])).toEqual([
      ["connector", 1],
      ["text", 2],
    ]);
    expect(result.elementCount).toBe(3);
  });
});
