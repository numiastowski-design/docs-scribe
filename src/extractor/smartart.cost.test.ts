/**
 * The one case of `smartart.test.ts` that asserts a duration, on its own (18/11/2026). A wall-clock bound belongs to
 * « pdf-serial », away from the load of the parallel workers; the other cases of that file assert no duration and run
 * in « main », with the rest of the PowerPoint reader's tests, so a lot that touches `smartart.ts` runs them.
 * Measured alone: the timed call takes 130 to 300 ms against its bound of 3 000 ms.
 * Called by: nothing (a test).
 */
import { describe, expect, it } from "vitest";

import { parseXml } from "./ooxml/xml";
import { readDataModelPieces } from "./smartart";

const DGM =
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const pt = (id: string, text: string, type?: string): string =>
  `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ""}><dgm:t><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>`;
const cxn = (src: string, dest: string, srcOrd: number): string =>
  `<dgm:cxn modelId="c${src}${dest}" srcId="${src}" destId="${dest}" srcOrd="${srcOrd}" destOrd="0"/>`;
const model = (points: string[], cxns: string): string =>
  `<dgm:dataModel ${DGM}><dgm:ptLst>${points.join("")}</dgm:ptLst><dgm:cxnLst>${cxns}</dgm:cxnLst></dgm:dataModel>`;

// Review of 01/10/2026 (security): the child lists were copied on every edge (a star of 40 000 points: 14 s). Past the
// point/edge bound the document order of the points stands: every word is still published.
describe("readDataModelPieces: bounded in time on a hostile model", () => {
  it("a star of 40 000 points is read in bounded time, in document order, no word lost", () => {
    const n = 40_000;
    const points = [pt("R", "root", "doc"), ...Array.from({ length: n }, (_, i) => pt(`p${i}`, `w${i}`))];
    // srcOrd runs backwards: the tree order would reverse the leaves, document order does not.
    const cxns = Array.from({ length: n }, (_, i) => cxn("R", `p${i}`, n - i)).join("");
    const root = parseXml(model(points, cxns))!; // the parse is not what is timed
    const t0 = Date.now();
    const got = readDataModelPieces(root).map((x) => x.text);
    // Failure value: 23 s (child lists copied per edge).
    expect(Date.now() - t0).toBeLessThan(3_000);
    expect(got).toEqual(["root", ...Array.from({ length: n }, (_, i) => `w${i}`)]);
  }, 60_000);
});
