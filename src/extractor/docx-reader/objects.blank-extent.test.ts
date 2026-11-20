import { describe, expect, it } from "vitest";

import { isOffPage } from "./objects";
import { parseXml } from "../ooxml/xml";

// Decision n° 12 (12/10/2026): a blank `wp:extent` size is unknown, not 0 — only what is PROVABLY off the page is dropped.
// Run: npx vitest run --project=main src/extractor/docx-reader/objects.blank-extent.test.ts

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const anchor = (cx: string, cy: string): string =>
  `<w:drawing ${NS}><wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>99999999</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/></wp:anchor></w:drawing>`;
const PAGE = { widthEmu: 7_560_000, heightEmu: 10_692_000 };
const drawing = (cx: string, cy: string) => {
  const root = parseXml(anchor(cx, cy));
  if (!root) throw new Error("fixture does not parse");
  return root;
};

describe("isOffPage with a blank extent", () => {
  it("a written extent far right of the page is off the page (control)", () => {
    expect(isOffPage(drawing("100", "100"), PAGE)).toBe(true);
  });
  it("a blank cx/cy is an unknown size: the drawing is kept", () => {
    expect(isOffPage(drawing("", ""), PAGE)).toBe(false);
  });
});
