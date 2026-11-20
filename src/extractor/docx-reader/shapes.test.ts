// A minimal unit test on a crafted snippet.
// Covers the silent loss fixed here: a `pic:pic` or `wpg:graphicFrame` inside
// a `wpg:wgp` group (schema-legal, ECMA-376/[MS-DOCX] 2.5.78,
// `WordprocessingGroup`'s `CT_WordprocessingGroup` choice — Microsoft Learn,
// 29/09/2026) was walked by nothing and counted by nothing.

import { describe, expect, it } from "vitest";

import { MAX_SLIDE_ELEMENTS, PAGE_SIZED_LIMITS } from "../page-caps.test.support";
import { drawShapeBlocks, type DrawingImageRegistrar } from "./shapes";
import { parseXml, type XNode } from "../ooxml/xml";

const NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';

const WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

const filledRect =
  "<wps:wsp><wps:spPr>" +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
  "</wps:spPr></wps:wsp>";

const nestedPicture =
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="p"/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></pic:spPr></pic:pic>';

const nestedGraphicFrame =
  '<wpg:graphicFrame><wpg:cNvPr id="3" name="gf"/><wpg:cNvFrPr/>' +
  '<wpg:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></wpg:xfrm>' +
  '<a:graphic><a:graphicData uri="urn:test:unrelated"/></a:graphic></wpg:graphicFrame>';

const drawing = (groupMembers: string): XNode => {
  const xml =
    `<w:drawing ${NS}><wp:inline><wp:extent cx="914400" cy="914400"/>` +
    `<a:graphic><a:graphicData uri="${WPG_URI}"><wpg:wgp>${groupMembers}</wpg:wgp></a:graphicData></a:graphic>` +
    "</wp:inline></w:drawing>";
  const node = parseXml(xml);
  if (!node) throw new Error("test snippet failed to parse");
  return node;
};

const registrar: DrawingImageRegistrar = (bytes, mimeType) => ({ sha256: `sha-${bytes.length}`, mimeType });

describe("drawShapeBlocks — caps (review 30/09/2026, finding 20)", () => {
  it("a drawing over MAX_SLIDE_ELEMENTS shapes draws nothing and counts them all as not drawn", () => {
    let calls = 0;
    const counting: DrawingImageRegistrar = (bytes, mimeType) => (calls++, registrar(bytes, mimeType));
    const result = drawShapeBlocks(drawing(filledRect.repeat(MAX_SLIDE_ELEMENTS + 1)), counting, PAGE_SIZED_LIMITS);
    expect(result).toEqual({ blocks: [], shapesNotDrawn: MAX_SLIDE_ELEMENTS + 1 });
    expect(calls).toBe(0);
  });

  it("a drawing at the cap is still drawn", () => {
    const result = drawShapeBlocks(drawing(filledRect.repeat(MAX_SLIDE_ELEMENTS)), registrar, PAGE_SIZED_LIMITS);
    expect(result.blocks).toHaveLength(1);
  });

  it("the cap is a parameter (ReaderLimits.slideElements): lowered, a smaller drawing is not drawn; absent, the product's", () => {
    // FAILS WITH (limits ignored): expected { blocks: [ { …(2) } ], … } to deeply equal { blocks: [], shapesNotDrawn: 3 }.
    expect(drawShapeBlocks(drawing(filledRect.repeat(3)), registrar, { slideElements: 2 })).toEqual({
      blocks: [],
      shapesNotDrawn: 3,
    });
    expect(drawShapeBlocks(drawing(filledRect.repeat(2)), registrar, { slideElements: 2 }).blocks).toHaveLength(1);
    expect(drawShapeBlocks(drawing(filledRect.repeat(3)), registrar, {}).blocks).toHaveLength(1);
  });

  it("a registrar that refuses the picture (byte budget spent) leaves the shapes not drawn, no block", () => {
    const refusing: DrawingImageRegistrar = () => undefined;
    const result = drawShapeBlocks(drawing(filledRect.repeat(3)), refusing, PAGE_SIZED_LIMITS);
    expect(result).toEqual({ blocks: [], shapesNotDrawn: 3 });
  });
});

describe("drawShapeBlocks — pic:pic and wpg:graphicFrame inside a group", () => {
  it("counts a pic:pic alongside a drawn shape, instead of losing it silently", () => {
    const result = drawShapeBlocks(drawing(filledRect + nestedPicture), registrar, PAGE_SIZED_LIMITS);
    expect(result.blocks).toHaveLength(1); // the rectangle is drawn
    expect(result.shapesNotDrawn).toBe(1); // the picture is not
  });

  it("counts a wpg:graphicFrame alongside a drawn shape, instead of losing it silently", () => {
    const result = drawShapeBlocks(drawing(filledRect + nestedGraphicFrame), registrar, PAGE_SIZED_LIMITS);
    expect(result.blocks).toHaveLength(1);
    expect(result.shapesNotDrawn).toBe(1);
  });

  it("counts both together, and a lone pic:pic in a group (no other shape) still counts even though nothing is drawn", () => {
    const both = drawShapeBlocks(
      drawing(filledRect + nestedPicture + nestedGraphicFrame),
      registrar,
      PAGE_SIZED_LIMITS,
    );
    expect(both.shapesNotDrawn).toBe(2);

    // The exact bug: today `shapes.length === 0` short-circuits to
    // `shapesNotDrawn: 0`, so a group holding only a `pic:pic` (no wps:wsp,
    // no wpg:wgp) never surfaces the loss at all.
    const onlyPicture = drawShapeBlocks(drawing(nestedPicture), registrar, PAGE_SIZED_LIMITS);
    expect(onlyPicture.blocks).toEqual([]);
    expect(onlyPicture.shapesNotDrawn).toBe(1);
  });
});

// Exhaustive audit of 07/10/2026 (F2.1): ONE rule for « a Word connector's line shows ». With no `a:ln` written, Word's own connector
// shows its line through `wps:style/a:lnRef idx > 0` (Doc1.docx); `readDrawingDiagram` read it, the drawn picture did not.
describe("drawShapeBlocks — a Word connector styled by its theme line is drawn", () => {
  const connector = (lnRefIdx: number): string =>
    "<wps:wsp><wps:cNvCnPr/><wps:spPr>" +
    '<a:xfrm><a:off x="0" y="200"/><a:ext cx="300" cy="0"/></a:xfrm>' +
    '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>' +
    "</wps:spPr>" +
    `<wps:style><a:lnRef idx="${lnRefIdx}"><a:schemeClr val="accent1"/></a:lnRef></wps:style></wps:wsp>`;
  const svgOf = (members: string): string => {
    let svg = "";
    drawShapeBlocks(
      drawing(members),
      (bytes, mimeType) => ((svg = new TextDecoder().decode(bytes)), registrar(bytes, mimeType)),
      PAGE_SIZED_LIMITS,
    );
    return svg;
  };

  it("🔴 a connector with lnRef idx 1 and no a:ln draws a line that idx 0 does not", () => {
    expect(svgOf(filledRect + connector(1))).not.toBe(svgOf(filledRect + connector(0)));
  });
});

describe("drawShapeBlocks — a connector the drawer cannot trace does not make a picture", () => {
  // A visible straight connector that writes no `a:xfrm`: the reader gives it no ends, so the drawer traces no line.
  const endless =
    "<wps:wsp><wps:cNvCnPr/><wps:spPr>" +
    '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>' +
    '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
    "</wps:spPr></wps:wsp>";

  it("alone, it gives no picture: nothing non-textual would be in it", () => {
    let registered = 0;
    const counting: DrawingImageRegistrar = (bytes, mimeType) => (registered++, registrar(bytes, mimeType));
    // Failure value: one image block (a picture with no line in it) and 1 registration.
    expect(drawShapeBlocks(drawing(endless), counting, PAGE_SIZED_LIMITS).blocks).toEqual([]);
    expect(registered).toBe(0);
  });

  it("beside a drawn shape, the picture is made and the connector counts as not drawn", () => {
    const result = drawShapeBlocks(drawing(filledRect + endless), registrar, PAGE_SIZED_LIMITS);
    expect(result.blocks).toHaveLength(1);
    expect(result.shapesNotDrawn).toBe(1);
  });
});

describe("drawShapeBlocks — a shape without a box does not make a picture", () => {
  // A filled rectangle that writes no `a:xfrm`: the reader gives it no box, so the drawer leaves it out (`no-box`).
  const boxless =
    "<wps:wsp><wps:spPr>" +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
    '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
    "</wps:spPr></wps:wsp>";

  it("alone, it gives no picture: nothing would be in it", () => {
    let registered = 0;
    const counting: DrawingImageRegistrar = (bytes, mimeType) => (registered++, registrar(bytes, mimeType));
    // Failure value: one image block (a picture with no shape in it) and 1 registration.
    expect(drawShapeBlocks(drawing(boxless), counting, PAGE_SIZED_LIMITS).blocks).toEqual([]);
    expect(registered).toBe(0);
  });

  it("beside a drawn shape, the picture is made and the boxless shape counts as not drawn", () => {
    const result = drawShapeBlocks(drawing(filledRect + boxless), registrar, PAGE_SIZED_LIMITS);
    expect(result.blocks).toHaveLength(1);
    expect(result.shapesNotDrawn).toBe(1);
  });
});

// Exhaustive audit of 07/10/2026 (P1-F2.2): the docx copy of the shape reader never set `isTurnedByGroup`, so a text box in a turned
// group was drawn upright in the SVG (its box and text at the wrong angle), where PowerPoint's reader leaves it out (« turned »).
describe("drawShapeBlocks — a text box in a turned group is left out of the picture, as in PowerPoint", () => {
  const textBox =
    "<wps:wsp><wps:spPr>" +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
    '</wps:spPr><wps:txbx><w:txbxContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Tourné</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>';
  const turnedGroup = (rot: number): XNode => {
    const xml =
      `<w:drawing ${NS}><wp:inline><wp:extent cx="1828800" cy="1828800"/>` +
      `<a:graphic><a:graphicData uri="${WPG_URI}"><wpg:wgp><wpg:grpSpPr><a:xfrm rot="${rot}"><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/>` +
      '<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/></a:xfrm></wpg:grpSpPr>' +
      `${textBox}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const node = parseXml(xml);
    if (!node) throw new Error("test snippet failed to parse");
    return node;
  };
  const svgOf = (node: XNode): string => {
    let svg = "";
    drawShapeBlocks(
      node,
      (bytes, mimeType) => ((svg = new TextDecoder().decode(bytes)), registrar(bytes, mimeType)),
      PAGE_SIZED_LIMITS,
    );
    return svg;
  };

  // The picture of a Word drawing carries the boxes; the text itself is published as words. Turned, the box is left out.
  it("🔴 turned by its group: its box is not drawn", () => {
    expect(svgOf(turnedGroup(0)), "precondition: upright, the box is drawn").toContain('stroke="#000"');
    expect(svgOf(turnedGroup(5400000))).not.toContain('stroke="#000"');
  });
});
