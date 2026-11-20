/**
 * Test support: made-up PowerPoint decks written by a test (a deck of given slide trees, text and title shapes, a
 * graphic frame, the shapes of a diagram slide, the two-slide diagram deck), and the two readers of the document a deck
 * gives (its tabs, its diagram-text sections).
 * Decides: the XML these decks are made of, and what « the tabs of a deck » means for a test (one tabs block, then at
 * most one diagram-text block, nothing else).
 * NOT here: any assertion about what the reader does with a deck (the tests), the SmartArt decks
 * (`smartart-drawing.test.support.ts`), nor the caps a test reads with (`page-caps.test.support.ts`).
 * A TWIN, on purpose: the application's private repository builds the same decks for its page tests
 * (`src/readers/pptx-deck.test.support.ts`), because a public repository cannot lend it a test support. The two are
 * NOT the same text and no test holds them equal: a fix to one is carried to the other by hand.
 * Called by: `pptx.test.ts`.
 */

import JSZip from "jszip";
import { expect } from "vitest";

import { encodePng } from "doc-ir/internal";
import type { Inline, IRBlock, IRTab } from "doc-ir";

/** One tabs block — then, at most, ONE « diagram text » block for the whole
 *  deck (29/09/2026; before, one per diagram slide — option A of
 *  27/09). Since lot A of docs/CONCEPTION-CARTE-ET-REMARQUES.md (29/09/2026)
 *  the display-cap notice is dashboard-only, never a paragraph on the page
 *  any more: anything else here fails. */
export const tabsOf = (blocks: readonly IRBlock[]): IRTab[] => {
  const [block, ...rest] = blocks;
  const diagrams = rest.filter((b) => b.kind === "diagramText");
  const others = rest.filter((b) => b.kind !== "diagramText");
  expect(others).toEqual([]);
  // Failure value (until 29/09/2026): one block per diagram slide.
  expect(diagrams.length).toBeLessThanOrEqual(1);
  for (const diagram of diagrams) {
    if (diagram.kind === "diagramText")
      for (const section of diagram.sections) expect(section.label).toMatch(/^Slide \d+$/);
  }
  if (block?.kind !== "tabs") throw new Error("expected one tabs block");
  return block.tabs;
};

/** The sections of the deck's one diagram-text block, by their slide label. */
export const diagramsOf = (blocks: readonly IRBlock[]): Map<string, Inline[]> => {
  const out = new Map<string, Inline[]>();
  for (const b of blocks.slice(1))
    if (b.kind === "diagramText") for (const section of b.sections) out.set(section.label ?? "", section.inlines);
  return out;
};

// ---------------------------------------------------------------------------
// Made-up decks

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
/** The base of every relationship type a deck declares. */
export const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

export interface MadeSlide {
  tree: string;
  hidden?: boolean;
  rels?: string;
}

/** A deck of the given slides, in order; no `p:sldSz` unless `sldSz` is
 *  given, so every slide is a flow. `files`: other parts by archive path. */
export const deckOf = async (
  slides: readonly MadeSlide[],
  { sldSz = "", files = {} }: { sldSz?: string; files?: Readonly<Record<string, Uint8Array>> } = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>',
  );
  for (const [path, bytes] of Object.entries(files)) zip.file(path, bytes);
  zip.file(
    "_rels/.rels",
    RELS(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="ppt/presentation.xml"/>`),
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join("") +
      `</p:sldIdLst>${sldSz}</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      slides
        .map(
          (_, index) =>
            `<Relationship Id="rId${index + 1}" Type="${OFFICE_REL}/slide" Target="slides/slide${index + 1}.xml"/>`,
        )
        .join(""),
    ),
  );
  slides.forEach((slide, index) => {
    zip.file(
      `ppt/slides/slide${index + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}${slide.hidden ? ' show="0"' : ""}><p:cSld><p:spTree>${slide.tree}</p:spTree></p:cSld></p:sld>`,
    );
    if (slide.rels) zip.file(`ppt/slides/_rels/slide${index + 1}.xml.rels`, RELS(slide.rels));
  });
  return zip.generateAsync({ type: "uint8array" });
};

/** A text shape of one paragraph; `ph` is its placeholder declaration, if any. */
export const textShape = (id: number, words: string, ph = ""): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
/** A text shape declared as the slide's title. */
export const titleShape = (id: number, words: string): string => textShape(id, words, '<p:ph type="title"/>');
/** A graphic frame holding `data` under the given `uri` (a table, a chart, a SmartArt, an OLE object). */
export const graphicFrame = (id: number, uri: string, data: string): string =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="f${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;

// ---------------------------------------------------------------------------
// A diagram slide: placed boxes, a connector with an arrow head, a picture, a table

/** A placement, in EMU. */
export const shapeAt = (x: number, y: number, cx: number, cy: number): string =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
/** A rectangle placed by `xfrm`, holding the paragraphs of `body`. */
export const placedBox = (id: number, xfrm: string, body: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`;
/** A straight connector with an arrow head. */
export const arrowConnector = (id: number): string =>
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="c${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
  `<p:spPr>${shapeAt(2_000_000, 500_000, 3_000_000, 0)}<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>` +
  '<a:ln><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>';
const embeddedPicture = (id: number): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr>${shapeAt(7_000_000, 4_000_000, 500_000, 500_000)}</p:spPr></p:pic>`;
const TWO_CELL_TABLE =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t6"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="0" y="3000000"/><a:ext cx="4000000" cy="500000"/></p:xfrm><a:graphic>` +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>' +
  "<a:tr><a:tc><a:txBody><a:p><a:r><a:t>c1</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>c2</a:t></a:r></a:p></a:txBody></a:tc></a:tr>" +
  "</a:tbl></a:graphicData></a:graphic></p:graphicFrame>";
/** A black PNG of one pixel: the picture the diagram slides embed. */
export const ONE_PIXEL_PNG = encodePng(1, 1, "rgb-24", [0, 0, 0])!;
const DIAGRAM_RELS =
  `<Relationship Id="rId2" Type="${OFFICE_REL}/image" Target="../media/image1.png"/>` +
  `<Relationship Id="rId5" Type="${OFFICE_REL}/hyperlink" Target="https://example.com/x?a=1&amp;b=2" TargetMode="External"/>`;

/** Two diagram slides: a title, two boxes (one with an empty paragraph, a line break and a link), a connector, a
 *  table and a picture; then a connector and a picture only, with no text at all. */
export const diagramDeck = (): Promise<Uint8Array> =>
  deckOf(
    [
      {
        tree:
          titleShape(2, "Flow") +
          placedBox(
            3,
            shapeAt(0, 0, 2_000_000, 1_000_000),
            "<a:p><a:r><a:t>Alpha</a:t></a:r></a:p><a:p/>" +
              '<a:p><a:r><a:t>R&amp;D</a:t></a:r><a:br/><a:r><a:rPr><a:hlinkClick r:id="rId5"/></a:rPr><a:t>team</a:t></a:r></a:p>',
          ) +
          arrowConnector(4) +
          placedBox(5, shapeAt(5_000_000, 0, 2_000_000, 1_000_000), "<a:p><a:r><a:t>Omega</a:t></a:r></a:p>") +
          TWO_CELL_TABLE +
          embeddedPicture(7),
        rels: DIAGRAM_RELS,
      },
      // Only a connector and a picture: no text, so no expand at all.
      { tree: arrowConnector(4) + embeddedPicture(7), rels: DIAGRAM_RELS },
    ],
    { sldSz: '<p:sldSz cx="9144000" cy="5143500"/>', files: { "ppt/media/image1.png": ONE_PIXEL_PNG } },
  );
