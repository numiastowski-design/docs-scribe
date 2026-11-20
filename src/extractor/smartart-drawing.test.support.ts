/**
 * Builders for the SmartArt-drawn-as-SVG tests: tiny Word and PowerPoint packages
 * whose SmartArt parts (data model + `dsp:drawing`) the test writes itself, so a test exercises ONE property of the
 * drawing (a rotation, a colour modifier, a missing part) on top of the real fixtures.
 * Pairing (the file's own): a frame names its data model (`dgm:relIds r:dm`), the model names its drawing
 * (`dsp:dataModelExt relId`), both relationships live in the part that holds the frame ([MS-ODRAWXML] §2.3.6, §2.5.44).
 * NOT here: any assertion (the tests), the code under test (`extractor/smartart-drawing.ts`).
 * A TWIN, on purpose: the application's private repository builds the same packages for its tests of the published
 * page (`src/readers/smartart-drawing.test.support.ts`), because a public repository cannot lend it a test support.
 * The two are NOT the same text and no test holds them equal: a fix to one is carried to the other by hand.
 * Called by: `smartart-drawing.test.ts`, the two tests of the real files (`smartart-drawing.fixtures.test.ts`,
 * `smartart-drawing.scale.fixtures.test.ts`) and the test of the second reading
 * (`pipeline/source-reading-smartart.drawing.test.ts`).
 */
import JSZip from "jszip";

import type { IRBlock } from "doc-ir";
import type { RawImage } from "doc-ir/internal";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const DRAWING_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const TYPES =
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>';
export const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

export const EMU_PER_PT = 12_700;

/** One `dsp:sp` of a SmartArt drawing, as Office writes it (geometry, optional rotations, one paragraph per text). */
export interface DspShape {
  /** A preset geometry (`rect`, `hexagon`, `chord`, `star5`…) or `custGeom`. */
  geom?: string;
  /** Box in EMU, in the frame's own space (origin: the frame's top-left corner). */
  box?: [x: number, y: number, cx: number, cy: number];
  /** `a:xfrm rot`, 60 000ths of a degree. */
  rot?: number;
  /** `dsp:txXfrm` is written; `txRot` is its `rot` (absent: none written). */
  txXfrm?: boolean;
  txRot?: number;
  /** The `dsp:txXfrm` rectangle when it is not the shape's own (EMU, the shape's own space before its rotation). */
  txBox?: [x: number, y: number, cx: number, cy: number];
  /** `a:bodyPr` attributes: `vert` (default `horz`), `upright`, `rot`. */
  vert?: string;
  upright?: boolean;
  bodyRot?: number;
  /** `a:xfrm flipV`. */
  flipV?: boolean;
  /** One paragraph per entry; none: a shape without text. */
  text?: string[];
  /** Raw fill element(s) of `dsp:spPr`; default a visible exact grey (`VISIBLE_FILL`): a shape that writes neither fill nor line draws
   *  nothing at all (`extractor/slide-svg.ts#drawOutline`), and the tests that look for an outline need one. */
  fill?: string;
  /** Colour of the text runs (`FFFFFF`); default `000000`, written explicitly. */
  textColor?: string;
}

const VISIBLE_FILL = '<a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>';

/** The `a:bodyPr` of a shape's text: `vert`, and `upright` / `rot` when the shape writes them. */
const bodyPrXml = (s: DspShape): string => {
  const upright = s.upright === true ? ' upright="1"' : "";
  const rot = s.bodyRot === undefined ? "" : ` rot="${s.bodyRot}"`;
  return `<a:bodyPr spcFirstLastPara="0" vert="${s.vert ?? "horz"}"${upright}${rot} wrap="square" lIns="83820" tIns="83820" rIns="83820" bIns="83820" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr>`;
};

export const dspShape = (s: DspShape): string => {
  const [x, y, cx, cy] = s.box ?? [0, 0, 1_500_000, 600_000];
  const geometry =
    s.geom === "custGeom"
      ? '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom>'
      : `<a:prstGeom prst="${s.geom ?? "rect"}"><a:avLst/></a:prstGeom>`;
  const rot = (s.rot === undefined ? "" : ` rot="${s.rot}"`) + (s.flipV === true ? ' flipV="1"' : "");
  const paragraphs = (s.text ?? [])
    .map(
      (t) =>
        `<a:p><a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:srgbClr val="${s.textColor ?? "000000"}"/></a:solidFill></a:rPr><a:t>${t}</a:t></a:r></a:p>`,
    )
    .join("");
  const body = s.text === undefined ? "" : `<dsp:txBody>${bodyPrXml(s)}<a:lstStyle/>${paragraphs}</dsp:txBody>`;
  const txRotAttr = s.txRot === undefined ? "" : ' rot="' + String(s.txRot) + '"';
  const [tx, ty, tcx, tcy] = s.txBox ?? [x, y, cx, cy];
  const txXfrm =
    s.txXfrm === true || s.txRot !== undefined || s.txBox !== undefined
      ? `<dsp:txXfrm${txRotAttr}><a:off x="${tx}" y="${ty}"/><a:ext cx="${tcx}" cy="${tcy}"/></dsp:txXfrm>`
      : "";
  return (
    `<dsp:sp modelId="{00000000-0000-0000-0000-000000000000}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>` +
    `<dsp:spPr><a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>${geometry}${s.fill ?? VISIBLE_FILL}<a:ln><a:noFill/></a:ln></dsp:spPr>` +
    `${body}${txXfrm}</dsp:sp>`
  );
};

/** A solid fill of an explicit colour followed by the four offsets Office writes at zero (neutral). */
export const ZERO_OFFSETS_FILL =
  '<a:solidFill><a:srgbClr val="FF0000"><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/></a:srgbClr></a:solidFill>';
export const fillWith = (modifier: string): string =>
  `<a:solidFill><a:srgbClr val="FF0000">${modifier}</a:srgbClr></a:solidFill>`;

const drawingXml = (shapes: readonly string[]): string =>
  `<?xml version="1.0"?><dsp:drawing xmlns:dgm="${DGM}" xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${shapes.join("")}</dsp:spTree></dsp:drawing>`;

const dataXml = (texts: readonly string[], drawingRelId: string | undefined): string => {
  const points = texts
    .map(
      (t, i) =>
        `<dgm:pt modelId="{1000000${i}}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
    )
    .join("");
  const ext =
    drawingRelId === undefined
      ? ""
      : `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${DSP}" relId="${drawingRelId}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst>`;
  return `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst><dgm:pt modelId="{00000000}" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>${points}</dgm:ptLst><dgm:cxnLst/>${ext}</dgm:dataModel>`;
};

/** How the drawing of a SmartArt can fail to be there. */
export type DrawingState = "ok" | "part-absent" | "no-extension" | "damaged";

export interface SmartArtSpec {
  /** Part number: `data{n}.xml`, `drawing{n}.xml`, relationship ids `rIdD{n}`, `rIdW{n}`. */
  n: number;
  /** Texts of the data model's points (one point each, document order). */
  model: string[];
  /** The `dsp:sp` elements; `undefined`: the drawing part is not written. */
  shapes?: string[];
  /** Default `ok`. `part-absent`: the relationship names a part the package does not hold. `no-extension`: the model
   *  has no `dsp:dataModelExt`. `damaged`: the drawing's XML is unreadable. */
  state?: DrawingState;
  /** Frame position (pptx), EMU. */
  frame?: [x: number, y: number, cx: number, cy: number];
}

interface Parts {
  files: Record<string, string>;
  rels: string;
}

const partsOf = (dir: string, spec: SmartArtSpec): Parts => {
  const state = spec.state ?? "ok";
  const data = `rIdD${spec.n}`;
  const drawing = `rIdW${spec.n}`;
  const files: Record<string, string> = {};
  let rels = `<Relationship Id="${data}" Type="${R}/diagramData" Target="diagrams/data${spec.n}.xml"/>`;
  files[`${dir}/diagrams/data${spec.n}.xml`] = dataXml(
    spec.model,
    state === "no-extension" || spec.shapes === undefined ? undefined : drawing,
  );
  if (spec.shapes !== undefined) {
    rels += `<Relationship Id="${drawing}" Type="${DRAWING_REL}" Target="diagrams/drawing${spec.n}.xml"/>`;
    if (state !== "part-absent")
      files[`${dir}/diagrams/drawing${spec.n}.xml`] =
        state === "damaged" ? '<?xml version="1.0"?><dsp:drawing <<< not xml' : drawingXml(spec.shapes);
  }
  return { files, rels };
};

const relIds = (n: number): string =>
  `<dgm:relIds xmlns:dgm="${DGM}" xmlns:r="${R}" r:dm="rIdD${n}" r:lo="rIdL${n}" r:qs="rIdQ${n}" r:cs="rIdC${n}"/>`;

/** A Word document: one paragraph of text, then one SmartArt frame per spec (own paragraph, `wp:inline` 5 486 400 × 3 200 400
 *  EMU, as Word writes it), then one paragraph of text. */
export const docxWith = async (specs: readonly SmartArtSpec[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", TYPES);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>`));
  let rels = "";
  const frames: string[] = [];
  for (const spec of specs) {
    const parts = partsOf("word", spec);
    for (const [path, xml] of Object.entries(parts.files)) zip.file(path, xml);
    rels += parts.rels;
    frames.push(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="${spec.n}" name="d${spec.n}"/><a:graphic><a:graphicData uri="${DGM}">${relIds(spec.n)}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    );
  }
  zip.file("word/_rels/document.xml.rels", RELS(rels));
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${A}" xmlns:r="${R}"><w:body>` +
      `<w:p><w:r><w:t>beforewords</w:t></w:r></w:p>${frames.join("")}<w:p><w:r><w:t>afterwords</w:t></w:r></w:p></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array" });
};

const P_NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const SLIDE_SIZE = '<p:sldSz cx="12192000" cy="6858000"/>';

/** A text box (`p:sp`) of a slide. */
export const slideBox = (
  id: number,
  text: string,
  box: [number, number, number, number] = [100_000, 100_000, 1_000_000, 300_000],
): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="b${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;

/** A straight connector: its presence makes the slide a « diagram » (`layout/slide-layout.ts`). */
export const slideConnector = (id: number): string =>
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="c${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="4000000" y="100000"/><a:ext cx="500000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`;

export const smartArtFrame = (spec: SmartArtSpec): string => {
  const [x, y, cx, cy] = spec.frame ?? [1_000_000, 1_000_000, 5_000_000, 3_000_000];
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${100 + spec.n}" name="sa${spec.n}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${DGM}">${relIds(spec.n)}</a:graphicData></a:graphic></p:graphicFrame>`
  );
};

export interface SlideSpec {
  /** SmartArts of the slide, in shape-tree order, each after the `before` shapes. */
  smartArts: SmartArtSpec[];
  /** Raw `p:sp` / `p:cxnSp` / `p:grpSp` xml written before the frames. */
  before?: string;
  /** Wraps every frame in a `p:grpSp` with an identity matrix. */
  grouped?: boolean;
}

/** A deck with a declared slide size; `slides[i]` is slide `i + 1`. */
export const pptxWith = async (
  slides: readonly SlideSpec[],
  opts: { slideSize?: boolean } = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", TYPES);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/>`));
  const slideIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}>${opts.slideSize === false ? "" : SLIDE_SIZE}<p:sldIdLst>${slideIds}</p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      slides
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  slides.forEach((slide, i) => {
    let rels = "";
    const frames = slide.smartArts.map((spec) => {
      const parts = partsOf("ppt", spec);
      for (const [path, xml] of Object.entries(parts.files)) zip.file(path, xml);
      rels += parts.rels.replaceAll('Target="diagrams/', 'Target="../diagrams/');
      return smartArtFrame(spec);
    });
    const inner = slide.grouped
      ? `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/><a:chOff x="0" y="0"/><a:chExt cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>${frames.join("")}</p:grpSp>`
      : frames.join("");
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, RELS(rels));
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${slide.before ?? ""}${inner}</p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

// ---- Reading what an extraction produced ----

const SVG = "image/svg+xml";

/** Every `image` block of an IR tree, in document order (tabs, table cells, lists… included). */
export const imageRefs = (
  blocks: unknown,
  out: { sha256: string; mimeType: string }[] = [],
): { sha256: string; mimeType: string }[] => {
  if (Array.isArray(blocks)) for (const b of blocks) imageRefs(b, out);
  else if (typeof blocks === "object" && blocks !== null) {
    const o = blocks as Record<string, unknown>;
    if (o["kind"] === "image" && typeof o["ref"] === "object" && o["ref"] !== null)
      out.push(o["ref"] as { sha256: string; mimeType: string });
    else for (const v of Object.values(o)) imageRefs(v, out);
  }
  return out;
};

/** The SVG pictures an extraction publishes, as text, in document order (one entry per image BLOCK). */
export const svgsOf = (result: { doc: { blocks: readonly IRBlock[] }; images: readonly RawImage[] }): string[] => {
  const bySha = new Map(result.images.map((i) => [i.sha256, i]));
  return imageRefs(result.doc.blocks)
    .filter((r) => r.mimeType === SVG)
    .map((r) => new TextDecoder().decode(bySha.get(r.sha256)!.bytes));
};

const unescapeXml = (s: string): string =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

/** The text an SVG draws: its `<text>` elements' character data in paint order, tspans joined, ALL whitespace removed
 *  (a wrap, a run boundary or a line break never decides a comparison). */
export const svgCompact = (svg: string): string =>
  [...svg.matchAll(/<text\b[^>]*>(.*?)<\/text>/gs)]
    // eslint-disable-next-line sonarjs/super-linear-regex -- judged safe 20/10/2026: runs on the SVG our drawer writes (every literal < is escaped)
    .map((m) => unescapeXml(m[1]!.replace(/<[^>]+>/g, "")))
    .join("")
    .replace(/\s+/g, "");

/** The text of a `dsp:drawing` XML as `svgCompact` reads it: shapes in document order, paragraphs joined, whitespace removed. */
export const drawingCompact = (xml: string): string =>
  (xml.match(/<dsp:sp\b.*?<\/dsp:sp>/gs) ?? [])
    .map((sp) => [...sp.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unescapeXml(m[1]!)).join(""))
    .join("")
    .replace(/\s+/g, "");

/** Every `<rect>` of an SVG but its white background, as `{x, y, width, height}` in points. */
export const outlineRects = (svg: string): { x: number; y: number; width: number; height: number }[] =>
  [...svg.matchAll(/<rect\b([^>]*)>/g)]
    .filter((m) => !/fill="#fff"/.test(m[1]!))
    .map((m) => {
      const num = (name: string): number =>
        Number(new RegExp(`\\b${name}="(-?[\\d.]+)"`).exec(m[1]!)?.[1] ?? Number.NaN);
      return { x: num("x"), y: num("y"), width: num("width"), height: num("height") };
    });
