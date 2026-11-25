import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { decode as decodeJpeg, encode as encodeJpeg } from "jpeg-js";

import { parseXml, type XNode } from "../ooxml/xml";
import { encodePng } from "doc-ir/internal";
import { resolveFillRef } from "./color";
import { readGroupLevel, placePicture } from "./geometry";
import { decodePng } from "../png-decode";
import { type Size, slideItemFlags } from "../../layout/slide-elements";
import { PAGE_SIZED_LIMITS } from "../page-caps.test.support";
import {
  type PptxConnectorShape,
  type PptxParagraph,
  type PptxPictureShape,
  type PptxShape,
  type PptxSlide,
  type PptxTable,
  type PptxTextShape,
  readPptx,
} from "./index";

// Witness: `fixtures/pptx/03`.
// Every expected value below was read in the fixture's raw XML before the
// test was written (docs/FORMAT-PPTX.md [purged 27/09/2026, git history] §4).
const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../../fixtures", path)));

// Each real deck is parsed once per file and shared; no test mutates the result (checked 30/09/2026).
const parsed = new Map<string, ReturnType<typeof readPptx>>();
const readReal = (path: string): ReturnType<typeof readPptx> => {
  let p = parsed.get(path);
  if (p === undefined) {
    p = readPptx(fixture(path), { limits: PAGE_SIZED_LIMITS });
    parsed.set(path, p);
  }
  return p;
};

/** The text shapes of a slide (tables and other graphic frames left out). */
const textShapes = (slide: PptxSlide): PptxTextShape[] =>
  slide.shapes.filter((s): s is PptxTextShape => s.kind === "text");

/** Every run of a slide's text shapes, in document order. */
const runsOf = (slide: PptxSlide): string[] =>
  textShapes(slide).flatMap((s) => s.paragraphs.flatMap((p) => p.runs.map((r) => r.text)));

/** Every paragraph of a slide's text shapes, its runs joined. */
const paragraphsOf = (slide: PptxSlide): string[] =>
  textShapes(slide).flatMap((s) => s.paragraphs.map((p) => p.runs.map((r) => r.text).join("")));

/** Each text shape's paragraphs as lists of run texts. */
const textsOf = (slide: PptxSlide): string[][][] =>
  textShapes(slide).map((s) => s.paragraphs.map((p) => p.runs.map((r) => r.text)));

const slideNumbered = (slides: readonly PptxSlide[], n: number): PptxSlide => {
  const slide = slides.find((s) => s.number === n);
  if (!slide) throw new Error(`no slide ${n}`);
  return slide;
};

describe("step 0 — package, slide order, hidden slides and shapes, text", () => {
  it("03: six slides in p:sldIdLst order, slide 1's text exactly as written", async () => {
    const { slides } = await readReal("pptx/03-Presentation-Resultats-T1.pptx");
    expect(slides.map((s) => [s.number, s.partPath])).toEqual([
      [1, "ppt/slides/slide1.xml"],
      [2, "ppt/slides/slide2.xml"],
      [3, "ppt/slides/slide3.xml"],
      [4, "ppt/slides/slide4.xml"],
      [5, "ppt/slides/slide5.xml"],
      [6, "ppt/slides/slide6.xml"],
    ]);
    expect(paragraphsOf(slideNumbered(slides, 1))).toEqual([
      "Résultats T1 2026",
      "Atria Solutions — Comité de direction",
    ]);
  });

  it("a file that is not a zip archive fails with our own message", async () => {
    await expect(readPptx(new Uint8Array([1, 2, 3, 4]), { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid PowerPoint presentation",
    });
  });
});

// ---------------------------------------------------------------------------
// A synthetic package, for what no witness exercises: a numeric character
// reference, a hidden slide, a hidden shape, a group, a slide order that is
// not the part-name order. Minimal on purpose: only the parts OPC requires.

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A text shape; `ph` is the attribute list of its `p:ph`, when it has one. */
const shape = (id: number, body: string, hidden = false, ph?: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"${hidden ? ' hidden="1"' : ""}/><p:cNvSpPr/>` +
  `<p:nvPr>${ph === undefined ? "" : `<p:ph ${ph}/>`}</p:nvPr></p:nvSpPr>` +
  `<p:spPr/><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`;

const slideXml = (tree: string, show?: string): string =>
  `<?xml version="1.0"?><p:sld ${P_NS}${show === undefined ? "" : ` show="${show}"`}><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
  `${tree}</p:spTree></p:cSld></p:sld>`;

/** A package whose `p:sldIdLst` names `ppt/slides/slide<n>.xml` in `order`;
 *  `slideRels[n]` is the `Relationship` elements of slide n's part; `files`
 *  other parts by archive path (media); `presentationExtra` elements written
 *  after `p:sldIdLst` (`p:sldSz`). */
const packageOf = async (
  order: readonly number[],
  parts: Readonly<Record<number, string>>,
  slideRels: Readonly<Record<number, string>> = {},
  files: Readonly<Record<string, Uint8Array>> = {},
  presentationExtra = "",
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/></Types>',
  );
  for (const [path, bytes] of Object.entries(files)) zip.file(path, bytes);
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      order.map((n, i) => `<p:sldId id="${256 + i}" r:id="rId${n}"/>`).join("") +
      `</p:sldIdLst>${presentationExtra}</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      Object.keys(parts)
        .map((n) => `<Relationship Id="rId${n}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`)
        .join("") +
      "</Relationships>",
  );
  for (const [n, xml] of Object.entries(parts)) zip.file(`ppt/slides/slide${n}.xml`, xml);
  for (const [n, xml] of Object.entries(slideRels)) {
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

// Declared order: slide3, slide1 (hidden), slide2 — not the part-name order.
const synthetic = (): Promise<Uint8Array> =>
  packageOf([3, 1, 2], {
    3: slideXml(
      shape(2, "<a:p><a:r><a:t>Caf&#233; 007</a:t></a:r><a:br/><a:r><a:t>2</a:t></a:r></a:p>") +
        shape(3, "<a:p><a:r><a:t>Hidden shape</a:t></a:r></a:p>", true) +
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
        shape(5, '<a:p><a:fld id="{0}" type="datetime1"><a:t>25/09/2026</a:t></a:fld></a:p>') +
        "</p:grpSp>",
    ),
    1: slideXml(shape(2, "<a:p><a:r><a:t>Hidden slide</a:t></a:r></a:p>"), "0"),
    2: slideXml(shape(2, "<a:p><a:r><a:t>Last</a:t></a:r></a:p>")),
  });

describe("step 0 — what no witness exercises", () => {
  it("a numeric character reference is decoded; hidden slide and hidden shape skipped; groups walked", async () => {
    const { slides } = await readPptx(await synthetic(), { limits: PAGE_SIZED_LIMITS });
    // Failure values differ from the expected ones: an undecoded reference
    // reads "Caf&#233; 007"; a numeric parse drops "2" or turns "007" into 7;
    // a skipped hidden slide or shape would add "Hidden slide" / "Hidden shape".
    expect(slides.map((s) => [s.number, s.partPath])).toEqual([
      [1, "ppt/slides/slide3.xml"],
      [2, "ppt/slides/slide2.xml"],
    ]);
    expect(textsOf(slides[0]!)).toEqual([[["Café 007", "\n", "2"]], [["25/09/2026"]]]);
    expect(textsOf(slides[1]!)).toEqual([[["Last"]]]);
  });
});

// ---------------------------------------------------------------------------
// Step 1 — the declared title (MIKSUDOC « Spécification fonctionnelle » §9.3.1, FORMAT-PPTX [purged 27/09/2026, git history] §5
// point 5). Expected values read in the raw slide XML before the test.

describe("step 1 — declared title", () => {
  it("synthetic: idx title not taken, ctrTitle taken, repeat flagged against the previous VISIBLE slide", async () => {
    const t = (text: string): string => `<a:p><a:r><a:t xml:space="preserve">${text}</a:t></a:r></a:p>`;
    const { slides } = await readPptx(
      await packageOf([1, 2, 3, 4, 5, 6], {
        // A column header (`title` with idx) and the real title (`ctrTitle`).
        1: slideXml(
          shape(2, t("Header"), false, 'type="title" idx="3"') + shape(3, t("Review"), false, 'type="ctrTitle"'),
        ),
        // Same title with spaces around it: repeated.
        2: slideXml(shape(2, t("  Review ") + t(""), false, 'type="title"') + shape(3, t("Body"))),
        // Hidden slide with another title: skipped, so it breaks no repeat.
        3: slideXml(shape(2, t("Other"), false, 'type="title"'), "0"),
        // Repeated against slide 2, the previous visible one.
        4: slideXml(shape(2, t("Review"), false, 'type="title"')),
        // Case differs: not the same text.
        5: slideXml(shape(2, t("review"), false, 'type="title"')),
        // Only a column header: no title, never guessed.
        6: slideXml(shape(2, t("Big text"), false, 'type="title" idx="1"')),
      }),
    );
    expect(slides.map((s) => [s.number, s.title, s.isTitleRepeated])).toEqual([
      [1, "Review", false],
      [2, "Review", true],
      [3, "Review", true],
      [4, "review", false],
      [5, undefined, false],
    ]);
    expect(slides.map(paragraphsOf)).toEqual([["Header"], ["Body"], [], [], ["Big text"]]);
  });
});

// ---------------------------------------------------------------------------
// Step 2 — marks, links, fields, declared bullets (MIKSUDOC « Spécification fonctionnelle » §9.3
// §5). Only what the slide WRITES on the run or the paragraph; inheritance is
// not read (§8). Expected values read in the raw slide XML before the test.

const textOfParagraph = (p: PptxParagraph): string => p.runs.map((r) => r.text).join("");

describe("step 2 — marks, links, fields, declared bullets", () => {
  it("🔴 a run in the Symbol font (a:sym or a:latin typeface « Symbol ») publishes the characters its private points draw; another font keeps them (26/10/2026)", async () => {
    // FAILS WITH: « \uF071 » kept — Confluence draws no θ for it, and search cannot find it.
    const run = (rPr: string, text: string): string =>
      `<a:r><a:rPr lang="fr-FR">${rPr}</a:rPr><a:t>${text}</a:t></a:r>`;
    const body =
      "<a:p>" +
      run('<a:sym typeface="Symbol"/>', "\uF071") +
      run('<a:latin typeface="Symbol"/>', "\uF061") +
      run('<a:sym typeface="Wingdings"/>', "\uF0B7") +
      "</a:p>";
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(shape(2, body)) }), {
      limits: PAGE_SIZED_LIMITS,
    });
    expect(runsOf(slides[0]!)).toEqual(["θ", "α", "\uF0B7"]);
  });

  it("🔴 a:sym decides the symbols' font when it is written: a:latin Symbol with a:sym Wingdings keeps the point (counter-review of 26/10/2026)", async () => {
    // FAILS WITH: « θ » — a Wingdings box (U+F071) translated by the Symbol table.
    const body =
      '<a:p><a:r><a:rPr lang="fr-FR"><a:latin typeface="Symbol"/><a:sym typeface="Wingdings"/></a:rPr><a:t>\uF071</a:t></a:r></a:p>';
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(shape(2, body)) }), {
      limits: PAGE_SIZED_LIMITS,
    });
    expect(runsOf(slides[0]!)).toEqual(["\uF071"]);
  });

  it("synthetic: each mark, baseline ±, noStrike, u=none, link kinds, fields, bullets", async () => {
    const r = (rPr: string, text: string, inner = ""): string =>
      `<a:r><a:rPr lang="fr-FR"${rPr}>${inner}</a:rPr><a:t xml:space="preserve">${text}</a:t></a:r>`;
    const p = (pPr: string, body: string): string => `<a:p>${pPr}${body}</a:p>`;
    const link = (id: string, action?: string): string =>
      `<a:hlinkClick r:id="${id}"${action === undefined ? "" : ` action="${action}"`}/>`;
    const body =
      p(
        "",
        r(' b="1"', "b") +
          r(' i="true"', "i") +
          r(' u="sng"', "u") +
          r(' u="none"', "uNone") +
          r(' strike="sngStrike"', "s") +
          r(' strike="noStrike"', "sNone") +
          r(' baseline="30000"', "sup") +
          r(' baseline="-25000"', "sub") +
          r(' baseline="0"', "base0") +
          r(' b="0" i="false"', "off") +
          r(' b="1" i="1" u="dbl" strike="dblStrike"', "all") +
          "<a:r><a:t>bare</a:t></a:r>",
      ) +
      p(
        "",
        r("", "ext", link("rId2")) +
          r("", "drive", link("rId3")) +
          r("", "jump", link("rId4", "ppaction://hlinksldjump")) +
          r("", "next", link("", "ppaction://hlinkshowjump?jump=nextslide")) +
          r("", "internal", link("rId4")),
      ) +
      p(
        "",
        '<a:fld id="{1}" type="slidenum"><a:rPr lang="fr-FR"/><a:t>7</a:t></a:fld>' +
          '<a:fld id="{2}" type="datetime1"><a:rPr lang="fr-FR" b="1"/><a:t>25/09/2026</a:t></a:fld>',
      ) +
      p('<a:pPr lvl="1"><a:buChar char="•"/></a:pPr>', r("", "bullet")) +
      p('<a:pPr><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>', r("", "three")) +
      p('<a:pPr lvl="0"><a:buAutoNum type="alphaLcParenR"/></a:pPr>', r("", "one")) +
      p('<a:pPr><a:buBlip><a:blip r:embed="rId9"/></a:buBlip></a:pPr>', r("", "picture")) +
      p('<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>', r("", "none")) +
      p('<a:pPr lvl="2"/>', r("", "undeclared"));
    const { slides } = await readPptx(
      await packageOf(
        [1, 2],
        { 1: slideXml(shape(2, body)), 2: slideXml(shape(2, p("", r("", "x")))) },
        {
          1:
            `<Relationship Id="rId2" Type="${REL}/hyperlink" Target="https://example.com/a?x=1&amp;y=2" TargetMode="External"/>` +
            `<Relationship Id="rId3" Type="${REL}/hyperlink" Target="https://drive.google.com/file/d/abc/view" TargetMode="External"/>` +
            `<Relationship Id="rId4" Type="${REL}/slide" Target="slide2.xml"/>`,
        },
      ),
    );
    const [marks, links, fields, ...lists] = textShapes(slides[0]!)[0]!.paragraphs;
    // Failure values: no marks at all, or "uNone"/"sNone"/"off"/"base0" marked.
    expect(marks!.runs).toEqual([
      { text: "b", marks: ["strong"] },
      { text: "i", marks: ["em"] },
      { text: "u", marks: ["underline"] },
      { text: "uNone" },
      { text: "s", marks: ["strike"] },
      { text: "sNone" },
      { text: "sup", marks: ["sup"] },
      { text: "sub", marks: ["sub"] },
      { text: "base0" },
      { text: "off" },
      { text: "all", marks: ["strong", "em", "underline", "strike"] },
      { text: "bare" },
    ]);
    // An action (`ppaction://`) or an internal target is text without a link;
    // a Drive link stays a plain URL (normalize-links rewrites it later).
    expect(links!.runs).toEqual([
      { text: "ext", href: "https://example.com/a?x=1&y=2" },
      { text: "drive", href: "https://drive.google.com/file/d/abc/view" },
      { text: "jump" },
      { text: "next" },
      { text: "internal" },
    ]);
    // Every field keeps its displayed text and marks, the slide number too (19/10/2026).
    expect(fields!.runs).toEqual([{ text: "1" }, { text: "25/09/2026", marks: ["strong"] }]);
    expect(lists.map((l) => [textOfParagraph(l), l.list])).toEqual([
      ["bullet", { ordered: false, level: 1 }],
      ["three", { ordered: true, level: 0, start: 3 }],
      ["one", { ordered: true, level: 0 }],
      ["picture", { ordered: false, level: 0 }],
      ["none", undefined],
      ["undeclared", undefined],
    ]);
  });
});

// ---------------------------------------------------------------------------
// Step 3 — tables (`a:tbl`) and the other graphic frames
// (MIKSUDOC « Spécification fonctionnelle » §9.3.5). Expected values read in the raw slide
// XML before the test.

/** A table as rows of "text" or "text|cs2|rs3|h" (cell paragraphs joined by "\n"). */
const gridOf = (table: PptxTable): string[][] =>
  table.rows.map((row) =>
    row.cells.map((cell) =>
      [
        cell.paragraphs.map(textOfParagraph).join("\n"),
        ...(cell.colspan === undefined ? [] : [`cs${cell.colspan}`]),
        ...(cell.rowspan === undefined ? [] : [`rs${cell.rowspan}`]),
        ...(cell.header === true ? ["h"] : []),
      ].join("|"),
    ),
  );

const tablesOf = (slide: PptxSlide): PptxTable[] => slide.shapes.flatMap((s) => (s.kind === "table" ? [s.table] : []));

describe("step 3 — tables and graphic frames", () => {
  it("03: the chart frame is kept as a chart with its part and the part's data; the slide's text shape is intact", async () => {
    const { slides } = await readReal("pptx/03-Presentation-Resultats-T1.pptx");
    // slide3.xml: one text box, then a `p:graphicFrame` whose `c:chart r:id="rId1"`
    // targets `/ppt/charts/chart1.xml`. Failure value: the frame absent.
    const three = slideNumbered(slides, 3);
    expect(three.shapes.map((s) => s.kind)).toEqual(["text", "graphic"]);
    expect(three.shapes[1]).toEqual({
      kind: "graphic",
      spTreeIndex: 1,
      id: "3",
      // `p:xfrm`, not `a:xfrm`, on a graphic frame (step 5).
      box: { x: 914_400, y: 1_188_720, cx: 10_058_400, cy: 4_846_320 },
      graphic: "chart",
      part: "ppt/charts/chart1.xml",
      // chart1.xml: one bar series over four categories, a title, no
      // rendering in the archive (step A, ported from the former reader).
      chart: {
        title: "Évolution trimestrielle du chiffre d'affaires",
        data: {
          headers: ["", "Chiffre d'affaires"],
          rows: [
            ["T2 2025", "1,840"],
            ["T3 2025", "1,955"],
            ["T4 2025", "2,100"],
            ["T1 2026", "2,276"],
          ],
        },
        // The macro draws it: the numbers are the displayed ones ("1,840" is 1840).
        drawing: { type: "bar", values: [[1840, 1955, 2100, 2276]] },
      },
    });
    expect(paragraphsOf(three)).toEqual(["Chiffre d'affaires par trimestre (k€)"]);
  });

  it("synthetic: gridSpan, rowSpan, absorbed cells, firstRow on and off, hidden frame, other frames", async () => {
    const tc = (attrs: string, ...paragraphs: string[]): string =>
      `<a:tc${attrs}><a:txBody><a:bodyPr/>${paragraphs.map((p) => `<a:p>${p}</a:p>`).join("")}</a:txBody><a:tcPr/></a:tc>`;
    const r = (text: string, rPr = ""): string => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`;
    const tr = (...cells: string[]): string => `<a:tr h="100">${cells.join("")}</a:tr>`;
    const frame = (id: number, data: string, uri: string, hidden = false): string =>
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="f${id}"${hidden ? ' hidden="1"' : ""}/>` +
      '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>' +
      `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;
    const TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table";
    const table = (tblPr: string, columns: number, ...rows: string[]): string =>
      `<a:tbl>${tblPr}<a:tblGrid>${Array.from({ length: columns }, (_, i) => `<a:gridCol w="${(i + 1) * 100}"/>`).join("")}</a:tblGrid>${rows.join("")}</a:tbl>`;
    // Absorbed cells carry text on purpose: kept, it would show.
    const on = table(
      '<a:tblPr firstRow="1" bandRow="1"/>',
      3,
      tr(tc(' gridSpan="2"', r("A")), tc(' hMerge="1"', r("absorbed-h")), tc("", r("B"))),
      tr(
        tc(' rowSpan="2"', r("C")),
        tc("", '<a:pPr lvl="1"><a:buChar char="•"/></a:pPr>' + r("D", '<a:rPr b="1"/>')),
        tc("", r("E")),
      ),
      tr(tc(' vMerge="1"', r("absorbed-v")), tc("", r("F")), tc("", r("G1"), r("G2"))),
    );
    const off = table(
      '<a:tblPr firstRow="0"/>',
      3,
      tr(tc(' gridSpan="2" rowSpan="2"', r("X")), tc(' hMerge="1"'), tc("", r("Y"))),
      tr(tc(' vMerge="true"'), tc(' hMerge="1" vMerge="1"'), tc("", r("Z"))),
    );
    const { slides } = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            shape(2, "<a:p><a:r><a:t>Before</a:t></a:r></a:p>") +
              frame(3, on, TABLE) +
              frame(4, table("", 1, tr(tc("", r("hidden")))), TABLE, true) +
              frame(5, off, TABLE) +
              frame(
                6,
                '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>',
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
              ) +
              frame(
                7,
                '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId3"/>',
                "http://schemas.openxmlformats.org/drawingml/2006/diagram",
              ) +
              frame(
                8,
                '<p:oleObj r:id="rId4" progId="Excel.Sheet.12"/>',
                "http://schemas.openxmlformats.org/presentationml/2006/ole",
              ) +
              frame(9, "", "urn:unknown") +
              frame(10, "<a:tbl></a:tbl>", TABLE),
          ),
        },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart7.xml"/>` },
      ),
    );
    const shapes = slides[0]!.shapes;
    // spTreeIndex counts the hidden frame (rank 2) too.
    expect(shapes.map((s) => [s.spTreeIndex, s.kind])).toEqual([
      [0, "text"],
      [1, "table"],
      [3, "table"],
      [4, "graphic"],
      [5, "graphic"],
      [6, "graphic"],
      [7, "graphic"],
      [8, "table"],
    ]);
    const [first, second, empty] = tablesOf(slides[0]!);
    expect(first!.columnWidths).toEqual([100, 200, 300]);
    expect(gridOf(first!)).toEqual([
      ["A|cs2|h", "B|h"],
      ["C|rs2", "D", "E"],
      ["F", "G1\nG2"],
    ]);
    // A cell's paragraphs are read like any other: marks and bullet kept.
    expect(first!.rows[1]!.cells[1]!.paragraphs).toEqual([
      { runs: [{ text: "D", marks: ["strong"] }], list: { ordered: false, level: 1 }, level: 1 },
    ]);
    // firstRow="0": no header; a 2×2 merge leaves one cell; the second row,
    // absorbed but for its last cell, stays and holds only "Z".
    expect(gridOf(second!)).toEqual([["X|cs2|rs2", "Y"], ["Z"]]);
    // A table with no grid and no row is still a table, empty.
    expect(empty).toEqual({ columnWidths: [], rows: [] });
    expect(shapes.filter((s) => s.kind === "graphic")).toEqual([
      {
        kind: "graphic",
        spTreeIndex: 4,
        id: "6",
        box: { x: 0, y: 0, cx: 1, cy: 1 },
        graphic: "chart",
        part: "ppt/charts/chart7.xml",
        chart: "withheld",
      },
      { kind: "graphic", spTreeIndex: 5, id: "7", box: { x: 0, y: 0, cx: 1, cy: 1 }, graphic: "smartart" },
      { kind: "graphic", spTreeIndex: 6, id: "8", box: { x: 0, y: 0, cx: 1, cy: 1 }, graphic: "ole" },
      { kind: "graphic", spTreeIndex: 7, id: "9", box: { x: 0, y: 0, cx: 1, cy: 1 }, graphic: "other" },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Step 4 — pictures (MIKSUDOC « Spécification fonctionnelle » §9.3.5, §9.3.6). Expected values
// read in the raw slide XML and the PNG headers before the test.

const picturesOf = (slide: PptxSlide): PptxPictureShape[] =>
  slide.shapes.filter((s): s is PptxPictureShape => s.kind === "picture");

/** Width and height from a PNG's IHDR or a baseline/progressive JPEG's start-of-frame. */
const dimensionsOf = (bytes: Uint8Array): [number, number] => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes[0] === 0x89) return [view.getUint32(16), view.getUint32(20)];
  let at = 2;
  for (;;) {
    const marker = bytes[at + 1]!;
    if (marker >= 0xc0 && marker <= 0xc2) return [view.getUint16(at + 7), view.getUint16(at + 5)];
    at += 2 + view.getUint16(at + 2);
  }
};

describe("step 4 — pictures, width, crop", () => {
  /** Audit 07/10, P2 « non-doublons » 1. `image-reduction.ts` leaves a JPEG whose EXIF turns the image untouched, because re-encoding
   *  drops the turn; the crop re-encoded without reading it, so a phone JPEG (orientation 6) cropped in a pptx came out unturned.
   *  Same policy now: such a JPEG is published whole (every word of the picture stays), never cut. */
  it("a cropped JPEG that its EXIF turns is published whole; the same JPEG without a turn is cut", async () => {
    const pixels = new Uint8Array(16 * 16 * 4);
    for (let k = 0; k < 256; k++) pixels.set(k % 16 < 8 ? [255, 0, 0, 255] : [0, 0, 255, 255], k * 4);
    const plain = new Uint8Array(encodeJpeg({ width: 16, height: 16, data: pixels }, 90).data);
    const payload = [
      0x45, 0x78, 0x69, 0x66, 0, 0, 0x4d, 0x4d, 0x00, 0x2a, 0, 0, 0, 8, 0, 1, 0x01, 0x12, 0, 3, 0, 0, 0, 1, 0, 6, 0, 0,
      0, 0, 0, 0, 0, 0,
    ];
    const exif = new Uint8Array([0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload]);
    const turned = new Uint8Array(plain.length + exif.length);
    turned.set(plain.subarray(0, 2), 0);
    turned.set(exif, 2);
    turned.set(plain.subarray(2), 2 + exif.length);
    const pic = (id: number, rid: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:srcRect r="50000"/><a:stretch/></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm></p:spPr></p:pic>`;
    const { slides } = await readPptx(
      await packageOf(
        [1],
        { 1: slideXml(pic(2, "rId2") + pic(3, "rId3")) },
        {
          1: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.jpeg"/><Relationship Id="rId3" Type="${REL}/image" Target="../media/image2.jpeg"/>`,
        },
        { "ppt/media/image1.jpeg": turned, "ppt/media/image2.jpeg": plain },
      ),
    );
    const [first, second] = picturesOf(slides[0]!);
    if (!first || !("image" in first) || !second || !("image" in second))
      throw new Error("both pictures must be published");
    expect(first.image.bytes).toEqual(turned); // failure value: an 8-pixel-wide cut, unturned
    expect(decodeJpeg(second.image.bytes, { useTArray: true }).width).toBe(8); // the control still cuts
  });

  it("a cropped JPEG whose data is damaged becomes the cropped-unsupported marker, never a cut of grey blocks", async () => {
    const pixels = new Uint8Array(32 * 32 * 4);
    for (let k = 0; k < 1024; k++) pixels.set([(k * 37) & 255, (k * 91) & 255, (k * 13) & 255, 255], k * 4);
    const whole = new Uint8Array(encodeJpeg({ width: 32, height: 32, data: pixels }, 90).data);
    const damaged = whole.subarray(0, Math.floor(whole.length / 2));
    const pic = (rid: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="2" name="p2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:srcRect r="50000"/><a:stretch/></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm></p:spPr></p:pic>`;
    const read = async (jpeg: Uint8Array) =>
      picturesOf(
        (
          await readPptx(
            await packageOf(
              [1],
              { 1: slideXml(pic("rId2")) },
              { 1: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.jpeg"/>` },
              { "ppt/media/image1.jpeg": jpeg },
            ),
          )
        ).slides[0]!,
      )[0]!;
    expect(await read(damaged)).toMatchObject({ notPublished: "cropped-unsupported" });
    const control = await read(whole);
    if (!("image" in control)) throw new Error("the intact JPEG must be cut");
    expect(decodeJpeg(control.image.bytes, { useTArray: true }).width).toBe(16);
  });

  it("synthetic: PNG and JPEG cut, GIF cropped → marker, window outside → whole, zeros untouched, external link, group, width, no alt", async () => {
    // A 4×2 RGB PNG, every pixel distinct: row y, column x = (10y + x, 100 + x, 200).
    const png = encodePng(
      4,
      2,
      "rgb-24",
      [0, 1].flatMap((y) => [0, 1, 2, 3].flatMap((x) => [10 * y + x, 100 + x, 200])),
    )!;
    // A 16×16 JPEG, left half red, right half blue.
    const jpegPixels = new Uint8Array(16 * 16 * 4);
    for (let i = 0; i < 256; i++) jpegPixels.set(i % 16 < 8 ? [255, 0, 0, 255] : [0, 0, 255, 255], i * 4);
    const jpeg = new Uint8Array(encodeJpeg({ width: 16, height: 16, data: jpegPixels }, 90).data);
    // A 1×1 GIF: never decoded, only recognised.
    const gif = new Uint8Array(Buffer.from("R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==", "base64"));
    const pic = (
      id: number,
      blip: string,
      srcRect = "",
      opts: { hidden?: boolean; nvPr?: string; cx?: number } = {},
    ): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}" descr="alt text ${id}"${opts.hidden ? ' hidden="1"' : ""}/>` +
      `<p:cNvPicPr/><p:nvPr>${opts.nvPr ?? ""}</p:nvPr></p:nvPicPr><p:blipFill><a:blip ${blip}/>${srcRect}<a:stretch/></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.cx ?? 1270000}" cy="635000"/></a:xfrm></p:spPr></p:pic>`;
    const { slides } = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            pic(2, 'r:embed="rId2"', '<a:srcRect l="50000" t="50000"/>') +
              pic(3, 'r:embed="rId3"', '<a:srcRect r="50000"/>') +
              pic(4, 'r:embed="rId4"', '<a:srcRect b="10000"/>') +
              pic(5, 'r:embed="rId4"', '<a:srcRect b="0" l="0" r="0" t="0"/>') +
              pic(6, 'r:link="rId5"') +
              pic(7, 'r:embed="rId5"') +
              pic(8, 'r:embed="rId99"') +
              pic(9, 'r:embed="rId2"', '<a:srcRect t="-50000" b="50000"/>') +
              pic(10, 'r:embed="rId2"', "", { nvPr: '<a:videoFile r:link="rId5"/>' }) +
              pic(11, 'r:embed="rId2"', "", { hidden: true }) +
              '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="12" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
              pic(13, 'r:embed="rId2"', "", { cx: 63500 }) +
              "</p:grpSp>" +
              // Window entirely above the image (y from -150 % to -50 %).
              pic(14, 'r:embed="rId2"', '<a:srcRect t="-150000" b="150000"/>') +
              // Window touching the image's top edge (y from -100 % to 0).
              pic(15, 'r:embed="rId2"', '<a:srcRect t="-100000" b="100000"/>') +
              // Window partly outside: negative left inset only.
              pic(16, 'r:embed="rId2"', '<a:srcRect l="-25000"/>') +
              // Window entirely outside a GIF: no cut needed, published whole.
              pic(17, 'r:embed="rId4"', '<a:srcRect l="120000" r="-120000"/>'),
          ),
        },
        {
          1:
            `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>` +
            `<Relationship Id="rId3" Type="${REL}/image" Target="../media/image2.jpeg"/>` +
            `<Relationship Id="rId4" Type="${REL}/image" Target="../media/image3.gif"/>` +
            `<Relationship Id="rId5" Type="${REL}/image" Target="https://example.com/a.png" TargetMode="External"/>`,
        },
        { "ppt/media/image1.png": png, "ppt/media/image2.jpeg": jpeg, "ppt/media/image3.gif": gif },
      ),
    );
    const pictures = picturesOf(slides[0]!);
    // The hidden picture (rank 9) is skipped; the group (rank 10) is walked.
    expect(pictures.map((p) => [p.spTreeIndex, "image" in p ? p.image.contentType : p.notPublished])).toEqual([
      [0, "image/png"],
      [1, "image/jpeg"],
      [2, "cropped-unsupported"],
      [3, "image/gif"],
      [4, "external-link"],
      [5, "external-link"],
      [6, "missing-part"],
      [7, "cropped-unsupported"],
      [8, "video-or-audio"],
      [11, "image/png"],
      [12, "image/png"],
      [13, "image/png"],
      [14, "cropped-unsupported"],
      [15, "image/gif"],
    ]);
    const image = (i: number): Uint8Array => {
      const p = pictures[i]!;
      if (!("image" in p)) throw new Error(`picture ${i} not published`);
      return p.image.bytes;
    };
    // PNG l=50 %, t=50 %: the bottom-right 2×1. Failure values: the whole
    // 4×2, or the top-left quarter (0, 100, 200), (1, 101, 200).
    const cutPng = decodePng(image(0), 100)!;
    expect([cutPng.width, cutPng.height, [...cutPng.pixels]]).toEqual([2, 1, [12, 102, 200, 13, 103, 200]]);
    // JPEG r=50 %: the left 8×16, red (lossy: read with a margin).
    // Failure values: 16 wide, or blue pixels on the right.
    const cutJpeg = decodeJpeg(image(1), { useTArray: true });
    expect([cutJpeg.width, cutJpeg.height]).toEqual([8, 16]);
    for (let i = 0; i < 8 * 16; i++) {
      expect(cutJpeg.data[i * 4]).toBeGreaterThan(200);
      expect(cutJpeg.data[i * 4 + 2]).toBeLessThan(60);
    }
    // srcRect all zeros: the GIF untouched.
    expect(image(3)).toEqual(gif);
    // Window entirely outside (or only touching the edge): the part's bytes
    // untouched. Failure value: the marker.
    expect(image(10)).toEqual(png);
    expect(image(11)).toEqual(png);
    expect(image(13)).toEqual(gif);
    // Declared width: cx 1 270 000 EMU = 100 pt; the file is 2 px wide. The reader hands both; the caller that
    // renders the page holds, on these two numbers, that it draws the picture 2 px wide, never 133 (100 pt at 96 dpi).
    const first = pictures[0]!;
    const firstPixels = "image" in first ? first.image.pixelWidth : undefined;
    expect([first.widthPt, firstPixels]).toEqual([100, 2]);
    // Inside a group that declares no frame: its own `a:ext` (5 pt).
    expect(pictures[9]!.widthPt).toBe(5);
    // No alternative text anywhere in the model.
    expect(JSON.stringify(pictures, (_k, v: unknown) => (v instanceof Uint8Array ? v.length : v))).not.toContain(
      "alt text",
    );
  });

  it("a JPEG as large as the largest real picture (2 500 × 1 668) is cut under the decoding limits", async () => {
    // 4.2 MP, the USA deck's largest picture (image6.png; image5.jpg is
    // 2 500 × 1 666). Failure value: limits too tight → the marker.
    const data = new Uint8Array(2500 * 1668 * 4).fill(128);
    const jpeg = new Uint8Array(encodeJpeg({ width: 2500, height: 1668, data }, 90).data);
    const { slides } = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            '<p:pic><p:nvPicPr><p:cNvPr id="2" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
              '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="10000"/><a:stretch/></p:blipFill><p:spPr/></p:pic>',
          ),
        },
        { 1: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/big.jpeg"/>` },
        { "ppt/media/big.jpeg": jpeg },
      ),
    );
    const picture = picturesOf(slides[0]!)[0]!;
    expect("image" in picture ? dimensionsOf(picture.image.bytes) : picture.notPublished).toEqual([2250, 1668]);
  }, 60_000);
});

// ---------------------------------------------------------------------------
// Step 5 — shapes in mc:AlternateContent, connectors, boxes on the slide,
// groups, slide size (MIKSUDOC « Spécification fonctionnelle » §9.3.2, FORMAT-PPTX [purged 27/09/2026, git history] §3
// quater). Expected values read in the raw slide XML and checked against the
// study's counts before the test was written.

const connectorsOf = (slide: PptxSlide): PptxConnectorShape[] =>
  slide.shapes.filter((s): s is PptxConnectorShape => s.kind === "connector");

const shapeText = (shape: PptxShape): string =>
  shape.kind === "text" ? shape.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("|") : "";

/** The elements of a slide as the layout will see them: text shapes with
 *  words, pictures, tables and other frames. */
const elementsOf = (slide: PptxSlide): PptxShape[] =>
  slide.shapes.filter(
    (s) =>
      s.kind === "picture" ||
      s.kind === "table" ||
      s.kind === "graphic" ||
      (s.kind === "text" && shapeText(s).trim() !== ""),
  );

/** Each element's two flags, in `elementsOf` order. */
const flagsOf = (elements: readonly PptxShape[], size: Size): { background: boolean; posedOnPicture: boolean }[] =>
  slideItemFlags(
    elements.map((e) => ({ ...(e.box === undefined ? {} : { box: e.box }), picture: e.kind === "picture" })),
    size,
  );

describe("step 5 — alternate content, connectors, boxes, groups, slide size", () => {
  it("synthetic: Choice read and Fallback never, nested group scale, quarter turn, connector, background and posed on", async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';
    const sp = (id: number, text: string, xfrm: string, extra = ""): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm${extra}>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const pic = (id: number, xfrm: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr><a:xfrm>${xfrm}</a:xfrm></p:spPr></p:pic>`;
    const at = (x: number, y: number, cx: number, cy: number): string =>
      `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`;
    const group = (id: number, xfrm: string, members: string): string =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="g${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm>${xfrm}</a:xfrm></p:grpSpPr>${members}</p:grpSp>`;
    const connector = (id: number, glue: string, ln: string): string =>
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="c${id}"/><p:cNvCxnSpPr>${glue}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
      `<p:spPr><a:xfrm flipH="1">${at(300_000, 125_000, 330_000, 225_000)}</a:xfrm>` +
      `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln>${ln}</a:ln></p:spPr></p:cxnSp>`;
    const png = encodePng(1, 1, "rgb-24", [0, 0, 0])!;
    const { slideSize, slides } = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            // rank 0: a picture covering the whole 1 000 000 × 500 000 slide.
            pic(2, at(0, 0, 1_000_000, 500_000)) +
              // rank 1: words over it, centre (200 000, 125 000): no other picture there.
              sp(3, "Over the background", at(100_000, 100_000, 200_000, 50_000)) +
              // rank 2: a small picture; rank 3: a digit whose centre (650 000, 350 000) is on it.
              pic(4, at(600_000, 300_000, 100_000, 100_000)) +
              sp(5, "7", at(630_000, 330_000, 40_000, 40_000)) +
              // rank 4: a group at half scale (ext 200 000 for chExt 400 000),
              // holding rank 5, a nested group (rank 6: its member rank 7, a
              // third of its size), and rank 8, a picture 20 pt wide.
              group(
                6,
                `${at(100_000, 200_000, 200_000, 100_000)}<a:chOff x="0" y="0"/><a:chExt cx="400000" cy="200000"/>`,
                sp(7, "Outer member", at(0, 0, 100_000, 50_000)) +
                  group(
                    8,
                    `${at(200_000, 0, 200_000, 200_000)}<a:chOff x="1000" y="1000"/><a:chExt cx="600000" cy="600000"/>`,
                    sp(9, "Inner member", at(301_000, 1_000, 150_004, 3_000)),
                  ) +
                  pic(14, at(0, 100_000, 254_000, 50_000)),
              ) +
              // rank 9: turned a quarter clockwise.
              sp(10, "Turned", at(400_000, 100_000, 200_000, 50_001), ' rot="5400000"') +
              // rank 10: the Choice's shape; the Fallback's is never read.
              `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${sp(11, "Choice words", at(0, 400_000, 100_000, 50_000))}</mc:Choice>` +
              `<mc:Fallback>${sp(11, "Fallback words", at(0, 400_000, 100_000, 50_000))}</mc:Fallback></mc:AlternateContent>` +
              // ranks 11 and 12: connectors.
              connector(
                12,
                '<a:stCxn id="3" idx="2"/><a:endCxn id="5" idx="0"/>',
                '<a:headEnd type="none"/><a:tailEnd type="triangle"/>',
              ) +
              connector(13, "", '<a:headEnd type="arrow"/>'),
          ),
        },
        { 1: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>` },
        { "ppt/media/image1.png": png },
        '<p:sldSz cx="1000000" cy="500000"/>',
      ),
    );
    expect(slideSize).toEqual({ cx: 1_000_000, cy: 500_000 });
    const shapes = slides[0]!.shapes;
    const byId = (id: string): PptxShape => shapes.find((s) => s.id === id)!;

    // AlternateContent: failure values are no "Choice words" (step 0 skipped
    // it) or "Fallback words" too.
    expect(shapes.map(shapeText).filter((t) => t.endsWith("words"))).toEqual(["Choice words"]);
    expect([byId("11").spTreeIndex, byId("12").spTreeIndex]).toEqual([10, 11]);

    // Groups: the outer one halves, the inner one divides by three after
    // moving (1 000, 1 000) to (200 000, 0). Inner member, left edge:
    // 100 000 + (200 000 + 300 000 / 3) / 2 = 250 000; right edge
    // 100 000 + (200 000 + 450 004 / 3) / 2 = 275 000.67 → 275 001 (edges
    // rounded, not the width: 25 000.67 would truncate to 25 000). Failure
    // value: the declared (301 000, 1 000, 150 004, 3 000).
    expect(byId("7").box).toEqual({ x: 100_000, y: 200_000, cx: 50_000, cy: 25_000 });
    expect(byId("9").box).toEqual({ x: 250_000, y: 200_000, cx: 25_001, cy: 500 });
    // Every member, however deep, belongs to the top-level group (rank 4).
    expect(shapes.filter((s) => s.group !== undefined).map((s) => [s.id, s.group])).toEqual([
      ["7", 4],
      ["9", 4],
      ["14", 4],
    ]);
    // The picture in the group: 20 pt declared, 10 pt on the slide.
    expect([byId("14").box, (byId("14") as PptxPictureShape).widthPt]).toEqual([
      { x: 100_000, y: 250_000, cx: 127_000, cy: 25_000 },
      10,
    ]);

    // Quarter turn about the centre (500 000, 125 000.5): 50 001 wide, 200 000
    // tall; edges 474 999.5 → 475 000, 25 000.5 → 25 001 (half up). Failure
    // value: the declared (400 000, 100 000, 200 000, 50 001).
    expect([byId("10").box, byId("10").rot]).toEqual([{ x: 475_000, y: 25_001, cx: 50_001, cy: 200_000 }, 5_400_000]);

    // Connectors: glued ends, arrow heads; a flip leaves the box.
    const seen = connectorsOf(slides[0]!).map(({ id, from, to, startArrow, endArrow, box }) => ({
      id,
      from,
      to,
      startArrow,
      endArrow,
      box,
    }));
    expect(seen).toEqual([
      {
        id: "12",
        from: "3",
        to: "5",
        startArrow: false,
        endArrow: true,
        box: { x: 300_000, y: 125_000, cx: 330_000, cy: 225_000 },
      },
      {
        id: "13",
        from: undefined,
        to: undefined,
        startArrow: true,
        endArrow: false,
        box: { x: 300_000, y: 125_000, cx: 330_000, cy: 225_000 },
      },
    ]);

    // Background and posed on: the words over the background are no posed
    // element, the digit on the small picture is; the small picture itself
    // lies on the background, which does not count.
    const elements = elementsOf(slides[0]!);
    const flags = flagsOf(elements, slideSize!);
    expect(elements.map((e, i) => [e.id, flags[i]!.background, flags[i]!.posedOnPicture])).toEqual([
      ["2", true, false],
      ["3", false, false],
      ["4", false, false],
      ["5", false, true],
      ["7", false, false],
      ["9", false, false],
      ["14", false, false],
      ["10", false, false],
      ["11", false, false],
    ]);
  });
});

describe("step 7 — what the slide drawing needs", () => {
  it("synthetic: line ends through flips, rotation and a flipped group; edge and line shown or not; size, alignment, anchor, turned text, corner", async () => {
    const at = (x: number, y: number, cx: number, cy: number): string =>
      `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`;
    const line = (id: number, xfrmAttrs: string, ln = "<a:solidFill/>", style = ""): string =>
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="c${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
      `<p:spPr><a:xfrm${xfrmAttrs}>${at(100_000, 100_000, 200_000, 50_000)}</a:xfrm>` +
      `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln>${ln}</a:ln></p:spPr>${style}</p:cxnSp>`;
    const box = (id: number, spPr: string, bodyPr = "<a:bodyPr/>", paragraphs = "<a:p/>", style = ""): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm>${at(0, 0, 100_000, 100_000)}</a:xfrm>${spPr}</p:spPr>${style}` +
      `<p:txBody>${bodyPr}${paragraphs}</p:txBody></p:sp>`;
    const rect = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    const { slides } = await readPptx(
      await packageOf([1], {
        1: slideXml(
          line(2, "") +
            line(3, ' flipH="1"') +
            line(4, ' flipV="1"') +
            line(5, ' rot="10800000"') +
            // A group flipped horizontally about its centre x = 500 000,
            // child coordinates equal to its own: the line mirrored.
            `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
            `<p:grpSpPr><a:xfrm flipH="1">${at(0, 0, 1_000_000, 1_000_000)}<a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>` +
            `${line(7, "")}</p:grpSp>` +
            line(8, "", "<a:noFill/>") +
            line(9, "", "", '<p:style><a:lnRef idx="1"/><a:fillRef idx="0"/></p:style>') +
            line(10, "", "") +
            box(11, `${rect}<a:noFill/><a:ln><a:noFill/></a:ln>`) +
            box(12, `${rect}<a:solidFill/><a:ln><a:noFill/></a:ln>`) +
            box(13, `${rect}<a:noFill/><a:ln><a:solidFill/></a:ln>`) +
            box(14, rect, "<a:bodyPr/>", "<a:p/>", '<p:style><a:lnRef idx="0"/><a:fillRef idx="1"/></p:style>') +
            box(
              15,
              '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>',
              '<a:bodyPr anchor="b"/>',
              '<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1350"/><a:t>A</a:t></a:r><a:r><a:t>B</a:t></a:r></a:p><a:p><a:pPr algn="just"/></a:p>',
            ) +
            box(16, rect, '<a:bodyPr vert="vert270" anchor="t"/>') +
            box(17, rect, '<a:bodyPr rot="5400000"/>') +
            box(18, rect, '<a:bodyPr vert="horz"/>'),
        ),
      }),
    );
    const shapes = slides[0]!.shapes;
    const byId = (id: string): PptxShape => shapes.find((s) => s.id === id)!;
    const ends = (id: string): unknown => (byId(id) as PptxConnectorShape).ends;
    // The frame is (100 000, 100 000) to (300 000, 150 000). Failure value
    // for 3, 4, 5 and 7: the unflipped { 100 000, 100 000, 300 000, 150 000 }.
    expect(ends("2")).toEqual({ x1: 100_000, y1: 100_000, x2: 300_000, y2: 150_000 });
    expect(ends("3")).toEqual({ x1: 300_000, y1: 100_000, x2: 100_000, y2: 150_000 });
    expect(ends("4")).toEqual({ x1: 100_000, y1: 150_000, x2: 300_000, y2: 100_000 });
    expect(ends("5")).toEqual({ x1: 300_000, y1: 150_000, x2: 100_000, y2: 100_000 });
    expect(ends("7")).toEqual({ x1: 900_000, y1: 100_000, x2: 700_000, y2: 150_000 });
    expect([byId("3").flipH, byId("4").flipV, byId("2").flipH]).toEqual([true, true, undefined]);
    // Line shown: a written fill; `a:noFill` hides it; nothing written takes
    // the style's `a:lnRef`, and without a style there is none.
    expect(["2", "8", "9", "10"].map((id) => (byId(id) as PptxConnectorShape).hasVisibleLine)).toEqual([
      true,
      false,
      true,
      false,
    ]);
    // Fill and line: a text box writing `a:noFill` twice has neither; a fill
    // without a colour, or the style's fill with none, is a fill of unknown
    // colour; a line alone shows the line (step 7 ter).
    const text = (id: string): PptxTextShape => byId(id) as PptxTextShape;
    expect(["11", "12", "13", "14"].map((id) => [text(id).fill, text(id).hasVisibleLine])).toEqual([
      ["none", false],
      ["unknown", false],
      ["none", true],
      ["unknown", false],
    ]);
    // Size in points where written, alignment "ctr"/"r" only, anchor, corner.
    expect(text("15").paragraphs).toEqual([
      { runs: [{ text: "A", size: 13.5 }, { text: "B" }], align: "ctr" },
      { runs: [] },
    ]);
    expect([text("15").anchor, text("15").adj, text("11").anchor, text("11").adj]).toEqual([
      "b",
      50_000,
      undefined,
      undefined,
    ]);
    // Turned text: `vert` other than "horz", or a body `rot`.
    expect(["16", "17", "18"].map((id) => text(id).isTextTurned)).toEqual([true, true, undefined]);
  });
});

describe("step 7 bis — bent connectors, autofit, wrap, default text size", () => {
  it("synthetic: normAutofit in both notations, wrap none, paragraph level, default adj, flips and a flipped group, defaultTextStyle per level", async () => {
    const at = (x: number, y: number, cx: number, cy: number): string =>
      `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`;
    const bent = (id: number, prst: string, xfrmAttrs: string, frame: string, avLst = ""): string =>
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="c${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
      `<p:spPr><a:xfrm${xfrmAttrs}>${frame}</a:xfrm><a:prstGeom prst="${prst}"><a:avLst>${avLst}</a:avLst></a:prstGeom>` +
      `<a:ln><a:solidFill/></a:ln></p:spPr></p:cxnSp>`;
    const body = (id: number, bodyPr: string, paragraphs = "<a:p/>"): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm>${at(0, 0, 100_000, 100_000)}</a:xfrm></p:spPr><p:txBody>${bodyPr}${paragraphs}</p:txBody></p:sp>`;
    const result = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            body(2, '<a:bodyPr wrap="none"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>') +
              body(3, '<a:bodyPr wrap="square"><a:normAutofit fontScale="85%" lnSpcReduction="10%"/></a:bodyPr>') +
              body(
                4,
                "<a:bodyPr><a:normAutofit/></a:bodyPr>",
                '<a:p><a:pPr lvl="2"/><a:r><a:t>deep</a:t></a:r></a:p>',
              ) +
              // Defaults: adj1 = adj2 = 50 000.
              bent(5, "bentConnector4", "", at(0, 0, 100_000, 50_000)) +
              bent(6, "bentConnector2", ' flipV="1"', at(100_000, 100_000, 200_000, 50_000)) +
              `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
              `<p:grpSpPr><a:xfrm flipH="1">${at(0, 0, 1_000_000, 1_000_000)}<a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>` +
              `${bent(8, "bentConnector3", "", at(100_000, 100_000, 200_000, 50_000), '<a:gd name="adj1" fmla="val 25000"/>')}</p:grpSp>` +
              bent(9, "curvedConnector3", "", at(0, 0, 100_000, 50_000)) +
              bent(10, "straightConnector1", "", at(0, 0, 100_000, 50_000)),
          ),
        },
        {},
        {},
        '<p:defaultTextStyle><a:defPPr><a:defRPr sz="3000"/></a:defPPr><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>' +
          '<a:lvl2pPr><a:defRPr/></a:lvl2pPr><a:lvl3pPr><a:defRPr sz="1000"/></a:lvl3pPr></p:defaultTextStyle>',
      ),
    );
    const shapes = result.slides[0]!.shapes;
    const byId = (id: string): PptxShape => shapes.find((s) => s.id === id)!;
    const text = (id: string): PptxTextShape => byId(id) as PptxTextShape;
    const pick = (s: PptxTextShape): unknown => [s.fontScale, s.lineSpaceReduction, s.isUnwrapped];
    // 62 500 thousandths of a percent = 62.5 %; "85%" the Strict notation.
    // Failure value: [undefined, undefined, undefined] (nothing read).
    expect(pick(text("2"))).toEqual([0.625, 0.2, true]);
    expect(pick(text("3"))).toEqual([0.85, 0.1, undefined]);
    expect(pick(text("4"))).toEqual([undefined, undefined, undefined]);
    expect(text("4").paragraphs).toEqual([{ runs: [{ text: "deep" }], level: 2 }]);
    const path = (id: string): unknown => (byId(id) as PptxConnectorShape).path;
    // bentConnector4: x1 = 50 000, y2 = 25 000: (0,0) (x1,0) (x1,y2) (w,y2) (w,h).
    expect(path("5")).toEqual([
      { x: 0, y: 0 },
      { x: 50_000, y: 0 },
      { x: 50_000, y: 25_000 },
      { x: 100_000, y: 25_000 },
      { x: 100_000, y: 50_000 },
    ]);
    // bentConnector2 (l,t) (r,t) (r,b), flipped about y = 125 000.
    expect(path("6")).toEqual([
      { x: 100_000, y: 150_000 },
      { x: 300_000, y: 150_000 },
      { x: 300_000, y: 100_000 },
    ]);
    // bentConnector3 adj1 25 000: (100 000,100 000) (150 000,100 000)
    // (150 000,150 000) (300 000,150 000), mirrored by the group: x' = 1 000 000 − x.
    expect(path("8")).toEqual([
      { x: 900_000, y: 100_000 },
      { x: 850_000, y: 100_000 },
      { x: 850_000, y: 150_000 },
      { x: 700_000, y: 150_000 },
    ]);
    expect([path("9"), path("10")]).toEqual([undefined, undefined]);
    // lvl1pPr and lvl3pPr write a size; lvl2pPr writes none; defPPr is not a level.
    expect(result.defaultTextSizes).toEqual({ 0: 18, 2: 10 });
  });

  it("synthetic (calibration lot, 27/09/2026): written insets, a:lnSpc in percent and in points, a:spAutoFit", async () => {
    const body = (id: number, bodyPr: string, paragraphs = "<a:p/>"): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm></p:spPr><p:txBody>${bodyPr}${paragraphs}</p:txBody></p:sp>`;
    const para = (pPr: string): string => `<a:p>${pPr}<a:r><a:t>w</a:t></a:r></a:p>`;
    const result = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            // All four insets written (IT 3's labels write 68 575 / 34 275),
            // and the box grows with its text.
            body(2, '<a:bodyPr lIns="68575" tIns="34275" rIns="0" bIns="12700"><a:spAutoFit/></a:bodyPr>') +
              // Only one inset written: the others stay PowerPoint's defaults.
              body(3, '<a:bodyPr tIns="0"><a:noAutofit/></a:bodyPr>') +
              // Nothing written.
              body(4, "<a:bodyPr/>") +
              body(
                5,
                "<a:bodyPr/>",
                // 115 % (USA 24 writes 115 000), Strict "90%", 18 pt exactly
                // (1 800 hundredths), and a paragraph writing none.
                para('<a:pPr><a:lnSpc><a:spcPct val="115000"/></a:lnSpc></a:pPr>') +
                  para('<a:pPr><a:lnSpc><a:spcPct val="90%"/></a:lnSpc></a:pPr>') +
                  para('<a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc></a:pPr>') +
                  para("<a:pPr/>"),
              ) +
              // A line break writing its own size (IT 3's « RMS… PREPROD »), and one writing none.
              body(
                6,
                "<a:bodyPr/>",
                '<a:p><a:r><a:t>a</a:t></a:r><a:br><a:rPr sz="900"/></a:br><a:br/><a:r><a:t>b</a:t></a:r></a:p>',
              ),
          ),
        },
        {},
        {},
      ),
    );
    const shapes = result.slides[0]!.shapes as PptxTextShape[];
    const byId = (id: string): PptxTextShape => shapes.find((s) => s.id === id)!;
    // Failure value: undefined everywhere (nothing read, the drawing's flat
    // defaults 91 440 / 45 720 applied to every box).
    expect(byId("2").insets).toEqual({ left: 68_575, top: 34_275, right: 0, bottom: 12_700 });
    expect(byId("3").insets).toEqual({ top: 0 });
    expect(byId("4").insets).toBeUndefined();
    // Failure value: undefined (a:spAutoFit not read).
    expect([byId("2").hasSpAutoFit, byId("3").hasSpAutoFit, byId("4").hasSpAutoFit]).toEqual([
      true,
      undefined,
      undefined,
    ]);
    // Failure value: four undefined (the flat 1.2 applied to every line).
    expect(byId("5").paragraphs.map((p) => p.lineSpacing)).toEqual([
      { percent: 1.15 },
      { percent: 0.9 },
      { points: 18 },
      undefined,
    ]);
    // Failure value: { text: "\n" } twice (the break's size not read; the
    // drawing then gave it the deck's 14 pt default and IT 3's 9 pt
    // paragraph grew to 14 pt).
    expect(byId("6").paragraphs[0]!.runs).toEqual([
      { text: "a" },
      { text: "\n", size: 9 },
      { text: "\n" },
      { text: "b" },
    ]);
  });
});

describe("step 7 ter — exact colours for the slide drawing (25/09/2026)", () => {
  it("synthetic: a modifier, a gradient, a pattern or a picture fill is unknown; lowercase hex is written uppercase", async () => {
    const at =
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    const box = (id: number, fill: string, run = ""): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${at}${fill}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1000">${run}</a:rPr><a:t>t${id}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const { slides } = await readPptx(
      await packageOf([1], {
        1: slideXml(
          box(
            2,
            '<a:solidFill><a:srgbClr val="3c78d8"/></a:solidFill>',
            '<a:solidFill><a:srgbClr val="ff9900"/></a:solidFill>',
          ) +
            box(
              3,
              '<a:solidFill><a:srgbClr val="3C78D8"><a:lumMod val="75000"/></a:srgbClr></a:solidFill>',
              '<a:solidFill><a:srgbClr val="FF9900"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
            ) +
            box(4, "<a:gradFill><a:gsLst/></a:gradFill>", "<a:gradFill><a:gsLst/></a:gradFill>") +
            box(5, '<a:pattFill prst="pct5"/>') +
            box(6, "<a:blipFill/>") +
            box(7, '<a:solidFill><a:srgbClr val="nothex"/></a:solidFill>'),
        ),
      }),
    );
    const shapes = slides[0]!.shapes as PptxTextShape[];
    // Failure value for 3: "#3C78D8" (the modifier ignored).
    expect(shapes.map((s) => s.fill)).toEqual(["#3C78D8", "unknown", "unknown", "unknown", "unknown", "unknown"]);
    expect(shapes.map((s) => s.paragraphs[0]!.runs[0]!.color)).toEqual([
      "#FF9900",
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
  });
});

describe("step 7 ter (5) — theme colours: schemeClr through the colour map, sysClr lastClr, p:style references", () => {
  it("synthetic: sysClr lastClr, a modifier, phClr, the slide's colour map override, fillRef through the theme's fill styles, fontRef", async () => {
    const enc = (x: string): Uint8Array => new TextEncoder().encode(x);
    const rels = (xml: string): Uint8Array =>
      enc(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`,
      );
    const at =
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    const box = (id: number, fill: string, run: string, style = "", lstStyle = ""): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${at}${fill}</p:spPr>${style}<p:txBody><a:bodyPr/><a:lstStyle>${lstStyle}</a:lstStyle>` +
      `<a:p><a:r><a:rPr sz="1000">${run}</a:rPr><a:t>t${id}</a:t></a:r><a:r><a:rPr sz="1000"/><a:t>u${id}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const solid = (color: string): string => `<a:solidFill>${color}</a:solidFill>`;
    const style = (fillIdx: number, fillColor: string, fontColor: string): string =>
      `<p:style><a:lnRef idx="0"/><a:fillRef idx="${fillIdx}">${fillColor}</a:fillRef><a:effectRef idx="0"/><a:fontRef idx="minor">${fontColor}</a:fontRef></p:style>`;
    const theme =
      `<?xml version="1.0"?><a:theme ${P_NS} name="t"><a:themeElements><a:clrScheme name="t">` +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="AA0000"/></a:accent1><a:accent2><a:srgbClr val="00AA00"/></a:accent2>' +
      '<a:accent3><a:srgbClr val="0000AA"/></a:accent3><a:accent4><a:srgbClr val="AAAA00"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="00AAAA"/></a:accent5><a:accent6><a:srgbClr val="AA00AA"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>' +
      '<a:fontScheme name="t"/><a:fmtScheme name="t"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      "<a:gradFill><a:gsLst/></a:gradFill><a:gradFill><a:gsLst/></a:gradFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>";
    const MAP =
      'bg1="lt1" tx1="dk1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
    const master = `<?xml version="1.0"?><p:sldMaster ${P_NS}><p:cSld><p:spTree/></p:cSld><p:clrMap ${MAP}/></p:sldMaster>`;
    const layout = `<?xml version="1.0"?><p:sldLayout ${P_NS}><p:cSld><p:spTree/></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
    const slide =
      `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>` +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      box(2, solid('<a:sysClr val="window" lastClr="FAFAFA"/>'), solid('<a:schemeClr val="tx1"/>')) +
      box(
        3,
        solid('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'),
        solid('<a:schemeClr val="accent1"><a:lumOff val="10000"/></a:schemeClr>'),
      ) +
      box(4, solid('<a:schemeClr val="bg2"/>'), solid('<a:schemeClr val="phClr"/>')) +
      box(
        5,
        "",
        solid('<a:srgbClr val="123456"/>'),
        style(1, '<a:schemeClr val="accent2"/>', '<a:schemeClr val="lt1"/>'),
      ) +
      box(
        6,
        "",
        "",
        style(2, '<a:schemeClr val="accent2"/>', '<a:schemeClr val="lt1"/>'),
        "<a:lvl1pPr><a:defRPr/></a:lvl1pPr>",
      ) +
      box(
        7,
        "",
        "",
        style(
          1,
          '<a:schemeClr val="accent3"><a:shade val="50000"/></a:schemeClr>',
          '<a:schemeClr val="dk1"><a:tint val="50000"/></a:schemeClr>',
        ),
      ) +
      // The slide's own mapping: bg2 → lt2 (the master says dk2).
      `</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping ${MAP.replace('bg2="dk2"', 'bg2="lt2"')}/></p:clrMapOvr></p:sld>`;
    const { slides } = await readPptx(
      await packageOf(
        [1],
        { 1: slide },
        { 1: `<Relationship Id="rIdL" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` },
        {
          "ppt/slideLayouts/slideLayout1.xml": enc(layout),
          "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels(
            `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`,
          ),
          "ppt/slideMasters/slideMaster1.xml": enc(master),
          "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels(
            `<Relationship Id="rId1" Type="${REL}/theme" Target="../theme/theme1.xml"/>`,
          ),
          "ppt/theme/theme1.xml": enc(theme),
        },
      ),
    );
    const shapes = slides[0]!.shapes as PptxTextShape[];
    const colours = (s: PptxTextShape): unknown => s.paragraphs[0]!.runs.map((r) => r.color);
    // 2: sysClr lastClr; tx1 → dk1 → its sysClr lastClr. Failure: "unknown".
    // 3: a modifier on the fill and on the run: unknown both.
    // 4: bg2 through the slide's mapping → lt2 EEEEEE (the master's dk2
    //    would give 111111); phClr outside a style: unknown.
    // 5: fillRef 1 → the theme's first fill style, solid phClr → accent2;
    //    fontRef lt1 for the run writing no colour, the other keeps 123456.
    // 6: fillRef 2 → a gradient: unknown; the shape's lstStyle lvl1pPr
    //    writes an `a:defRPr` without a fill: the chain goes on to the
    //    fontRef, lt1 (step 7 quater; before it, nothing was taken).
    // 7: a modifier on the fillRef and fontRef colours: unknown both.
    expect(shapes.map((s) => s.fill)).toEqual(["#FAFAFA", "unknown", "#EEEEEE", "#00AA00", "unknown", "unknown"]);
    expect(shapes.map(colours)).toEqual([
      ["#000000", undefined],
      [undefined, undefined],
      [undefined, undefined],
      ["#123456", "#FFFFFF"],
      ["#FFFFFF", "#FFFFFF"],
      [undefined, undefined],
    ]);
  });
});

describe("step 7 quater — the colour a run inherits, for the slide drawing (25/09/2026, STATUS 12 bis)", () => {
  it("synthetic: layout matched by idx, master txStyles at level 2, title and other styles, defaultTextStyle for a free shape; a modifier stops the chain", async () => {
    const enc = (x: string): Uint8Array => new TextEncoder().encode(x);
    const rels = (xml: string): Uint8Array =>
      enc(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`,
      );
    const color = (hex: string, modifier = ""): string =>
      `<a:solidFill><a:srgbClr val="${hex}">${modifier}</a:srgbClr></a:solidFill>`;
    const level = (n: number, fill: string): string => `<a:lvl${n}pPr><a:defRPr>${fill}</a:defRPr></a:lvl${n}pPr>`;
    const sp = (id: number, ph: string, lstStyle: string, paragraphs: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle>${lstStyle}</a:lstStyle>${paragraphs}</p:txBody></p:sp>`;
    const para = (text: string, lvl = 0): string =>
      `<a:p><a:pPr lvl="${lvl}"/><a:r><a:rPr/><a:t>${text}</a:t></a:r></a:p>`;
    const tree = (shapes: string): string =>
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      `${shapes}</p:spTree></p:cSld>`;
    const master =
      `<?xml version="1.0"?><p:sldMaster ${P_NS}>` +
      tree(sp(2, '<p:ph type="body" idx="1"/>', level(1, color("AAAAAA")), "")) +
      `<p:txStyles><p:titleStyle>${level(1, color("111111"))}</p:titleStyle>` +
      `<p:bodyStyle>${level(1, color("0000AA")) + level(2, color("00AA00"))}</p:bodyStyle>` +
      `<p:otherStyle>${level(1, color("222222"))}</p:otherStyle></p:txStyles></p:sldMaster>`;
    const layout =
      `<?xml version="1.0"?><p:sldLayout ${P_NS}>` +
      tree(
        sp(2, '<p:ph type="body" idx="1"/>', level(1, color("CC0000")), "") +
          sp(3, '<p:ph idx="2"/>', level(1, color("CC0000", '<a:lumMod val="75000"/>')), ""),
      ) +
      "</p:sldLayout>";
    const slide =
      `<?xml version="1.0"?><p:sld ${P_NS}>` +
      tree(
        // idx 1, no type: its layout placeholder is found by idx (by type,
        // "obj" is not "body", it would not be).
        sp(4, '<p:ph idx="1"/>', "", para("a0") + para("a1", 1)) +
          sp(5, '<p:ph idx="2"/>', "", para("b0")) +
          // A title WITH an idx stays a body shape (step 1).
          sp(6, '<p:ph type="title" idx="5"/>', "", para("c0")) +
          sp(7, '<p:ph type="dt" idx="6"/>', "", para("d0")) +
          sp(8, "", "", para("e0")) +
          sp(9, "", level(1, color("FF9900", '<a:alpha val="50000"/>')), para("f0")),
      ) +
      "</p:sld>";
    const { slides } = await readPptx(
      await packageOf(
        [1],
        { 1: slide },
        { 1: `<Relationship Id="rIdL" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` },
        {
          "ppt/slideLayouts/slideLayout1.xml": enc(layout),
          "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels(
            `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`,
          ),
          "ppt/slideMasters/slideMaster1.xml": enc(master),
        },
        `<p:defaultTextStyle>${level(1, color("333333"))}</p:defaultTextStyle>`,
      ),
    );
    const colours = (slides[0]!.shapes as PptxTextShape[]).map((s) => s.paragraphs.map((p) => p.runs[0]!.color));
    expect(colours).toEqual([
      // a0: layout idx 1, lvl1pPr CC0000. a1 (lvl 1 → lvl2pPr): nothing on
      // the layout nor on the master's body placeholder (AAAAAA is level 1)
      // → master bodyStyle lvl2pPr 00AA00. Failure values: 0000AA (level
      // 1 read), undefined (txStyles not read).
      ["#CC0000", "#00AA00"],
      // b0: layout idx 2 writes CC0000 with lumMod: unknown, and the chain
      // STOPS. Failure value: "#AAAAAA" or "#0000AA" (read past it).
      [undefined],
      // c0: no title placeholder on the layout or the master → titleStyle.
      ["#111111"],
      // d0: dt → otherStyle.
      ["#222222"],
      // e0: free shape, no p:style → defaultTextStyle 333333, not the
      // master's otherStyle (222222).
      ["#333333"],
      // f0: its own lstStyle writes an alpha modifier: unknown, stop.
      // Failure value: "#333333".
      [undefined],
    ]);
  });
});

/** 29/09/2026: a placeholder's text takes the size PowerPoint
 *  gives it — the same chain as the colour (`text-color.ts`), for the size
 *  only, never the position. */
describe("the size a placeholder's run inherits (MIKSUDOC « Spécification fonctionnelle » §9.3.7, lifted 29/09/2026)", () => {
  it("synthetic: own lstStyle, layout by idx then by type, the master's placeholder, the master's txStyles by level; a free shape gets none", async () => {
    const enc = (x: string): Uint8Array => new TextEncoder().encode(x);
    const rels = (xml: string): Uint8Array =>
      enc(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`,
      );
    const level = (n: number, sz: number): string => `<a:lvl${n}pPr><a:defRPr sz="${sz}"/></a:lvl${n}pPr>`;
    const sp = (id: number, ph: string, lstStyle: string, paragraphs: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle>${lstStyle}</a:lstStyle>${paragraphs}</p:txBody></p:sp>`;
    const para = (text: string, lvl = 0): string =>
      `<a:p><a:pPr lvl="${lvl}"/><a:r><a:rPr/><a:t>${text}</a:t></a:r></a:p>`;
    const tree = (shapes: string): string =>
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      `${shapes}</p:spTree></p:cSld>`;
    const master =
      `<?xml version="1.0"?><p:sldMaster ${P_NS}>` +
      tree(sp(2, '<p:ph type="body" idx="1"/>', level(1, 2000), "")) +
      `<p:txStyles><p:titleStyle>${level(1, 4400)}</p:titleStyle>` +
      `<p:bodyStyle>${level(1, 2800) + level(2, 2400)}</p:bodyStyle>` +
      `<p:otherStyle>${level(1, 1200)}</p:otherStyle></p:txStyles></p:sldMaster>`;
    const layout =
      `<?xml version="1.0"?><p:sldLayout ${P_NS}>` +
      tree(
        sp(2, '<p:ph type="body" idx="1"/>', level(1, 1600), "") +
          sp(3, '<p:ph idx="2"/>', "", "") +
          sp(4, '<p:ph type="subTitle" idx="3"/>', level(1, 1200), ""),
      ) +
      "</p:sldLayout>";
    const slide =
      `<?xml version="1.0"?><p:sld ${P_NS}>` +
      tree(
        sp(4, '<p:ph idx="1"/>', "", para("a0") + para("a1", 1)) +
          sp(5, '<p:ph idx="2"/>', "", para("b0")) +
          sp(6, '<p:ph type="title" idx="5"/>', "", para("c0")) +
          sp(7, '<p:ph idx="1"/>', level(1, 900), para("d0")) +
          sp(8, "", "", para("e0")) +
          sp(9, '<p:ph type="subTitle" idx="7"/>', "", para("f0")),
      ) +
      "</p:sld>";
    const { slides } = await readPptx(
      await packageOf(
        [1],
        { 1: slide },
        { 1: `<Relationship Id="rIdL" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` },
        {
          "ppt/slideLayouts/slideLayout1.xml": enc(layout),
          "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels(
            `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`,
          ),
          "ppt/slideMasters/slideMaster1.xml": enc(master),
        },
      ),
    );
    // Failure value (until 29/09/2026): undefined for every shape.
    expect((slides[0]!.shapes as PptxTextShape[]).map((s) => s.inheritedSizes)).toEqual([
      // a0: the layout's idx 1, 16 pt. a1 (lvl2pPr): nothing on the layout
      // nor on the master's body placeholder (its 20 pt is level 1) → the
      // master's bodyStyle, 24 pt.
      { 0: 16, 1: 24 },
      // b0: the layout's idx 2 writes none; its type (none: "obj") is a body
      // on the master → the master's body placeholder, 20 pt.
      { 0: 20 },
      // c0: no title on the layout or the master → titleStyle, 44 pt.
      { 0: 44 },
      // d0: its own lstStyle first, 9 pt.
      { 0: 9 },
      // e0: a free shape: `p:defaultTextStyle`, read by the drawing, as before.
      undefined,
      // f0: idx 7 is on no layout placeholder → by type, the subTitle, 12 pt.
      { 0: 12 },
    ]);
  });
});

describe("step 7 ter (7)(8) — pictures and text in turned, mirrored or stretched groups (25/09/2026)", () => {
  it("synthetic: the rule is checked at EVERY uneven level, against what lies below it; text turned by a group is flagged", async () => {
    const at = (x: number, y: number, cx: number, cy: number): string =>
      `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`;
    const ch = (cx: number, cy: number): string => `<a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/>`;
    const group = (id: number, attrs: string, xfrm: string, members: string): string =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="g${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm${attrs}>${xfrm}</a:xfrm></p:grpSpPr>${members}</p:grpSp>`;
    const pic = (id: number, attrs: string, xfrm: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr><a:xfrm${attrs}>${xfrm}</a:xfrm></p:spPr></p:pic>`;
    const sp = (id: number, xfrm: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm>${xfrm}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:t>t${id}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const png = encodePng(1, 1, "rgb-24", [0, 0, 0])!;
    const { slides } = await readPptx(
      await packageOf(
        [1],
        {
          1: slideXml(
            // Group 3 turns a quarter (uniform); inside it, group 4 is
            // stretched unevenly (sx 2, sy 1) around picture 5, only mirrored:
            // below the uneven level, a diagonal matrix → drawn. Text 6
            // turned by group 3 → flagged.
            group(
              3,
              ' rot="5400000"',
              at(0, 0, 400_000, 400_000) + ch(400_000, 400_000),
              group(
                4,
                "",
                at(0, 0, 400_000, 100_000) + ch(200_000, 100_000),
                pic(5, ' flipH="1"', at(0, 0, 200_000, 100_000)),
              ) + sp(6, at(0, 200_000, 100_000, 50_000)),
            ) +
              // Group 7 is stretched unevenly (sx 2); inside it, group 8
              // turns 30° (uniform) around picture 9, not turned itself: the
              // turn lies below the uneven level → marker. Text 10, only
              // stretched → not flagged.
              group(
                7,
                "",
                at(500_000, 0, 400_000, 100_000) + ch(200_000, 100_000),
                group(
                  8,
                  ' rot="1800000"',
                  at(0, 0, 100_000, 100_000) + ch(100_000, 100_000),
                  pic(9, "", at(0, 0, 100_000, 50_000)),
                ) + sp(10, at(100_000, 50_000, 100_000, 50_000)),
              ),
          ),
        },
        { 1: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>` },
        { "ppt/media/image1.png": png },
        '<p:sldSz cx="1000000" cy="500000"/>',
      ),
    );
    const byId = (id: string): PptxShape => slides[0]!.shapes.find((s) => s.id === id)!;
    const p5 = byId("5") as PptxPictureShape;
    // Picture 5, by hand. Own flip about its centre (100 000, 50 000):
    // (x, y) ↦ (200 000 − x, y). Group 4: (x, y) ↦ (2x, y), so
    // (400 000 − 2x, y). Group 3, a quarter turn about (200 000, 200 000):
    // (X, Y) ↦ (400 000 − Y, X), so (x, y) ↦ (400 000 − y, 400 000 − 2x):
    // a = 0, b = −2, c = −1, d = 0, e = f = 400 000. Its frame as written.
    // Failure value: isTurnedByGroup, no matrix (until c98b204).
    expect([p5.matrix, p5.frame, p5.groupStretched]).toEqual([
      { a: 0, b: -2, c: -1, d: 0, e: 400_000, f: 400_000 },
      { x: 0, y: 0, cx: 200_000, cy: 100_000 },
      undefined,
    ]);
    const p9 = byId("9") as PptxPictureShape;
    // Failure value: a matrix (group 8's turn checked against group 8 only).
    expect([p9.groupStretched, p9.matrix, p9.frame]).toEqual([true, undefined, undefined]);
    expect([(byId("6") as PptxTextShape).isTurnedByGroup, (byId("10") as PptxTextShape).isTurnedByGroup]).toEqual([
      true,
      undefined,
    ]);
  });
});

describe("uneven group scale: a relative tolerance, not strict inequality", () => {
  const grpSpPr = (cx: number, cy: number, chCx: number, chCy: number): XNode =>
    parseXml(
      '<p:grpSpPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${chCx}" cy="${chCy}"/></a:xfrm></p:grpSpPr>`,
    )!;
  // A picture turned a quarter inside the group.
  const turned = { x: 0, y: 0, cx: 200_000, cy: 100_000, rot: 5_400_000, flipH: false, flipV: false };

  it("a group resized with locked aspect, 1 EMU off (1 000 001 × 500 000 over 1 000 000 × 500 000), still draws a turned picture", () => {
    const level = readGroupLevel(grpSpPr(1_000_001, 500_000, 1_000_000, 500_000))!;
    // Failure value: isUnevenlyScaled true, then "stretched-turn".
    expect(level.isUnevenlyScaled).toBe(false);
    expect(placePicture(turned, [level])).not.toBe("stretched-turn");
  });

  it("a genuine 1.5× stretch is still uneven: the turned picture is left out", () => {
    const level = readGroupLevel(grpSpPr(1_500_000, 500_000, 1_000_000, 500_000))!;
    expect(level.isUnevenlyScaled).toBe(true);
    expect(placePicture(turned, [level])).toBe("stretched-turn");
  });
});

describe("p:style/a:fillRef idx 0 and 1000: no fill (ECMA-376 Part 1 §20.1.4.2.10)", () => {
  const fillRef = (idx: number): XNode =>
    parseXml(
      `<a:fillRef xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" idx="${idx}"><a:srgbClr val="112233"/></a:fillRef>`,
    )!;
  const solid = parseXml(
    '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="phClr"/></a:solidFill>',
  )!;
  const context = { clrMap: {}, clrScheme: {}, fillStyleLst: [solid, solid, solid] };

  it("idx 0 and 1000 are none; idx 1 is the theme's first fill style; 1001 stays unknown", () => {
    // Failure value: "unknown" for both (drawn with a spurious outline).
    expect([0, 1000, 1, 1001].map((idx) => resolveFillRef(fillRef(idx), context))).toEqual([
      "none",
      "none",
      "#112233",
      "unknown",
    ]);
  });
});

// fast-check seed -492168866 found it: an unclosed tag between the `?` and
// the `>` of presentation.xml's XML declaration made the parser's own words
// (« Pi Tag is not closed. ») escape the reader. The parser tolerates an
// unclosed ELEMENT, so that spot is where it throws. Every XML part the
// reader parses goes through `package.ts#readXml`: one part per role holds it.
describe("a part whose XML does not parse fails with our own message, never the parser's", () => {
  const corruptDeclaration = async (path: string): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const xml = await zip.file(path)!.async("string");
    const offset = xml.indexOf("?>") + 1;
    zip.file(path, xml.slice(0, offset) + "<p:corruptedUnclosed" + xml.slice(offset));
    return zip.generateAsync({ type: "uint8array" });
  };

  it.each(["ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "ppt/slides/slide1.xml"])("%s", async (path) => {
    await expect(readPptx(await corruptDeclaration(path), { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid PowerPoint presentation",
    });
  });
});

// ---------------------------------------------------------------------------
// 🔴 Arbitrage of 28/09/2026 (option b): a damaged AUXILIARY part —
// a chart, a picture, the theme — gets the SAME treatment as that part
// MISSING (visible marker, default colours), and no longer refuses the whole
// deck. What keeps refusing: the skeleton (the describe above) and the
// inflation ceilings (docx-reader.test.ts, « the package layer names WHICH
// refusal »).

describe("arbitrage 28/09 — a damaged auxiliary part degrades, never the deck", () => {
  const corruptXml = async (path: string): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const xml = await zip.file(path)!.async("string");
    const offset = xml.indexOf("?>") + 1;
    zip.file(path, xml.slice(0, offset) + "<p:corruptedUnclosed" + xml.slice(offset));
    return zip.generateAsync({ type: "uint8array" });
  };

  it("a chart part whose XML does not parse: the deck reads, the chart is withheld", async () => {
    const { slides } = await readPptx(await corruptXml("ppt/charts/chart1.xml"), { limits: PAGE_SIZED_LIMITS });
    expect(slides).toHaveLength(6);
    const frame = slideNumbered(slides, 3).shapes.find((s) => s.kind === "graphic");
    // Failure value: no deck at all (EXTRACT_FAILED today), or a chart with
    // data read from a part that does not parse.
    expect(frame).toMatchObject({ kind: "graphic", graphic: "chart" });
    expect(frame !== undefined && "chart" in frame ? frame.chart : undefined).toBe("withheld");
  });

  it("a theme part whose XML does not parse: the deck reads, its text intact", async () => {
    const { slides } = await readPptx(await corruptXml("ppt/theme/theme1.xml"), { limits: PAGE_SIZED_LIMITS });
    expect(paragraphsOf(slideNumbered(slides, 1))).toEqual([
      "Résultats T1 2026",
      "Atria Solutions — Comité de direction",
    ]);
  });

  it("a picture part whose stream is corrupt: the deck reads, that picture alone is damaged-part", async () => {
    // The docx suite's corruptStream technique: re-deflate the fixture, then
    // break the entry's DEFLATE stream behind its intact local header.
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
    const name = new TextEncoder().encode("ppt/media/image-4-1.png");
    const view = new DataView(bytes.buffer, bytes.byteOffset);
    let nameAt = -1;
    for (let i = 30; i < bytes.length - name.length; i++) {
      if (!name.every((b, j) => bytes[i + j] === b)) continue;
      if (view.getUint32(i - 30, true) === 0x04034b50 && view.getUint16(i - 30 + 26, true) === name.length) {
        nameAt = i;
        break;
      }
    }
    expect(nameAt).toBeGreaterThan(0);
    const dataStart = nameAt + name.length + view.getUint16(nameAt - 30 + 28, true);
    for (let i = 0; i < 8; i++) bytes[dataStart + i] = (bytes[dataStart + i] ?? 0) ^ 0xff;
    // Failure value: no deck at all (EXTRACT_FAILED today); "missing-part"
    // would blame the FILE for a part it does carry.
    const { slides } = await readPptx(bytes, { limits: PAGE_SIZED_LIMITS });
    const reasons = slides.flatMap(picturesOf).map((p) => ("notPublished" in p ? p.notPublished : "published"));
    expect(reasons.filter((r) => r === "damaged-part")).toHaveLength(1);
    expect(reasons.filter((r) => r === "published")).toHaveLength(1);
  });
});

// ECMA-376 Part 3 (Markup Compatibility and Extensibility), AlternateContent:
// a consumer takes the FIRST Choice whose `Requires` namespaces it ALL
// understands, else the Fallback — exactly ONE branch (02/10/2026:
// we publish what Office DISPLAYS). This supersedes the net of 28/09, which
// read a Fallback when the Choice yielded nothing: PowerPoint shows the
// Choice it understands, never its Fallback.
describe("M15/M16 (02/10) — an AlternateContent publishes the ONE branch PowerPoint renders", () => {
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';
  const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
  const words = (id: number, text: string): string => shape(id, `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`);
  // An element of a namespace this reader takes nothing from.
  const opaque = '<u:frame xmlns:u="urn:example:unknown"><u:data/></u:frame>';
  const textsOf = async (tree: string): Promise<string[]> => {
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(tree) }), { limits: PAGE_SIZED_LIMITS });
    return textShapes(slides[0]!).flatMap((s) => s.paragraphs.flatMap((p) => p.runs.map((r) => r.text)));
  };

  it("M15: a Choice requiring a namespace PowerPoint does not understand is skipped for the Fallback", async () => {
    const tree =
      `<mc:AlternateContent ${MC} ${X14}><mc:Choice Requires="x14">${words(2, "excel only words")}</mc:Choice>` +
      `<mc:Fallback>${words(3, "fallback words")}</mc:Fallback></mc:AlternateContent>`;
    // Failure value: "excel only words" (the first Choice read blind).
    expect(await textsOf(tree)).toEqual(["fallback words"]);
  });

  it("M15: the first Choice whose Requires is understood wins, even when it is the second", async () => {
    const tree =
      `<mc:AlternateContent ${MC} ${X14} ${P14}><mc:Choice Requires="x14">${words(2, "first words")}</mc:Choice>` +
      `<mc:Choice Requires="p14">${words(3, "second words")}</mc:Choice>` +
      `<mc:Fallback>${words(4, "fallback words")}</mc:Fallback></mc:AlternateContent>`;
    // Failure value: "first words" (the first Choice, Requires unread).
    expect(await textsOf(tree)).toEqual(["second words"]);
  });

  it("M15: a Choice with no Requires (non-conforming) is never taken", async () => {
    const tree =
      `<mc:AlternateContent ${MC}><mc:Choice>${words(2, "no requires words")}</mc:Choice>` +
      `<mc:Fallback>${words(3, "fallback words")}</mc:Fallback></mc:AlternateContent>`;
    expect(await textsOf(tree)).toEqual(["fallback words"]);
  });

  it("M16: an understood Choice we cannot read, whose Fallback has no image either, publishes nothing", async () => {
    const tree =
      `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${opaque}</mc:Choice>` +
      `<mc:Fallback>${words(2, "fallback words")}</mc:Fallback></mc:AlternateContent>`;
    // Failure value: "fallback words" (the 28/09 net).
    expect(await textsOf(tree)).toEqual([]);
  });

  const picOf = (id: number, embed: string): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${embed}"/><a:stretch/></p:blipFill><p:spPr/></p:pic>`;
  const picturePackage = async (tree: string): Promise<Awaited<ReturnType<typeof readPptx>>> => {
    const real = encodePng(
      4,
      2,
      "rgb-24",
      [0, 1].flatMap(() => [0, 1, 2, 3].flatMap((x) => [x, 100, 200])),
    )!;
    return readPptx(
      await packageOf(
        [1],
        { 1: slideXml(tree) },
        { 1: `<Relationship Id="rIdReal" Type="${REL}/image" Target="../media/real.png"/>` },
        { "ppt/media/real.png": real },
      ),
    );
  };

  it("M16: an understood Choice we cannot read leaves its IMAGE-ONLY Fallback published (the same object Office draws)", async () => {
    const tree = `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${opaque}</mc:Choice><mc:Fallback>${picOf(2, "rIdReal")}</mc:Fallback></mc:AlternateContent>`;
    const { slides } = await picturePackage(tree);
    // Failure value: 0 pictures (the Fallback never read).
    expect(picturesOf(slides[0]!)).toHaveLength(1);
  });

  it("M16: an unread Choice whose image-only Fallback is a 1×1 pixel publishes nothing (junk, not content)", async () => {
    const pixel = encodePng(1, 1, "rgb-24", [255, 255, 255])!;
    const tree = `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${opaque}</mc:Choice><mc:Fallback>${picOf(2, "rIdPixel")}</mc:Fallback></mc:AlternateContent>`;
    const { slides } = await readPptx(
      await packageOf(
        [1],
        { 1: slideXml(tree) },
        { 1: `<Relationship Id="rIdPixel" Type="${REL}/image" Target="../media/pixel.png"/>` },
        { "ppt/media/pixel.png": pixel },
      ),
    );
    // Failure value: 1 picture (the pixel published).
    expect(picturesOf(slides[0]!)).toHaveLength(0);
  });

  it("M16: an unread Choice whose Fallback carries TEXT publishes nothing, picture included", async () => {
    const tree = `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${opaque}</mc:Choice><mc:Fallback>${picOf(2, "rIdReal")}${words(3, "old words")}</mc:Fallback></mc:AlternateContent>`;
    const { slides } = await picturePackage(tree);
    expect([picturesOf(slides[0]!).length, textShapes(slides[0]!).length]).toEqual([0, 0]);
  });

  it("M16: a readable Choice keeps its picture and its image-only Fallback is not read with it", async () => {
    const tree = `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${picOf(2, "rIdReal")}</mc:Choice><mc:Fallback>${picOf(3, "rIdReal")}</mc:Fallback></mc:AlternateContent>`;
    const { slides } = await picturePackage(tree);
    expect(picturesOf(slides[0]!)).toHaveLength(1);
  });

  it("an understood Choice that is read keeps its words, and its Fallback is never published with them", async () => {
    const tree =
      `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${words(2, "modern words")}</mc:Choice>` +
      `<mc:Fallback>${words(3, "old words")}</mc:Fallback></mc:AlternateContent>`;
    expect(await textsOf(tree)).toEqual(["modern words"]);
  });
});

// 8b of the arbitrages (28/09/2026 — « il pourrait se produire
// plusieurs fois non ? », and yes: a FAMILY, declared): `a:endParaRPr sz`
// is the one written size of a box whose runs carry none. Unread, the
// drawing omitted that box's text whole (« size-not-declared », the
// Chatroom false positive).
describe("arbitrage 28/09 (8b) — a:endParaRPr sz is read as the paragraph's declared fallback size", () => {
  it("a run without a size in a paragraph ending endParaRPr sz=1000: endSize 10 on the paragraph", async () => {
    const tree = shape(2, '<a:p><a:r><a:t>sized by the end mark</a:t></a:r><a:endParaRPr sz="1000"/></a:p>');
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(tree) }), { limits: PAGE_SIZED_LIMITS });
    const paragraph = textShapes(slides[0]!)[0]!.paragraphs[0]!;
    // Failure value: endSize absent — the mark is not read.
    expect(paragraph.endSize).toBe(10);
    expect(paragraph.runs[0]).toEqual({ text: "sized by the end mark" });
  });
});

// 7b of the arbitrages (28/09/2026): a `p:sp` with NO `p:txBody`
// used to be dropped whole — a decorative block arrow was invisible to the
// layout, and a slide of boxes joined by arrow SHAPES missed the diagram
// form. It now becomes an empty text shape carrying its geometry, which
// the layout already counts (`slide-layout.ts#ARROW_GEOMETRIES`).
describe("arbitrage 28/09 (7b) — a bodiless shape keeps its geometry; a bodiless non-arrow stays out", () => {
  it("a p:sp without txBody and prstGeom rightArrow is read as an empty text shape with that geometry", async () => {
    const bodiless = (id: number, prst: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="127000"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
    const tree = shape(2, "<a:p><a:r><a:t>box</a:t></a:r></a:p>") + bodiless(3, "rightArrow") + bodiless(4, "rect");
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(tree) }), { limits: PAGE_SIZED_LIMITS });
    // Failure values: the arrow absent (dropped whole, today), or the
    // bodiless rect read too (a decorative rectangle is no arrow and no
    // content: kept out, bias written).
    // The text box's helper writes no prstGeom: its geometry is undefined.
    expect(slides[0]!.shapes.map((s) => [s.kind, "geometry" in s ? s.geometry : undefined])).toEqual([
      ["text", undefined],
      ["text", "rightArrow"],
    ]);
  });
});

describe("revue 28/09 (N1) — an AlternateContent's variants count ONCE toward the element cap", () => {
  it("260 annotations (opaque Choice + junk Fallback each) count 260 elements, not 520", async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    // The reader takes nothing from the Choice; the Fallback yields one
    // contentless text shape the net then discards — yet BOTH used to keep
    // their ranks, so a legitimate ink-annotated slide passed 500 and was
    // wrongly withheld whole, with a lying total on its tab.
    const alt =
      `<mc:AlternateContent ${MC}><mc:Choice Requires="p14"><u:frame xmlns:u="urn:example:unknown"/></mc:Choice>` +
      `<mc:Fallback>${shape(9, "<a:p/>")}</mc:Fallback></mc:AlternateContent>`;
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(alt.repeat(260)) }), {
      limits: PAGE_SIZED_LIMITS,
    });
    expect(slides[0]!.shapes).toHaveLength(0);
    // Failure value: 520 — Choice and Fallback both ranked.
    expect(slides[0]!.elementCount).toBe(260);
  });
});

// 7a of the arbitrages (28/09/2026): at most MAX_SLIDE_ELEMENTS
// elements of one slide are READ — derived from the page's own ceilings
// as they stood then (800 KB, 2.5 MB since 07/11/2026, this cap not raised; over 21 tabs ≈ 38 KB a slide, ≈ 64 bytes the smallest element:
// saturation ≈ 600), never a refusal.
describe("arbitrage 28/09 (7a) — a slide's elements are read up to the cap, and the true total is told", () => {
  it("502 text shapes: the slide comes back withheld (no shapes) with the true count", async () => {
    const tree = Array.from({ length: 502 }, (_, i) => shape(i + 2, `<a:p><a:r><a:t>e${i}</a:t></a:r></a:p>`)).join("");
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(tree) }), { limits: PAGE_SIZED_LIMITS });
    // A withheld slide keeps NOTHING at the reader either (constat N2: its
    // charts are never read, its shapes have no consumer — the tab is one
    // line). Failure values: 500 shapes kept (the pre-N2 contract), or a
    // count of 500 (the notice would lie about what the slide holds).
    expect(slides[0]!.shapes).toHaveLength(0);
    expect(slides[0]!.elementCount).toBe(502);
  });
});

// The deck's USED pictures share the docx reader's 50 MB inflated budget
// (`ooxml/package.ts#IMAGE_MAX_TOTAL_BYTES`, T9 reasoning: the FILE
// downloads under 50 MB and pictures hardly compress, so more is filler).
// Past it the refusal is the CEILING class — a bomb, never degradable.
describe("arbitrage 28/09 — the deck's used pictures stay under the shared 50 MB budget", () => {
  const pic =
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr/></p:pic>';
  const imageRel = (file: string): string => `<Relationship Id="rId2" Type="${REL}/image" Target="../media/${file}"/>`;

  /** The media entries DEFLATED (zeros, a few KB in the archive): only the
   *  INFLATION weighs, which is exactly what the budget bounds — a STORED
   *  52 MB archive OOM'd the test worker for nothing. */
  const withDeflatedMedia = async (source: Uint8Array, sizes: Record<string, number>): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(source);
    for (const [path, size] of Object.entries(sizes)) zip.file(path, new Uint8Array(size), { compression: "DEFLATE" });
    return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  };

  // Re-revue du 28/09 (R1+R2, remplaçant le correctif N2 au remboursement):
  // the withhold is DECIDED on a structural count of the shape tree BEFORE
  // anything inflates — a withheld slide never touches its media, so no
  // refund, no cache purge, no double-counted package `spent`.
  it("a withheld slide inflates NOTHING: the deck publishes even when its media alone would burst the budget", async () => {
    const overCapSlide =
      Array.from({ length: 501 }, (_, i) => shape(i + 2, `<a:p><a:r><a:t>e${i}</a:t></a:r></a:p>`)).join("") + pic;
    // Slide 1 (normal) takes 30 MB of the 50 MB budget FIRST; slide 2 is
    // withheld and its own picture weighs 25 MB — more than what remains.
    // Failure value (the N2 refund's hole, R2): mediaOf inflates slide 2's
    // picture against the 20 MB left, hits the ceiling refusal and fails
    // the whole deck although that slide publishes as one line.
    const bytes = await withDeflatedMedia(
      await packageOf(
        [1, 2],
        { 1: slideXml(pic), 2: slideXml(overCapSlide) },
        { 1: imageRel("a.png"), 2: imageRel("b.png") },
      ),
      { "ppt/media/a.png": 30_000_000, "ppt/media/b.png": 25_000_000 },
    );
    const { slides } = await readPptx(bytes, { limits: PAGE_SIZED_LIMITS });
    const kept = picturesOf(slides[0]!)[0]!;
    expect("image" in kept && kept.image.bytes.length).toBe(30_000_000);
    expect(slides[1]!.elementCount).toBe(502);
    expect(slides[1]!.shapes).toHaveLength(0);
  }, 30_000);

  it("a part SHARED by a withheld and a normal slide is inflated once, by the normal one (R1: no purge, no double spend)", async () => {
    const overCapSlide =
      Array.from({ length: 501 }, (_, i) => shape(i + 2, `<a:p><a:r><a:t>e${i}</a:t></a:r></a:p>`)).join("") + pic;
    // Both slides reference THE SAME 30 MB part. Failure value (R1): the
    // refund purged the shared cache entry, the normal slide re-inflated
    // it, and the package's `spent` counter — never refunded — counted it
    // twice toward PACKAGE_MAX_BYTES.
    const bytes = await withDeflatedMedia(
      await packageOf(
        [1, 2],
        { 1: slideXml(overCapSlide), 2: slideXml(pic) },
        { 1: imageRel("shared.png"), 2: imageRel("shared.png") },
      ),
      { "ppt/media/shared.png": 30_000_000 },
    );
    const { slides } = await readPptx(bytes, { limits: PAGE_SIZED_LIMITS });
    expect(slides[0]!.elementCount).toBe(502);
    const kept = picturesOf(slides[1]!)[0]!;
    expect("image" in kept && kept.image.bytes.length).toBe(30_000_000);
  }, 30_000);

  it("two pictures inflating 52 MB together refuse the deck; one of 26 MB reads whole", async () => {
    const over = await withDeflatedMedia(
      await packageOf([1, 2], { 1: slideXml(pic), 2: slideXml(pic) }, { 1: imageRel("a.png"), 2: imageRel("b.png") }),
      { "ppt/media/a.png": 26_000_000, "ppt/media/b.png": 26_000_000 },
    );
    // Failure value: 52 MB inflated and the deck published — no budget
    // bites. Same `.then(() => "read")` shape as the docx bomb tests: a
    // failed `rejects` would otherwise pretty-print the 52 MB result and
    // OOM the worker (seen live writing this test).
    const outcome: unknown = await readPptx(over, { limits: PAGE_SIZED_LIMITS }).then(
      () => "read",
      (e: unknown) => e,
    );
    expect(outcome).toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid PowerPoint presentation",
    });
    const under = await withDeflatedMedia(await packageOf([1], { 1: slideXml(pic) }, { 1: imageRel("a.png") }), {
      "ppt/media/a.png": 26_000_000,
    });
    const { slides } = await readPptx(under, { limits: PAGE_SIZED_LIMITS });
    const picture = picturesOf(slides[0]!)[0]!;
    expect("image" in picture && picture.image.bytes.length).toBe(26_000_000);
    // 30 s: inflating ~78 MB across the two decks is legitimate work that
    // passed 5 s alone but not under the full check's load (seen 28/09).
  }, 30_000);
});

// ---------------------------------------------------------------------------
// 🔴 Review of the PowerPoint path, 27/09 — the docx hardening's method
// (docx-reader.test.ts « review 27/09 night »): every hostile document is
// built in memory, each fix has its red witness.

describe("review pptx 27/09 finding 1, centralized 28/09 — the reader hands hrefs RAW, except OUR internal forms", () => {
  it("every external target travels verbatim (the scheme guard lives at renderer/storage.ts); drive: and # from a file stay text", async () => {
    const targets: [string, string][] = [
      ["rIdJ", "javascript:alert(1)"],
      ["rIdD", "data:text/html,x"],
      ["rIdF", "file:///C:/secret.txt"],
      ["rIdR", "other.pptx"],
      ["rIdH", "https://example.org/a"],
      ["rIdM", "mailto:someone@example.org"],
      ["rIdV", "drive:1234567890"],
      ["rIdA", "#note-1"],
    ];
    const body = targets
      .map(([id]) => `<a:p><a:r><a:rPr lang="fr-FR"><a:hlinkClick r:id="${id}"/></a:rPr><a:t>${id}</a:t></a:r></a:p>`)
      .join("");
    const rels = targets
      .map(
        ([id, target]) => `<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${target}" TargetMode="External"/>`,
      )
      .join("");
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(shape(2, body)) }, { 1: rels }), {
      limits: PAGE_SIZED_LIMITS,
    });
    // Failure value: the old reader-side filter (no href on the first four),
    // or an internal form crossed from the file (an href on the last two).
    expect(textShapes(slides[0]!)[0]!.paragraphs.map((p) => p.runs[0])).toEqual([
      { text: "rIdJ", href: "javascript:alert(1)" },
      { text: "rIdD", href: "data:text/html,x" },
      { text: "rIdF", href: "file:///C:/secret.txt" },
      { text: "rIdR", href: "other.pptx" },
      { text: "rIdH", href: "https://example.org/a" },
      { text: "rIdM", href: "mailto:someone@example.org" },
      { text: "rIdV" },
      { text: "rIdA" },
    ]);
  });
});

// A chart part and the frame that points at it, for the chart findings.
const CHART_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const chartFrame = (id: number): string =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="f${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic>' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>' +
  "</a:graphicData></a:graphic></p:graphicFrame>";
const cache = (declared: string, points: readonly (readonly [number, string])[]): string =>
  `<c:ptCount val="${declared}"/>` + points.map(([i, v]) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
const chartXml = (
  declared: string,
  cats: readonly (readonly [number, string])[],
  vals: readonly (readonly [number, string])[],
): string =>
  `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart><c:ser>` +
  `<c:cat><c:strRef><c:f>A</c:f><c:strCache>${cache(declared, cats)}</c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>B</c:f><c:numCache>${cache(declared, vals)}</c:numCache></c:numRef></c:val>` +
  "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";

/** What the reader says of the one chart of the one slide of `bytes`: its data, `"withheld"`, or nothing. */
const chartOutcomeOf = async (bytes: Uint8Array) => {
  const { slides } = await readPptx(bytes, { limits: PAGE_SIZED_LIMITS });
  const frame = slides[0]!.shapes.find((s) => s.kind === "graphic");
  if (frame?.kind !== "graphic") throw new Error("no chart frame");
  return frame.chart;
};
/** The same, with a withheld chart as `undefined` (the data tests below). */
const chartOf = async (bytes: Uint8Array) => {
  const read = await chartOutcomeOf(bytes);
  return read === "withheld" ? undefined : read;
};

describe("🔴 review pptx 27/09, finding 2 — a declared c:ptCount never sizes an allocation", () => {
  it('ptCount val="2000000000" with 3 real points gives 3 rows, in bounded time (never new Array(2e9))', async () => {
    const chart = await chartOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "2000000000",
              [
                [0, "a"],
                [1, "b"],
                [2, "c"],
              ],
              [
                [0, "1"],
                [1, "2"],
                [2, "3"],
              ],
            ),
          ),
        },
      ),
    );
    // Failure value: the test never returns (2e9 entries filled), or an OOM.
    expect(chart?.data.rows).toEqual([
      ["a", "1"],
      ["b", "2"],
      ["c", "3"],
    ]);
  });

  // Décision B of the arbitrage of 28/09/2026 REPLACED this test's old
  // expectation: a point really PRESENT past MAX_CHART_POINTS used to be
  // dropped in silence (the old witness froze the loss of "far"); the
  // whole chart is now withheld instead — the frame keeps its marker,
  // never a table that looks complete with rows missing. A LYING ptCount
  // with no point past the bound withholds nothing (the test above).
  it("a point really present past MAX_CHART_POINTS withholds the whole chart; one at the last kept idx does not", async () => {
    const over = await chartOutcomeOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "2000000000",
              [
                [0, "a"],
                [1_999_999_999, "far"],
              ],
              [
                [0, "1"],
                [5, "beyond"],
              ],
            ),
          ),
        },
      ),
    );
    // Failure value: a 6-row table with "far" silently gone (the old rule).
    expect(over).toBe("withheld");
    const atEdge = await chartOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "10000",
              [
                [0, "a"],
                [9_999, "edge"],
              ],
              [
                [0, "1"],
                [9_999, "2"],
              ],
            ),
          ),
        },
      ),
    );
    // Revue HIGH du 28/09, constat N3: rowCount used to span 0..maxIdx, so
    // ONE written point at idx 9 999 fabricated 10 000 near-empty rows
    // (~400 KB of XHTML from a 30-byte cache — two such charts passed the
    // page limit). Only OCCUPIED rows come out now: a row whose label and
    // every value are blank carries nothing and is not fabricated (chosen
    // bias, written); no written label or value is ever lost.
    expect(atEdge?.data.rows).toEqual([
      ["a", "1"],
      ["edge", "2"],
    ]);
  });

  // Review of 01/10/2026 (security): every `c:lvl` of a multi-level cache
  // was kept, one label column per level × up to 10 000 rows — 2 000 levels
  // (102 KB of XML) built 4 002 000 cells, 10 000 levels ran out of memory
  // and the whole document failed. Past the level bound the chart is
  // withheld (décision B), like every other bound of this reader.
  const multiLevelChart = (levels: number, idxOf: (k: number) => number): string => {
    let lv = "";
    for (let k = 0; k < levels; k++) lv += `<c:lvl><c:pt idx="${idxOf(k)}"><c:v>w${k}</c:v></c:pt></c:lvl>`;
    return (
      `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart><c:ser>` +
      `<c:cat><c:multiLvlStrRef><c:f>A</c:f><c:multiLvlStrCache><c:ptCount val="10000"/>${lv}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>` +
      `<c:val><c:numRef><c:f>B</c:f><c:numCache>${cache("1", [[0, "1"]])}</c:numCache></c:numRef></c:val>` +
      "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
    );
  };
  const multiLevelOutcomeOf = async (xml: string) =>
    chartOutcomeOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        { "ppt/charts/chart1.xml": new TextEncoder().encode(xml) },
      ),
    );
  const multiLevelOf = async (xml: string) => {
    const read = await multiLevelOutcomeOf(xml);
    return read === "withheld" ? undefined : read;
  };
  it("a multi-level cache of 2 000 levels withholds the chart, in bounded time; the last allowed level count still publishes every word", async () => {
    const t0 = Date.now();
    const over = await multiLevelOutcomeOf(multiLevelChart(2000, (k) => (k * 7) % 9999));
    // Failure value: a chart of 2 000 + 1 columns (4 002 000 cells).
    expect(over).toBe("withheld");
    expect(Date.now() - t0).toBeLessThan(2000);
    const atEdge = await multiLevelOf(multiLevelChart(8, (k) => k));
    expect(atEdge?.data.headers?.length).toBe(9);
    expect(atEdge?.data.rows.flat().filter((c) => /^w\d$/.test(c))).toHaveLength(8);
  });
  it("a chart whose rows × columns pass the cell bound is withheld before any row is built", async () => {
    const cats = Array.from({ length: 9_999 }, (_, i) => [i, `c${i}`] as const);
    const sers = Array.from(
      { length: 21 },
      () =>
        `<c:ser><c:cat><c:strRef><c:f>A</c:f><c:strCache>${cache("9999", cats)}</c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>B</c:f><c:numCache>${cache(
          "9999",
          cats.map(([i]) => [i, "1"] as const),
        )}</c:numCache></c:numRef></c:val></c:ser>`,
    ).join("");
    const xml = `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart>${sers}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    // Failure value: a 9 999-row, 22-column table (219 978 cells).
    expect(await multiLevelOutcomeOf(xml)).toBe("withheld");
  }, 60_000);

  // Review of 01/10/2026 (security, measured): nothing bounded the NUMBER of
  // series, and the cell bound ran only after every series' dense array was
  // built — 5 000 series of ONE point at idx 9 999 (620 KB of XML) passed it
  // (1 occupied row × 5 001) and cost 13 s / 435 MB. Series are capped at
  // Excel's own 255, checked before any array is built, and a running
  // (columns × longest array) bound stops the build early.
  const sparseSeries = (count: number, idx: number): string => {
    const sers = Array.from(
      { length: count },
      () =>
        `<c:ser><c:cat><c:strRef><c:f>A</c:f><c:strCache>${cache("10000", [[idx, "k"]])}</c:strCache></c:strRef></c:cat>` +
        `<c:val><c:numRef><c:f>B</c:f><c:numCache>${cache("10000", [[idx, "1"]])}</c:numCache></c:numRef></c:val></c:ser>`,
    ).join("");
    return `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart>${sers}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  };
  it("5 000 one-point series are withheld in bounded time and heap; 255 series still publish, 256 do not", async () => {
    const heap0 = process.memoryUsage().heapUsed;
    const t0 = Date.now();
    const over = await multiLevelOutcomeOf(sparseSeries(5000, 9_999));
    const ms = Date.now() - t0;
    const heapMb = (process.memoryUsage().heapUsed - heap0) / 1e6;
    // Failure value: a 1-row, 5 001-column chart after 13 s / 435 MB.
    expect(over).toBe("withheld");
    expect(ms).toBeLessThan(4000); // 0.65 s alone (XML parse included), 13 s before
    expect(heapMb).toBeLessThan(100);
    expect((await multiLevelOf(sparseSeries(255, 0)))?.data.headers?.length).toBe(256);
    expect(await multiLevelOutcomeOf(sparseSeries(256, 0))).toBe("withheld");
  }, 60_000);
  it("the running bound stops a sparse many-series chart before the series cap does", async () => {
    // 200 series at idx 9 999: 201 columns × 10 000 long > 200 000 cells,
    // though only ONE row is occupied. Failure value: a 1-row chart.
    expect(await multiLevelOutcomeOf(sparseSeries(200, 9_999))).toBe("withheld");
    // A legitimate sparse chart is untouched.
    expect((await multiLevelOf(sparseSeries(3, 9_999)))?.data.rows).toEqual([["k", "1", "1", "1"]]);
  });

  // The withhold looks at the points PRESENT, whatever the count declares: a ptCount that lies low, or is missing,
  // must not drop a point that is really there.
  // `c:ptCount` is OPTIONAL in ECMA-376 (CT_NumData/CT_StrData, 0..1): the declared count is a hint and
  // the points PRESENT stand.
  it("a cache with NO ptCount keeps its written points (the bound alone applies); an explicit 0 is a read declaration", async () => {
    const noCount =
      `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart><c:ser>` +
      '<c:cat><c:strRef><c:f>A</c:f><c:strCache><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:f>B</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>' +
      "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
    const chart = await chartOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        { "ppt/charts/chart1.xml": new TextEncoder().encode(noCount) },
      ),
    );
    // Failure value: [] — every point read as « past the declared count ».
    expect(chart?.data.rows).toEqual([
      ["a", "1"],
      ["b", "2"],
      ["c", "3"],
    ]);
  });

  it("a really-present point past the bound withholds the chart even when ptCount lies low", async () => {
    const lyingLow = await chartOutcomeOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          // Declares 3 points but really carries one at idx 15 000: the old
          // order dropped it as « past the declared count » in silence.
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "3",
              [
                [0, "a"],
                [15_000, "far"],
              ],
              [[0, "1"]],
            ),
          ),
        },
      ),
    );
    expect(lyingLow).toBe("withheld");
  });
});

describe("🔴 recensement 28/09, constat 1 — a value whose trailing category cells are blank is not lost", () => {
  it("categories written at idx 0-2, values at idx 0-4 (Excel omits c:pt for blank cells): five rows, no number dropped", async () => {
    const chart = await chartOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "5",
              [
                [0, "a"],
                [1, "b"],
                [2, "c"],
              ],
              [
                [0, "1"],
                [1, "2"],
                [2, "3"],
                [3, "4"],
                [4, "5"],
              ],
            ),
          ),
        },
      ),
    );
    // Failure value: three rows — the values "4" and "5" vanished with the
    // blank category cells (the rows were driven by the categories alone).
    expect(chart?.data.rows).toEqual([
      ["a", "1"],
      ["b", "2"],
      ["c", "3"],
      ["", "4"],
      ["", "5"],
    ]);
  });
});

describe("🔴 review pptx 27/09, finding 7 — the chart picture is the first image relation whose part EXISTS", () => {
  it("rel1's part is absent from the archive, rel2's present: the published image is rel2's", async () => {
    const png = encodePng(1, 1, "rgb-24", [7, 8, 9])!;
    const chart = await chartOf(
      await packageOf(
        [1],
        { 1: slideXml(chartFrame(2)) },
        { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
        {
          "ppt/charts/chart1.xml": new TextEncoder().encode(
            chartXml(
              "2",
              [
                [0, "a"],
                [1, "b"],
              ],
              [
                [0, "1"],
                [1, "2"],
              ],
            ),
          ),
          "ppt/charts/_rels/chart1.xml.rels": new TextEncoder().encode(
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
              `<Relationship Id="rId1" Type="${REL}/image" Target="../media/missing.png"/>` +
              `<Relationship Id="rId2" Type="${REL}/image" Target="../media/present.png"/>` +
              "</Relationships>",
          ),
          "ppt/media/present.png": png,
        },
      ),
    );
    // Failure value: no image at all (the loop broke on the missing part —
    // the header promises « the picture only when the archive holds a
    // rendering »).
    expect(chart?.image?.contentType).toBe("image/png");
    expect([...(chart?.image?.bytes ?? [])]).toEqual([...png]);
  });
});

/** Audit 07/10, P1-F6 point 2. The picture a chart carries is only a FALLBACK (the data is the content): a damaged picture part
 *  must cost that picture, never the chart's own table (it used to withhold the whole chart). */
describe("a chart whose fallback picture part is damaged keeps its data", () => {
  it("the picture stream is corrupt: the chart is read, without a picture", async () => {
    const pixels = Array.from({ length: 32 * 32 }, (_, k) => [(k * 7) % 251, (k * 13) % 241, (k * 29) % 239]).flat();
    const png = encodePng(32, 32, "rgb-24", pixels)!;
    const built = await packageOf(
      [1],
      { 1: slideXml(chartFrame(2)) },
      { 1: `<Relationship Id="rId2" Type="${REL}/chart" Target="../charts/chart1.xml"/>` },
      {
        "ppt/charts/chart1.xml": new TextEncoder().encode(
          chartXml(
            "2",
            [
              [0, "a"],
              [1, "b"],
            ],
            [
              [0, "1"],
              [1, "2"],
            ],
          ),
        ),
        "ppt/charts/_rels/chart1.xml.rels": new TextEncoder().encode(
          '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            `<Relationship Id="rId1" Type="${REL}/image" Target="../media/fallback.png"/></Relationships>`,
        ),
        "ppt/media/fallback.png": png,
      },
    );
    const zip = await JSZip.loadAsync(built);
    const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
    const name = new TextEncoder().encode("ppt/media/fallback.png");
    const view = new DataView(bytes.buffer, bytes.byteOffset);
    let nameAt = -1;
    for (let k = 30; k < bytes.length - name.length; k++) {
      if (!name.every((b, j) => bytes[k + j] === b)) continue;
      if (view.getUint32(k - 30, true) === 0x04034b50 && view.getUint16(k - 30 + 26, true) === name.length) {
        nameAt = k;
        break;
      }
    }
    expect(nameAt).toBeGreaterThan(0);
    const dataStart = nameAt + name.length + view.getUint16(nameAt - 30 + 28, true);
    for (let k = 0; k < 8; k++) bytes[dataStart + k] = (bytes[dataStart + k] ?? 0) ^ 0xff;
    const chart = await chartOf(bytes);
    // Failure value: "withheld" (the whole chart kept out for a damaged picture).
    expect(chart).not.toBe("withheld");
    expect(chart?.data.rows).toEqual([
      ["a", "1"],
      ["b", "2"],
    ]);
    expect(chart?.image).toBeUndefined();
  });
});

describe("🔴 review pptx 27/09, finding 4 — a coordinate past ECMA's bound is treated as not declared", () => {
  it('off x="1e308" and ext cx="1e308" pass Number.isInteger: no box, and no Infinity nor NaN anywhere in the boxes', async () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="1e308" y="0"/><a:ext cx="1e308" cy="1000000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:p><a:r><a:t>far</a:t></a:r></a:p></p:txBody></p:sp>";
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(sp) }), { limits: PAGE_SIZED_LIMITS });
    // Failure value: box { x: 1e308, …, cx: Infinity } (1e308 + 1e308
    // overflows), which the drawing writes as width="Infinity".
    expect(slides[0]!.shapes[0]!.box).toBeUndefined();
  });

  it('p:sldSz cx="1e300" is no slide size', async () => {
    const { slideSize } = await readPptx(
      await packageOf(
        [1],
        { 1: slideXml(shape(2, "<a:p><a:r><a:t>x</a:t></a:r></a:p>")) },
        {},
        {},
        '<p:sldSz cx="1e300" cy="6858000"/>',
      ),
    );
    // Failure value: { cx: 1e300, cy: 6858000 }.
    expect(slideSize).toBeUndefined();
  });

  it('a picture whose ext cx="1e300" declares no width, as it declares no box', async () => {
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="p2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1e300" cy="1000000"/></a:xfrm></p:spPr></p:pic>';
    const { slides } = await readPptx(await packageOf([1], { 1: slideXml(pic) }), { limits: PAGE_SIZED_LIMITS });
    const picture = slides[0]!.shapes[0] as PptxPictureShape;
    expect(picture.box).toBeUndefined();
    // Failure value: 7.874015748031496e+295 (1e300 EMU in points).
    expect(picture.widthPt).toBeUndefined();
  });
});

describe("🔴 review pptx 27/09, finding 5 — table spans are clamped to the declared grid, exactly", () => {
  const tc = (attrs: string, text?: string): string =>
    `<a:tc${attrs}><a:txBody><a:bodyPr/><a:p>${text === undefined ? "" : `<a:r><a:t>${text}</a:t></a:r>`}</a:p></a:txBody><a:tcPr/></a:tc>`;
  const tbl = (columns: number, ...rows: string[]): string =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="t3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic>' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid>' +
    Array.from({ length: columns }, () => '<a:gridCol w="100"/>').join("") +
    `</a:tblGrid>${rows.map((r) => `<a:tr h="100">${r}</a:tr>`).join("")}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

  it('rowSpan="999999999" on a 2-row table is a rowspan of 2; gridSpan="1000000" on the last column falls', async () => {
    const { slides } = await readPptx(
      await packageOf([1], {
        1: slideXml(
          tbl(2, tc(' rowSpan="999999999"', "A") + tc(' gridSpan="1000000"', "B"), tc(' vMerge="1"') + tc("", "C")),
        ),
      }),
    );
    // Failure value: A|rs999999999 and B|cs1000000, written into the page.
    expect(gridOf(tablesOf(slides[0]!)[0]!)).toEqual([["A|rs2", "B"], ["C"]]);
  });

  it("a span is clamped to what remains, never below the declared reality: gridSpan 999 over 3 columns is 3", async () => {
    const { slides } = await readPptx(
      await packageOf([1], {
        1: slideXml(
          tbl(
            3,
            tc(' gridSpan="999"', "A") + tc(' hMerge="1"') + tc(' hMerge="1"'),
            tc("", "B") + tc(' rowSpan="7"', "C") + tc("", "D"),
          ),
        ),
      }),
    );
    // Failure values: A|cs999, or C|rs7 (2 rows only, 1 remaining below row 1).
    expect(gridOf(tablesOf(slides[0]!)[0]!)).toEqual([["A|cs3"], ["B", "C", "D"]]);
  });
});

describe("🔴 review pptx 27/09, finding 6 — buAutoNum startAt is clamped to ECMA's 1..32767", () => {
  it('startAt="-2147483648" reads 1; "999999999" reads 32767; "0" reads 1', async () => {
    const p = (startAt: string, text: string): string =>
      `<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="${startAt}"/></a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
    const { slides } = await readPptx(
      await packageOf([1], {
        1: slideXml(shape(2, p("-2147483648", "low") + p("999999999", "high") + p("0", "zero"))),
      }),
    );
    // Failure value: start -2147483648 and 999999999 written into the IR.
    expect(textShapes(slides[0]!)[0]!.paragraphs.map((p) => p.list)).toEqual([
      { ordered: true, level: 0 },
      { ordered: true, level: 0, start: 32_767 },
      { ordered: true, level: 0 },
    ]);
  });
});
