import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { readDocx } from "./docx-reader";
import { encodePng } from "doc-ir/internal";
import { readPptx } from "./pptx-reader";

// Charts and SmartArts where a docx can put a drawing: in a text box, in a
// group, in `mc:AlternateContent`; and a SmartArt whose data model is absent
// (lot 2 of 01/10/2026: no placement may lose their words). The chart
// and the SmartArt are REAL ones (Word's `g-fr-graphique-barres.docx` chart
// part, PowerPoint's SmartArt of `06-fr-la-veille.pptx`) set into minimal
// packages: no Word file with these placements is among the fixtures.
vi.setConfig({ testTimeout: 30_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = (path: string): Buffer => readFileSync(join(here, "../../fixtures", path));
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"`;

const chartGraphic =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic>';
const smartGraphic =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic>';
const inline = (graphic: string): string =>
  `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/>${graphic}</wp:inline></w:drawing>`;
const inTextBox = (inner: string): string =>
  `<w:drawing><wp:anchor><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="100" cy="100"/><wp:docPr id="2" name="t"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>inside box</w:t></w:r></w:p><w:p><w:r>${inner}</w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing>`;
const inGroup = (graphic: string): string =>
  `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="3" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">` +
  `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wpg:graphicFrame><wpg:cNvPr id="4" name="f"/><wpg:cNvFrPr/><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm>${graphic}</wpg:graphicFrame></wpg:wgp>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
const alternate = (drawing: string): string =>
  `<mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>`;

interface Parts {
  body: string;
  /** SmartArt data model kept (with its text), emptied of text, or absent. */
  model?: "text" | "empty" | "absent";
  chart?: boolean;
  /** The chart part's own text, in place of Word's real one (a withheld or damaged chart). */
  chartXml?: string;
  /** A PNG part `rIdP`, for a picture: 1×1 pixel by default, 4×2 with `large`. */
  image?: boolean;
  large?: boolean;
}

const build = async (parts: Parts): Promise<Uint8Array> => {
  const pptx = await JSZip.loadAsync(fixtures("diagrammes-reel/06-fr-la-veille.pptx"));
  const data = await pptx.file("ppt/diagrams/data1.xml")!.async("string");
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const model = parts.model ?? "text";
  const rels = [
    `<Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>`,
    ...(model === "absent" ? [] : [`<Relationship Id="rId2" Type="${REL}/diagramData" Target="diagrams/data1.xml"/>`]),
    ...(parts.image ? [`<Relationship Id="rIdP" Type="${REL}/image" Target="media/p.png"/>`] : []),
    ...(parts.chart ? [`<Relationship Id="rIdC" Type="${REL}/chart" Target="charts/chart1.xml"/>`] : []),
  ];
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${parts.body}</w:body></w:document>`);
  if (model === "text") zip.file("word/diagrams/data1.xml", data);
  if (model === "empty") zip.file("word/diagrams/data1.xml", data.replace(/<a:t>[^<]*<\/a:t>/g, "<a:t></a:t>"));
  zip.file("word/diagrams/drawing1.xml", await pptx.file("ppt/diagrams/drawing1.xml")!.async("string"));
  if (parts.image)
    zip.file(
      "word/media/p.png",
      parts.large
        ? encodePng(4, 2, "rgb-24", Array.from({ length: 8 }, () => [1, 2, 3]).flat())!
        : encodePng(1, 1, "rgb-24", [1, 2, 3])!,
    );
  if (parts.chart) {
    const docx = await JSZip.loadAsync(fixtures("diagrammes-generes/g-fr-graphique-barres.docx"));
    zip.file(
      "word/charts/chart1.xml",
      parts.chartXml ?? (await docx.file("word/charts/chart1.xml")!.async("nodebuffer")),
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

const blocksOfKind = (value: unknown, kind: string, out: IRBlock[] = []): IRBlock[] => {
  if (Array.isArray(value)) for (const v of value) blocksOfKind(v, kind, out);
  else if (typeof value === "object" && value !== null) {
    if ((value as { kind?: string }).kind === kind) out.push(value as IRBlock);
    for (const v of Object.values(value)) blocksOfKind(v, kind, out);
  }
  return out;
};

const smartText = (blocks: readonly IRBlock[]): string =>
  blocksOfKind(blocks, "diagramText")
    .flatMap((b) => (b.kind === "diagramText" ? b.sections.flatMap((s) => s.inlines.map((i) => i.text)) : []))
    .join("");

const SMART_WORDS = ["Mise sous surveillance des sources et mots-clés", "Repères de signaux significatifs"];

// A chart the reader withholds (a bound, a damaged part) is told apart from « no chart »: its title and labels are not
// on the page, so a remark says so, like the pptx path's notice (review of 01/10/2026).
describe("a withheld chart in a docx is reported, never silently dropped", () => {
  const CHART_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
  const seriesChart = (count: number): string =>
    `<?xml version="1.0"?><c:chartSpace ${CHART_NS}><c:chart><c:plotArea><c:barChart>` +
    Array.from(
      { length: count },
      () =>
        '<c:ser><c:val><c:numRef><c:f>B</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>',
    ).join("") +
    `</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const body = `<w:p><w:r>${inline(chartGraphic)}</w:r></w:p>`;

  it("a chart of 256 series (past the cap) leaves one « chart withheld » notice and no chart block", async () => {
    const { doc, notices } = await readDocx(await build({ body, chart: true, chartXml: seriesChart(256) }));
    // Failure value (before): no notice at all.
    expect(`${blocksOfKind(doc.blocks, "chart").length} chart(s), notices ${JSON.stringify(notices)}`).toBe(
      '0 chart(s), notices [{"kind":"chart-withheld","count":1}]',
    );
  });

  it("255 series still publish, with no notice", async () => {
    const { doc, notices } = await readDocx(await build({ body, chart: true, chartXml: seriesChart(255) }));
    expect(`${blocksOfKind(doc.blocks, "chart").length} chart(s), notices ${JSON.stringify(notices)}`).toBe(
      "1 chart(s), notices undefined",
    );
  });

  it("a damaged chart part leaves the same notice, and the document stands", async () => {
    const { doc, notices } = await readDocx(
      await build({
        body: `${body}<w:p><w:r><w:t>after</w:t></w:r></w:p>`,
        chart: true,
        chartXml: "<c:chartSpace <<< not xml",
      }),
    );
    expect(`${blocksOfKind(doc.blocks, "chart").length} chart(s), notices ${JSON.stringify(notices)}`).toBe(
      '0 chart(s), notices [{"kind":"chart-withheld","count":1}]',
    );
    expect(JSON.stringify(doc.blocks)).toContain("after");
  });

  // 07/10/2026: a missing part is said, the same notice in every format (PowerPoint says it too).
  it("🔴 a chart whose part is missing (relationship absent) leaves the same notice", async () => {
    const { doc, notices } = await readDocx(await build({ body: `${body}<w:p><w:r><w:t>after</w:t></w:r></w:p>` }));
    expect(`${blocksOfKind(doc.blocks, "chart").length} chart(s), notices ${JSON.stringify(notices)}`).toBe(
      '0 chart(s), notices [{"kind":"chart-withheld","count":1}]',
    );
    expect(JSON.stringify(doc.blocks)).toContain("after");
  });

  it("a chart with no series is « no chart »: nothing to publish, nothing to report", async () => {
    const { notices } = await readDocx(await build({ body, chart: true, chartXml: seriesChart(0) }));
    expect(notices).toBeUndefined();
  });
});

describe("chart and SmartArt placements in a docx", () => {
  for (const [name, wrap] of [
    ["a text box", () => inTextBox(inline(chartGraphic))],
    ["a group", () => inGroup(chartGraphic)],
  ] as const) {
    it(`a chart in ${name} is published as a chart block, once`, async () => {
      const { doc } = await readDocx(await build({ body: `<w:p><w:r>${wrap()}</w:r></w:p>`, chart: true }));
      // Failure value (before): 0 chart blocks.
      expect(`${blocksOfKind(doc.blocks, "chart").length} chart block(s)`).toBe("1 chart block(s)");
    });
  }

  it("a SmartArt in a text box and one in a group are both published", async () => {
    const body = `<w:p><w:r>${inTextBox(inline(smartGraphic))}</w:r></w:p><w:p><w:r>${inGroup(smartGraphic)}</w:r></w:p>`;
    const { doc } = await readDocx(await build({ body }));
    const text = smartText(doc.blocks);
    // Failure value (before): 0 occurrences of each.
    expect(`${text.split(SMART_WORDS[1]!).length - 1} occurrences`).toBe("2 occurrences");
  });

  it("a SmartArt and a chart inside mc:AlternateContent (the Choice is read) are published once each", async () => {
    const body = `<w:p><w:r>${alternate(inline(smartGraphic))}</w:r></w:p><w:p><w:r>${alternate(inline(chartGraphic))}</w:r></w:p>`;
    const { doc } = await readDocx(await build({ body, chart: true }));
    expect(`${smartText(doc.blocks).split(SMART_WORDS[1]!).length - 1} SmartArt`).toBe("1 SmartArt");
    expect(`${blocksOfKind(doc.blocks, "chart").length} chart(s)`).toBe("1 chart(s)");
  });

  it("a SmartArt whose data model part is absent is read from its drawing cache, named by the document's diagramDrawing relationship", async () => {
    const { doc } = await readDocx(
      await build({ body: `<w:p><w:r>${inline(smartGraphic)}</w:r></w:p>`, model: "absent" }),
    );
    const text = smartText(doc.blocks);
    // Failure value (before): no diagram text at all.
    for (const w of SMART_WORDS) expect(text).toContain(w);
  });

  it("a SmartArt WITH its model is not also read from its cache (the words come out once)", async () => {
    const body = `<w:p><w:r>${inline(smartGraphic)}</w:r></w:p>`;
    const { doc } = await readDocx(await build({ body, model: "text" }));
    expect(`${smartText(doc.blocks).split(SMART_WORDS[1]!).length - 1} occurrence`).toBe("1 occurrence");
  });
});

describe("a SmartArt without a data model in a pptx", () => {
  it("the cache named by the slide's diagramDrawing relationship gives its words", async () => {
    const zip = await JSZip.loadAsync(fixtures("diagrammes-reel/06-fr-la-veille.pptx"));
    zip.remove("ppt/diagrams/data1.xml");
    const relsPath = "ppt/slides/_rels/slide7.xml.rels";
    const rels = (await zip.file(relsPath)!.async("string")).replace(/<Relationship Id="rId2"[^>]*\/>/, "");
    zip.file(relsPath, rels);
    const deck = await readPptx(await zip.generateAsync({ type: "uint8array" }));
    const words = deck.slides.flatMap((s) =>
      s.shapes.flatMap((sh) => (sh.kind === "graphic" && sh.words ? sh.words : [])),
    );
    // Failure value (before): no words for slide 7's SmartArt.
    expect(words.join("|")).toContain("Mise sous surveillance des sources et mots-clés");
    expect(words.join("|")).toContain("Repères de signaux significatifs");
  });
});

describe("the picture a chart drawing offers instead of itself", () => {
  const pic =
    '<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="9" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdP"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';

  it("a Choice we do not read (an extended chart, cx1) leaves its image-only Fallback published (the same object, drawn by Office); a readable chart keeps its data table and its Fallback is not read", async () => {
    const extended = `<mc:AlternateContent><mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"><w:drawing/></mc:Choice><mc:Fallback>${pic}</mc:Fallback></mc:AlternateContent>`;
    const readable = `<mc:AlternateContent><mc:Choice Requires="wps">${inline(chartGraphic)}</mc:Choice><mc:Fallback>${pic}</mc:Fallback></mc:AlternateContent>`;
    const { doc, images } = await readDocx(
      await build({
        body: `<w:p><w:r>${extended}</w:r></w:p><w:p><w:r>${readable}</w:r></w:p>`,
        chart: true,
        image: true,
        large: true,
      }),
    );
    expect(
      `${blocksOfKind(doc.blocks, "chart").length} chart(s), ${blocksOfKind(doc.blocks, "image").length} image(s), ${images.length} file(s)`,
    ).toBe("1 chart(s), 1 image(s), 1 file(s)");
  });

  it("an unread Choice whose image-only Fallback is a 1×1 pixel publishes nothing (junk, not content)", async () => {
    const extended = `<mc:AlternateContent><mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"><w:drawing/></mc:Choice><mc:Fallback>${pic}</mc:Fallback></mc:AlternateContent>`;
    const { doc } = await readDocx(
      await build({ body: `<w:p><w:r>${extended}</w:r></w:p>`, chart: true, image: true }),
    );
    // Failure value: 1 image (the pixel published).
    expect(blocksOfKind(doc.blocks, "image")).toHaveLength(0);
  });

  // M16 (02/10/2026): the Fallback of an unread Choice is published ONLY when it carries no text.
  it("an unread Choice whose Fallback carries text publishes NOTHING (the article's trap); never both branches", async () => {
    const withText = `<mc:AlternateContent><mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"><w:drawing/></mc:Choice><mc:Fallback>${pic}<w:t>old words</w:t></mc:Fallback></mc:AlternateContent>`;
    const { doc, images } = await readDocx(
      await build({ body: `<w:p><w:r>${withText}</w:r></w:p>`, chart: true, image: true }),
    );
    // Failure value: 1 image (the Fallback published although it carries words).
    expect(
      `${blocksOfKind(doc.blocks, "image").length} image(s), ${images.length} file(s), ${JSON.stringify(doc.blocks).includes("old words")}`,
    ).toBe("0 image(s), 0 file(s), false");
  });
});
