import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { encodePng } from "doc-ir/internal";
import { type PptxPictureShape, type PptxShape, type PptxTextShape, readPptx } from "./index";
import { PAGE_SIZED_LIMITS } from "../page-caps.test.support";

// Characterisation of the steps `readPptx` and the shape walk are cut into (slide size, slide part, detail cap, withheld slide, alternate
// content bookkeeping, bodiless arrows). Written green on the code before the cut; each expected value is read from the declared XML.
const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';
const DAMAGED = { code: "EXTRACT_FAILED", message: "the file is damaged or is not a valid PowerPoint presentation" };

const box = (id: number, hidden = false): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"${hidden ? ' hidden="1"' : ""}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>t${id}</a:t></a:r></a:p></p:txBody></p:sp>`;

const slideXml = (tree: string, show?: string): string =>
  `<?xml version="1.0"?><p:sld ${P_NS}${show === undefined ? "" : ` show="${show}"`}><p:cSld><p:spTree>` +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
  `${tree}</p:spTree></p:cSld></p:sld>`;

/** A package whose `p:sldIdLst` names `ppt/slides/slide<n>.xml` in `order`. `presentationRels` replaces the slide relationships. */
const packageOf = async (
  order: readonly number[],
  parts: Readonly<Record<number, string>>,
  extra: {
    presentation?: string;
    slideRels?: Readonly<Record<number, string>>;
    files?: Readonly<Record<string, Uint8Array>>;
    presentationRels?: string;
  } = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>',
  );
  for (const [path, bytes] of Object.entries(extra.files ?? {})) zip.file(path, bytes);
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      order.map((n, i) => `<p:sldId id="${256 + i}" r:id="rId${n}"/>`).join("") +
      `</p:sldIdLst>${extra.presentation ?? ""}</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      (extra.presentationRels ??
        Object.keys(parts)
          .map((n) => `<Relationship Id="rId${n}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`)
          .join("")) +
      "</Relationships>",
  );
  for (const [n, xml] of Object.entries(parts)) zip.file(`ppt/slides/slide${n}.xml`, xml);
  for (const [n, xml] of Object.entries(extra.slideRels ?? {})) {
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xml}</Relationships>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

describe("readPptx stages: the slide size", () => {
  const sizeOf = async (sldSz: string) =>
    (
      await readPptx(await packageOf([1], { 1: slideXml(box(2)) }, { presentation: sldSz }), {
        limits: PAGE_SIZED_LIMITS,
      })
    ).slideSize;

  it("two positive integers within ECMA's bound are a size; the bound itself is allowed, one past it is not", async () => {
    expect(await sizeOf('<p:sldSz cx="9144000" cy="6858000"/>')).toEqual({ cx: 9144000, cy: 6858000 });
    expect(await sizeOf('<p:sldSz cx="27273042316900" cy="1"/>')).toEqual({ cx: 27273042316900, cy: 1 });
    expect(await sizeOf('<p:sldSz cx="1" cy="27273042316900"/>')).toEqual({ cx: 1, cy: 27273042316900 });
    expect(await sizeOf('<p:sldSz cx="27273042316901" cy="1"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="1" cy="27273042316901"/>')).toBeUndefined();
  });

  it("zero, a negative, a fraction, a missing side or no element at all is no size", async () => {
    expect(await sizeOf('<p:sldSz cx="0" cy="5"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="5" cy="0"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="-5" cy="5"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="1.5" cy="5"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="5" cy="1.5"/>')).toBeUndefined();
    expect(await sizeOf('<p:sldSz cx="5"/>')).toBeUndefined();
    expect(await sizeOf("")).toBeUndefined();
  });
});

describe("readPptx stages: the slide part a p:sldId names", () => {
  it("an external relationship (whose raw target happens to name a real part), a relationship of another kind or a part that is no p:sld is damage", async () => {
    const slide = { 1: slideXml(box(2)) };
    const external = await packageOf([1], slide, {
      presentationRels: `<Relationship Id="rId1" Type="${REL}/slide" Target="ppt/slides/slide1.xml" TargetMode="External"/>`,
    });
    await expect(readPptx(external, { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject(DAMAGED);
    const master = await packageOf([1], slide, {
      presentationRels: `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slides/slide1.xml"/>`,
    });
    await expect(readPptx(master, { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject(DAMAGED);
    const unknownId = await packageOf([7], slide, {
      presentationRels: `<Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/>`,
    });
    await expect(readPptx(unknownId, { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject(DAMAGED);
    const notSlide = await packageOf([1], {
      1: `<?xml version="1.0"?><p:sldMaster ${P_NS}><p:cSld><p:spTree/></p:cSld></p:sldMaster>`,
    });
    await expect(readPptx(notSlide, { limits: PAGE_SIZED_LIMITS })).rejects.toMatchObject(DAMAGED);
    const plain = await readPptx(await packageOf([1], slide), { limits: PAGE_SIZED_LIMITS });
    expect(plain.slides).toHaveLength(1);
  });
});

describe("readPptx stages: the detail cap, hidden slides and withheld slides", () => {
  const three = (): Promise<Uint8Array> =>
    packageOf([1, 2, 3, 4], {
      1: slideXml(box(2), "0"),
      2: slideXml(box(2) + box(3, true)),
      3: slideXml(box(2)),
      4: slideXml(box(2)),
    });

  it("slides past the cap are counted, not read; a hidden slide takes no place under the cap", async () => {
    const { slides, listedSlideCount } = await readPptx(await three(), { detailCap: 2, limits: PAGE_SIZED_LIMITS });
    expect(listedSlideCount).toBe(4);
    expect(
      slides.map((s) => [s.number, s.listIndex, s.partPath, s.elementCount, s.shapes.length, s.isTitleRepeated]),
    ).toEqual([
      [1, 1, "ppt/slides/slide2.xml", 2, 1, false],
      [2, 2, "ppt/slides/slide3.xml", 1, 1, false],
      [3, 3, "ppt/slides/slide4.xml", 0, 0, false],
    ]);
    expect(slides[0]!.hiddenShapes).toBe(1);
    expect(slides[1]!.hiddenShapes).toBeUndefined();
    expect(Object.keys(slides[2]!)).toEqual([
      "number",
      "listIndex",
      "partPath",
      "elementCount",
      "isTitleRepeated",
      "shapes",
    ]);
  });

  it("a cap of 1 reads one slide; no cap reads them all", async () => {
    const capped = await readPptx(await three(), { detailCap: 1, limits: PAGE_SIZED_LIMITS });
    expect(capped.slides.map((s) => s.shapes.length)).toEqual([1, 0, 0]);
    const all = await readPptx(await three(), { limits: PAGE_SIZED_LIMITS });
    expect(all.slides.map((s) => s.shapes.length)).toEqual([1, 1, 1]);
  });

  it("a slide over the element cap keeps its place, its true count and no title, and the next slide's title is not suppressed by it", async () => {
    const title = (id: number, text: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const many = Array.from({ length: 501 }, (_, i) => box(i + 2)).join("");
    const { slides } = await readPptx(
      await packageOf([1, 2, 3, 4], {
        1: slideXml(box(2), "0"),
        2: slideXml(title(2, "Same")),
        3: slideXml(title(2, "Same") + many),
        4: slideXml(title(2, "Same")),
      }),
      // The element cap is the PRODUCT's, asked for: the reader holds none of its own since 09/11/2026.
      { limits: PAGE_SIZED_LIMITS },
    );
    expect(
      slides.map((s) => [
        s.number,
        s.listIndex,
        s.partPath,
        s.elementCount,
        s.shapes.length,
        s.title,
        s.isTitleRepeated,
      ]),
    ).toEqual([
      [1, 1, "ppt/slides/slide2.xml", 1, 0, "Same", false],
      [2, 2, "ppt/slides/slide3.xml", 502, 0, undefined, false],
      [3, 3, "ppt/slides/slide4.xml", 1, 0, "Same", false],
    ]);
  });

  it("exactly the element cap is read, one more is withheld", async () => {
    const read = await readPptx(
      await packageOf([1], { 1: slideXml(Array.from({ length: 500 }, (_, i) => box(i + 2)).join("")) }),
      { limits: PAGE_SIZED_LIMITS },
    );
    expect([read.slides[0]!.elementCount, read.slides[0]!.shapes.length]).toEqual([500, 500]);
    const withheld = await readPptx(
      await packageOf([1], { 1: slideXml(Array.from({ length: 501 }, (_, i) => box(i + 2)).join("")) }),
      { limits: PAGE_SIZED_LIMITS },
    );
    expect([withheld.slides[0]!.elementCount, withheld.slides[0]!.shapes.length]).toEqual([501, 0]);
  });

  it("the hidden count and the default text sizes are present only when something was counted or written", async () => {
    const plain = await readPptx(await packageOf([1], { 1: slideXml(box(2)) }), { limits: PAGE_SIZED_LIMITS });
    expect("hiddenShapes" in plain.slides[0]!).toBe(false);
    expect("defaultTextSizes" in plain).toBe(false);
    expect("isThemeDamaged" in plain).toBe(false);
    const sized = await readPptx(
      await packageOf(
        [1],
        { 1: slideXml(box(2)) },
        {
          presentation:
            '<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr><a:lvl9pPr><a:defRPr sz="900"/></a:lvl9pPr></p:defaultTextStyle>',
        },
      ),
    );
    expect(sized.defaultTextSizes).toEqual({ 0: 18, 8: 9 });
    const tenth = await readPptx(
      await packageOf(
        [1],
        { 1: slideXml(box(2)) },
        {
          presentation:
            '<p:defaultTextStyle><a:lvl9pPr><a:defRPr sz="900"/></a:lvl9pPr><a:lvl2pPr><a:defRPr sz="1200"/></a:lvl2pPr></p:defaultTextStyle>',
        },
      ),
      { limits: PAGE_SIZED_LIMITS },
    );
    expect(tenth.defaultTextSizes).toEqual({ 1: 12, 8: 9 });
  });
});

describe("shape walk stages: a bodiless arrow", () => {
  const arrow = (id: number, extra = ""): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="a${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="127000"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>${extra}</p:spPr></p:sp>`;
  const group = (flip: string, members: string): string =>
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm${flip}><a:off x="0" y="0"/><a:ext cx="1270000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="127000"/></a:xfrm></p:grpSpPr>${members}</p:grpSp>`;
  const read = async (tree: string): Promise<PptxShape[]> =>
    (await readPptx(await packageOf([1], { 1: slideXml(tree) }), { limits: PAGE_SIZED_LIMITS })).slides[0]!.shapes;

  it("carries its written fill and line, and the turn of the group that mirrors it", async () => {
    const shapes = await read(
      arrow(
        2,
        '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
      ) +
        arrow(3, "<a:noFill/><a:ln><a:noFill/></a:ln>") +
        group(' flipH="1"', arrow(4)) +
        group("", arrow(5)),
    );
    const [red, none, mirrored, plain] = shapes as PptxTextShape[];
    expect([red!.fill, red!.hasVisibleLine, red!.isTurnedByGroup]).toEqual(["#FF0000", true, undefined]);
    expect([none!.fill, none!.hasVisibleLine, none!.isTurnedByGroup]).toEqual(["none", false, undefined]);
    expect([mirrored!.id, mirrored!.isTurnedByGroup, mirrored!.group]).toEqual(["4", true, 2]);
    expect([plain!.id, plain!.isTurnedByGroup, plain!.group]).toEqual(["5", undefined, 4]);
    expect(red!.paragraphs).toEqual([]);
  });
});

describe("shape walk stages: the ranks and counts of an AlternateContent whose Choice yields nothing", () => {
  const opaque = '<u:frame xmlns:u="urn:example:unknown"><u:data/></u:frame>';
  const picOf = (id: number, embed: string, hidden = false): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"${hidden ? ' hidden="1"' : ""}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${embed}"/><a:stretch/></p:blipFill><p:spPr/></p:pic>`;
  const alternate = (choice: string, fallback: string): string =>
    `<mc:AlternateContent ${MC} ${P14}><mc:Choice Requires="p14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const real = encodePng(
    4,
    2,
    "rgb-24",
    [0, 1].flatMap(() => [0, 1, 2, 3].flatMap((x) => [x, 100, 200])),
  )!;
  const pixel = encodePng(1, 1, "rgb-24", [255, 255, 255])!;
  const slideRels = `<Relationship Id="rIdReal" Type="${REL}/image" Target="../media/real.png"/><Relationship Id="rIdPixel" Type="${REL}/image" Target="../media/pixel.png"/>`;
  const read = async (tree: string) =>
    (
      await readPptx(
        await packageOf(
          [1],
          { 1: slideXml(tree) },
          { slideRels: { 1: slideRels }, files: { "ppt/media/real.png": real, "ppt/media/pixel.png": pixel } },
        ),
        { limits: PAGE_SIZED_LIMITS },
      )
    ).slides[0]!;

  it("the Fallback's ranks replace the Choice's: its picture is ranked from where the alternate stood", async () => {
    const slide = await read(box(2) + alternate(opaque.repeat(3), picOf(3, "rIdReal")) + box(4));
    expect(slide.shapes.map((s) => [s.kind, s.spTreeIndex])).toEqual([
      ["text", 0],
      ["picture", 1],
      ["text", 2],
    ]);
    expect(slide.elementCount).toBe(3);
  });

  it("the Fallback's hidden shapes replace the Choice's in the count", async () => {
    const slide = await read(alternate(box(2, true) + box(3, true), picOf(4, "rIdReal")));
    expect(slide.shapes.map((s) => s.kind)).toEqual(["picture"]);
    expect("hiddenShapes" in slide).toBe(false);
    const both = await read(
      alternate(box(2, true), picOf(4, "rIdReal") + picOf(5, "rIdReal", true) + picOf(6, "rIdReal", true)),
    );
    expect(both.hiddenShapes).toBe(2);
  });

  it("when the Fallback adds nothing, the larger of the two rank counts stands", async () => {
    const slide = await read(alternate(opaque.repeat(3), picOf(3, "rIdPixel")) + box(4));
    expect(slide.shapes.map((s) => [s.kind, s.spTreeIndex])).toEqual([["text", 3]]);
    expect(slide.elementCount).toBe(4);
    const small = await read(
      alternate(opaque, picOf(3, "rIdPixel") + picOf(5, "rIdPixel") + picOf(6, "rIdPixel")) + box(4),
    );
    expect(small.shapes.map((s) => s.spTreeIndex)).toEqual([3]);
    expect(small.elementCount).toBe(4);
  });

  it("only a Fallback made WHOLLY of 1×1 pixels is junk: one real picture beside a pixel keeps both", async () => {
    const slide = await read(alternate(opaque, picOf(3, "rIdPixel") + picOf(4, "rIdReal")));
    const pictures = slide.shapes.filter((s): s is PptxPictureShape => s.kind === "picture");
    expect(pictures.map((p) => p.id)).toEqual(["3", "4"]);
  });
});
