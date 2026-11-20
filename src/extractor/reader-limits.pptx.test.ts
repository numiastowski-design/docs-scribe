import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { isDeclaredFormatSource } from "../pipeline/conservation-source";
import { readPptxConservationSource } from "../pipeline/source-reading-pptx";
import { parseXml } from "./ooxml/xml";
import { readShapes } from "./pptx-reader/shapes";
import { encodePng } from "doc-ir/internal";
import { extractPptx } from "./pptx";

// Plan step 5: a deck is read under the limits it is given (`ReaderLimits`), the product's when it is given none, and the extraction and its
// independent second reading take the SAME limits. Run: npx vitest run --project=main src/extractor/reader-limits.pptx.test.ts

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

interface MadeSlide {
  tree: string;
  rels?: string;
}

const deckOf = async (
  slides: readonly MadeSlide[],
  files: Readonly<Record<string, Uint8Array | string>> = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>',
  );
  for (const [path, content] of Object.entries(files)) zip.file(path, content);
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("")}</p:sldIdLst></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      slides
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  slides.forEach((s, i) => {
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${s.tree}</p:spTree></p:cSld></p:sld>`,
    );
    if (s.rels) zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, RELS(s.rels));
  });
  return zip.generateAsync({ type: "uint8array" });
};

const text = (id: number, words: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
const frame = (uri: string, data: string): string =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="f2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;
const tc = (words: string): string =>
  `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
const tr = (...cells: string[]): string => `<a:tr h="100">${cells.join("")}</a:tr>`;
const table = (): string =>
  frame(
    "http://schemas.openxmlformats.org/drawingml/2006/table",
    `<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>${tr(tc("alpha"), tc("bravo"), tc("charlie"))}${tr(tc("delta"), tc("echo"), tc("foxtrot"))}</a:tbl>`,
  );

type Limits = NonNullable<Parameters<typeof extractPptx>[3]>;
const tabsOf = (blocks: readonly { kind: string }[]): { title: string }[] =>
  blocks.flatMap((b) => (b.kind === "tabs" ? (b as unknown as { tabs: { title: string }[] }).tabs : []));
const wordsOfSource = async (bytes: Uint8Array, limits?: Limits): Promise<{ words: string[]; excluded: number[] }> => {
  const source = await readPptxConservationSource(bytes, "en", limits);
  if (!isDeclaredFormatSource(source)) throw new Error("the second reading failed");
  return {
    words: source.units.flatMap((u) => u.tokens.map((t) => t.text)),
    excluded: [...source.excludedUnits].sort((a, b) => a - b),
  };
};

describe("pptx — the display cap of tabs is the limit it is given", () => {
  const slides = Array.from({ length: 5 }, (_, i) => ({ tree: text(2, `slide${i + 1}word`) }));

  it("five slides are five tabs under the product's cap", async () => {
    expect(tabsOf((await extractPptx(await deckOf(slides))).doc.blocks)).toHaveLength(5);
  });

  it("presentationTabs: 3 keeps three tabs and says which slides are left out (the extraction)", async () => {
    // FAILS WITH (limits ignored): expected [ { …(5) }, … ] to have a length of 3 but got 5.
    const result = await extractPptx(await deckOf(slides), undefined, "en", { presentationTabs: 3 });
    expect(tabsOf(result.doc.blocks)).toHaveLength(3);
    expect(result.notices).toContainEqual({ kind: "tab-cap", from: 4, to: 5 });
  });

  it("the second reading excludes the same slides under the SAME limits", async () => {
    const bytes = await deckOf(slides);
    expect((await wordsOfSource(bytes)).excluded).toEqual([]);
    // FAILS WITH (limits ignored): expected [] to deeply equal [ 3, 4 ].
    const limited = await wordsOfSource(bytes, { presentationTabs: 3 });
    expect(limited.excluded).toEqual([3, 4]);
    expect(limited.words).not.toContain("slide4word");
  });
});

describe("pptx — the element cap of a slide is the limit it is given", () => {
  const tree = Array.from({ length: 5 }, (_, i) => text(i + 2, `e${i}`)).join("");
  const bytes = deckOf([{ tree }, { tree: text(2, "next") }]);

  it("the extraction withholds the slide whole under slideElements: 4 and keeps it by default", async () => {
    const whole = await extractPptx(await bytes);
    expect(tabsOf(whole.doc.blocks)).toHaveLength(2);
    // FAILS WITH (limits ignored): expected [ { …(5) }, { …(5) } ] to have a length of 1 but got 2.
    const limited = await extractPptx(await bytes, undefined, "en", { slideElements: 4 });
    expect(tabsOf(limited.doc.blocks)).toHaveLength(1);
    expect(limited.notices).toContainEqual({ kind: "slide-withheld", slide: 1, elementCount: 5 });
  });

  it("the second reading withholds it under the SAME limits", async () => {
    expect((await wordsOfSource(await bytes)).excluded).toEqual([]);
    expect((await wordsOfSource(await bytes, { slideElements: 4 })).excluded).toEqual([0]);
  });

  it("the shape walk counts but never reads an element at the limit it is given", () => {
    const walk = (limits?: Limits) =>
      readShapes(parseXml(`<p:spTree ${P_NS}>${tree}</p:spTree>`)!, {
        relationships: new Map(),
        media: new Map(),
        sources: {},
        limits,
      });
    expect(walk().shapes).toHaveLength(5);
    // FAILS WITH (limits ignored): expected [ { … }, … ] to have a length of 4 but got 5.
    expect(walk({ slideElements: 4 }).shapes).toHaveLength(4);
    expect(walk({ slideElements: 4 }).elementCount).toBe(5);
  });
});

describe("pptx — the table caps are the limits it is given", () => {
  const bytes = deckOf([{ tree: table() + text(9, "kilo") }]);
  const flat = (r: Awaited<ReturnType<typeof extractPptx>>): string => JSON.stringify(r.doc.blocks);

  it("a 3-column table is published by default and withheld, and said, under tableCols: 2", async () => {
    expect(flat(await extractPptx(await bytes))).toContain("foxtrot");
    // FAILS WITH (limits ignored): expected '[{"kind":"tabs", … "foxtrot" …' not to contain 'foxtrot'.
    const limited = await extractPptx(await bytes, undefined, "en", { tableCols: 2 });
    expect(flat(limited)).not.toContain("foxtrot");
    expect(flat(limited)).toContain("kilo");
    expect(limited.notices).toContainEqual({ kind: "table-withheld", count: 1 });
  });

  it("the second reading expects none of its words under the SAME limits", async () => {
    expect((await wordsOfSource(await bytes)).words).toContain("foxtrot");
    const limited = await wordsOfSource(await bytes, { tableCols: 2 });
    expect(limited.words).not.toContain("foxtrot");
    expect(limited.words).toContain("kilo");
  });
});

describe("pptx — the byte allowance of the media is the limit it is given", () => {
  const png = encodePng(1, 1, "rgb-24", [1, 2, 3])!;
  const pic =
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr/></p:pic>';
  const bytes = deckOf([{ tree: pic, rels: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/p.png"/>` }], {
    "ppt/media/p.png": png,
  });

  it("a picture is read under the product's allowance and refused (never inflated) past the one given", async () => {
    expect((await extractPptx(await bytes)).images).toHaveLength(1);
    // FAILS WITH (limits ignored): Promise resolved "{ doc: …, images: [ …(1) ], … }" instead of rejecting.
    await expect(extractPptx(await bytes, undefined, "en", { payloadBytes: png.length - 1 })).rejects.toBeInstanceOf(
      ReaderError,
    );
    expect((await extractPptx(await bytes, undefined, "en", { payloadBytes: png.length })).images).toHaveLength(1);
  });
});
