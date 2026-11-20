import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractPptx } from "./pptx";

// 08/10/2026: hidden shapes of a published slide stay out of the page and the document says so, once,
// with a count only: { kind: "hidden-dropped", what: "shape", count }. A group counts for one, never its children.
// Run: npx vitest run --project=main src/extractor/pptx.hidden-notices.test.ts

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

type Slide = { tree: string; show?: string };
const deckOf = async (slides: readonly Slide[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("") +
      "</p:sldIdLst></p:presentation>",
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
      `<?xml version="1.0"?><p:sld ${P_NS}${s.show !== undefined ? ` show="${s.show}"` : ""}><p:cSld><p:spTree>${s.tree}</p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

let nextId = 2;
const shape = (text: string, hidden?: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${nextId++}" name="s${nextId}"${hidden !== undefined ? ` hidden="${hidden}"` : ""}/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const group = (children: string, hidden?: string): string =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId++}" name="g${nextId}"${hidden !== undefined ? ` hidden="${hidden}"` : ""}/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;

const run = async (slides: readonly Slide[]): Promise<{ hidden: unknown[]; text: string }> => {
  const { doc, notices } = await extractPptx(await deckOf(slides));
  return { hidden: (notices ?? []).filter((n) => n.kind === "hidden-dropped"), text: JSON.stringify(doc.blocks) };
};
const shapeNotice = (count: number): unknown => ({ kind: "hidden-dropped", what: "shape", count });

describe("pptx — hidden shapes of published slides are counted in one hidden-dropped/shape remark", () => {
  it("🔴 one hidden shape with text: count 1", async () => {
    const r = await run([{ tree: shape("Visible one") + shape("Secret one", "1") }]);
    expect(r.text).toContain("Visible one");
    expect(r.hidden).toEqual([shapeNotice(1)]);
  });

  it("🔴 two hidden shapes on two slides: ONE entry, count 2", async () => {
    const r = await run([
      { tree: shape("Keep A") + shape("Secret A", "1") },
      { tree: shape("Keep B") + shape("Secret B", "1") },
    ]);
    expect(r.hidden).toEqual([shapeNotice(2)]);
  });

  it("🔴 a hidden group of 3 shapes counts for ONE", async () => {
    const r = await run([{ tree: shape("Keep") + group(shape("g1") + shape("g2") + shape("g3"), "1") }]);
    expect(r.text).not.toContain("g2");
    expect(r.hidden).toEqual([shapeNotice(1)]);
  });

  it("🔴 a visible group with one hidden and one visible child: count 1, the visible text stays", async () => {
    const r = await run([{ tree: group(shape("Child visible") + shape("Child secret", "1")) }]);
    expect(r.text).toContain("Child visible");
    expect(r.hidden).toEqual([shapeNotice(1)]);
  });

  it("no hidden shape: no hidden-dropped entry", async () => {
    const r = await run([{ tree: shape("Alpha") + shape("Beta") }]);
    expect(r.hidden).toEqual([]);
  });

  it("a hidden shape on a HIDDEN slide (show=0): no entry (the whole slide is dropped, not our remark)", async () => {
    const r = await run([{ tree: shape("Shown") }, { show: "0", tree: shape("Off slide") + shape("Off secret", "1") }]);
    expect(r.text).toContain("Shown");
    expect(r.text).not.toContain("Off slide");
    expect(r.hidden).toEqual([]);
  });

  it("the hidden text is still absent from the published blocks (non-regression)", async () => {
    const r = await run([{ tree: shape("Public words") + shape("Concealed words", "1") }]);
    expect(r.text).toContain("Public words");
    expect(r.text).not.toContain("Concealed words");
  });

  it('🔴 hidden="true" counts like "1"; hidden="0" does not count', async () => {
    const r = await run([{ tree: shape("T1", "true") + shape("Z0", "0") + shape("Plain") }]);
    expect(r.text).toContain("Z0");
    expect(r.text).not.toContain("T1");
    expect(r.hidden).toEqual([shapeNotice(1)]);
  });

  // Decision n° 7 (12/10/2026): the « hidden » flag is read tolerantly (blanks and case ignored), as on the Word side.
  it.each(["TRUE", " 1 ", "True", " true\n"])('🔴 hidden="%s" hides the shape like "1"', async (value) => {
    const r = await run([{ tree: shape("Shown words") + shape("Tolerant secret", value) }]);
    expect(r.text).toContain("Shown words");
    expect(r.text).not.toContain("Tolerant secret");
    expect(r.hidden).toEqual([shapeNotice(1)]);
  });

  it.each(["0", "false", " FALSE ", "", "yes"])('hidden="%s" does not hide the shape', async (value) => {
    const r = await run([{ tree: shape("Still shown", value) }]);
    expect(r.text).toContain("Still shown");
    expect(r.hidden).toEqual([]);
  });
});
