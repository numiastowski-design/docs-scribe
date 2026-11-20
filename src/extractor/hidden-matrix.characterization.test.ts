import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { DocumentNotice } from "doc-ir";
import { extractDocx } from "./docx";
import { extractGoogleDoc } from "./google-doc";
import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";
import { extractPptx } from "./pptx";
import { extractXlsx } from "./xlsx";

// CHARACTERIZATION tests: the behaviour already existed, these tests pin it. No red run before; instead each assertion
// was flipped to its failure value once in this file and seen red (see the session report).
// What they pin: the "cases without test" of docs/audit-2026-10-07/matrice-masque.md (hidden / folded / out-of-view content,
// per format), read against the CODE of 09/10/2026, not the matrix: where the code says otherwise, the test says the code.
// Each test is named with its matrix id (H<n> <format>) and asserts what is published (hidden text ABSENT, visible text
// PRESENT) AND the exact list of "hidden-dropped" notices.
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/hidden-matrix.characterization.test.ts
vi.setConfig({ testTimeout: 30_000 });

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const hiddenOf = (notices: readonly DocumentNotice[] | undefined): DocumentNotice[] =>
  (notices ?? []).filter((n) => n.kind === "hidden-dropped");

// ---- docx (shape of docx-reader.hidden-notices.test.ts, copied)
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${REL}"`;
const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  zip.file("word/_rels/document.xml.rels", RELS(""));
  const a4 =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}${a4}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};
const run = (t: string, rPr = ""): string => `<w:r>${rPr}<w:t>${t}</w:t></w:r>`;
const para = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const cellOf = (inner: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const rowOf = (cells: string[], trPr = ""): string => `<w:tr>${trPr}${cells.map(cellOf).join("")}</w:tr>`;
const tableOf = (rows: string[]): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join("")}</w:tbl>`;
const DATE = 'w:author="a" w:date="2026-01-01T00:00:00Z"';

describe("docx / Google Doc — hidden table rows, tables, format changes, folded content, suggestions", () => {
  it("H2 docx table row with w:trPr/w:hidden: the row is NOT read as hidden, its text is published, no notice", async () => {
    const body = tableOf([
      rowOf([para(run("rowA1")), para(run("rowA2"))]),
      rowOf([para(run("hiddenrowB1")), para(run("hiddenrowB2"))], "<w:trPr><w:hidden/></w:trPr>"),
    ]);
    const { doc, notices } = await extractDocx(await docxOf(body));
    const text = JSON.stringify(doc.blocks);
    for (const t of ["rowA1", "rowA2", "hiddenrowB1", "hiddenrowB2"]) expect(text).toContain(t);
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H2 docx table whose runs are vanish (w:vanish): the hidden words are absent, the visible cell is present, no notice", async () => {
    const vanish = "<w:rPr><w:vanish/></w:rPr>";
    const body = tableOf([rowOf([para(run("vanishedwords", vanish)), para(run("shown cell"))])]);
    const { doc, notices } = await extractDocx(await docxOf(body));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("shown cell");
    expect(text).not.toContain("vanishedwords");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H11 docx rPrChange / pPrChange: format-change records are not text; the paragraph text is published once, no notice", async () => {
    const body =
      `<w:p><w:pPr><w:pPrChange w:id="1" ${DATE}><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange></w:pPr>` +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="2" ${DATE}><w:rPr/></w:rPrChange></w:rPr><w:t>formatchanged</w:t></w:r></w:p>`;
    const { doc, notices } = await extractDocx(await docxOf(body));
    expect(JSON.stringify(doc.blocks).split("formatchanged")).toHaveLength(2);
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H17 docx content control with w14:collapsed: its content is published, no notice", async () => {
    const body = `<w:sdt><w:sdtPr><w14:collapsed w14:val="1"/></w:sdtPr><w:sdtContent>${para(run("foldedcontrol"))}</w:sdtContent></w:sdt>${para(run("after"))}`;
    const { doc, notices } = await extractDocx(await docxOf(body));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("foldedcontrol");
    expect(text).toContain("after");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H11 Google Doc suggestions (export carries w:ins / w:del): insertion published, deletion absent and counted once", async () => {
    const ins = `<w:ins w:id="1" ${DATE}><w:r><w:t>suggestedadd</w:t></w:r></w:ins>`;
    const del = `<w:del w:id="2" ${DATE}><w:r><w:delText>suggestedcut</w:delText></w:r></w:del>`;
    const { doc, notices } = await extractGoogleDoc(await docxOf(para(run("base "), ins, del)));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("suggestedadd");
    expect(text).not.toContain("suggestedcut");
    expect(hiddenOf(notices)).toEqual([{ kind: "hidden-dropped", what: "tracked-deletion", count: 1 }]);
  });
});

// ---- xlsx (shape of xlsx.hidden-notices.test.ts, copied, with extra parts)
type Sheet = {
  name: string;
  state?: string;
  rows: string;
  tail?: string;
  cols?: string;
  parts?: Record<string, string>;
  sheetRels?: string;
};
const xlsxOf = async (sheets: readonly Sheet[], definedNames = ""): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets>` +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"${s.state ? ` state="${s.state}"` : ""}/>`,
        )
        .join("") +
      `</sheets>${definedNames}</workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  sheets.forEach((s, i) => {
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">${s.cols ?? ""}<sheetData>${s.rows}</sheetData>${s.tail ?? ""}</worksheet>`,
    );
    if (s.sheetRels !== undefined) zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, RELS(s.sheetRels));
    for (const [path, xml] of Object.entries(s.parts ?? {})) zip.file(path, xml);
  });
  return zip.generateAsync({ type: "uint8array" });
};
const cell = (ref: string, text: string): string => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const xrow = (n: number, a: string, b: string, attrs = ""): string =>
  `<row r="${n}"${attrs}>${cell(`A${n}`, a)}${cell(`B${n}`, b)}</row>`;

describe("xlsx — veryHidden tab, folded groups, comments, print setup", () => {
  it("H5 xlsx veryHidden tab: nothing of it is published (end to end), the visible tab is, no hidden-dropped notice", async () => {
    const { doc, notices } = await extractXlsx(
      await xlsxOf([
        { name: "Shown", rows: xrow(1, "shownA", "shownB") },
        { name: "Secret", state: "veryHidden", rows: xrow(1, "verysecretA", "verysecretB") },
      ]),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("shownA");
    expect(text).not.toContain("verysecret");
    expect(text).not.toContain("Secret");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H8 xlsx grouped rows and columns folded in the file (outlineLevel + collapsed, NOT hidden): published, no notice", async () => {
    const rows =
      `<row r="1" outlineLevel="0">${cell("A1", "head1")}${cell("B1", "head2")}</row>` +
      `<row r="2" outlineLevel="1">${cell("A2", "groupedrow")}${cell("B2", "g2")}</row>` +
      `<row r="3" collapsed="1">${cell("A3", "summaryrow")}${cell("B3", "s2")}</row>`;
    const cols = '<cols><col min="2" max="2" width="10" outlineLevel="1" collapsed="1"/></cols>';
    const { doc, notices } = await extractXlsx(await xlsxOf([{ name: "S", rows, cols }]));
    const text = JSON.stringify(doc.blocks);
    for (const t of ["head2", "groupedrow", "g2", "summaryrow", "s2"]) expect(text).toContain(t);
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H8 xlsx grouped rows that the file marks hidden=1 (what Excel writes when folded): absent, no notice (outside any filter)", async () => {
    const rows =
      xrow(1, "head1", "head2") +
      xrow(2, "foldedrow", "foldedb", ' outlineLevel="1" hidden="1"') +
      xrow(3, "summaryrow", "s2", ' collapsed="1"');
    const { doc, notices } = await extractXlsx(await xlsxOf([{ name: "S", rows }]));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("summaryrow");
    expect(text).not.toContain("foldedrow");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H9 xlsx cell comments (comments part linked from the sheet): the comment text is not published, the cell is", async () => {
    const comments =
      '<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList>' +
      '<comment ref="A1" authorId="0"><text><r><t>commentbodysecret</t></r></text></comment></commentList></comments>';
    const { doc, notices } = await extractXlsx(
      await xlsxOf([
        {
          name: "S",
          rows: xrow(1, "cellvalue", "b1"),
          parts: { "xl/comments1.xml": comments },
          sheetRels: `<Relationship Id="rId1" Type="${REL}/comments" Target="../comments1.xml"/>`,
        },
      ]),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("cellvalue");
    expect(text).not.toContain("commentbodysecret");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H12 xlsx print header/footer text and H16 print area: header/footer text not published; cells OUTSIDE the print area are published", async () => {
    const tail =
      "<headerFooter><oddHeader>&amp;Lprintheadertext</oddHeader><oddFooter>&amp;Rprintfootertext</oddFooter></headerFooter>";
    const names =
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S!$A$1:$B$1</definedName></definedNames>';
    const { doc, notices } = await extractXlsx(
      await xlsxOf([{ name: "S", rows: xrow(1, "insidearea", "b1") + xrow(2, "outsidearea", "b2"), tail }], names),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("insidearea");
    expect(text).toContain("outsidearea");
    expect(text).not.toContain("printheadertext");
    expect(text).not.toContain("printfootertext");
    expect(hiddenOf(notices)).toEqual([]);
  });
});

// ---- pptx (shape of pptx.hidden-notices.test.ts, copied, with layout / comments / sections)
const P_NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const sp = (text: string, id: number, off = "0", runAttrs = ""): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${off}" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr${runAttrs}/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const deckOf = async (
  trees: readonly string[],
  opts: { sldSz?: string; presExtra?: string; withLayout?: boolean; withComments?: boolean } = {},
): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      trees.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("") +
      `</p:sldIdLst>${opts.sldSz ?? ""}${opts.presExtra ?? ""}</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      trees
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  trees.forEach((tree, i) => {
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree>${tree}</p:spTree></p:cSld></p:sld>`,
    );
    const rels =
      (opts.withLayout
        ? `<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`
        : "") +
      (opts.withComments ? `<Relationship Id="rId2" Type="${REL}/comments" Target="../comments/comment1.xml"/>` : "");
    if (rels) zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, RELS(rels));
  });
  if (opts.withLayout) {
    zip.file(
      "ppt/slideLayouts/slideLayout1.xml",
      `<?xml version="1.0"?><p:sldLayout ${P_NS}><p:cSld><p:spTree>${sp("layoutfurniture", 99)}</p:spTree></p:cSld></p:sldLayout>`,
    );
    zip.file(
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
      RELS(`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`),
    );
    zip.file(
      "ppt/slideMasters/slideMaster1.xml",
      `<?xml version="1.0"?><p:sldMaster ${P_NS}><p:cSld><p:spTree>${sp("masterfurniture", 98)}</p:spTree></p:cSld></p:sldMaster>`,
    );
  }
  if (opts.withComments) {
    zip.file(
      "ppt/comments/comment1.xml",
      `<?xml version="1.0"?><p:cmLst ${P_NS}><p:cm authorId="0" dt="2026-01-01T00:00:00" idx="1"><p:pos x="0" y="0"/><p:text>slidecommentsecret</p:text></p:cm></p:cmLst>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
};

describe("pptx / Google Slides — master and layout shapes, off-slide text, sections, comments", () => {
  it("H13 pptx: shapes of the slide layout and of the master are not published, the slide shape is, no notice", async () => {
    const { doc, notices } = await extractPptx(await deckOf([sp("slidetextkept", 2)], { withLayout: true }));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("slidetextkept");
    expect(text).not.toContain("layoutfurniture");
    expect(text).not.toContain("masterfurniture");
    expect(hiddenOf(notices)).toEqual([]);
  });

  // Changed on 09/10/2026 (rule « Rien d'invisible », decision (e)): a shape wholly outside the slide is no longer
  // published; still no notice (it is not « unexpected »).
  it("H16 pptx: text of a shape placed wholly outside the slide is NOT published, no notice", async () => {
    const sldSz = '<p:sldSz cx="9144000" cy="6858000"/>';
    const { doc, notices } = await extractPptx(
      await deckOf([sp("insidetext", 2) + sp("offslidetext", 3, "90000000")], { sldSz }),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("insidetext");
    expect(text).not.toContain("offslidetext");
    expect(hiddenOf(notices)).toEqual([]);
  });

  // Characterisation of the four edges of the « wholly outside the slide » test (EMU, slide 9 144 000 × 6 858 000).
  it("H16b pptx: a top-level shape is left out on each of the four sides when wholly outside, and kept when it touches the slide", async () => {
    const box = (text: string, id: number, x: number, y: number, cx: number, cy: number, rot = ""): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const cases: readonly (readonly [string, string, boolean])[] = [
      // [name, shape, published]
      ["leftout", box("leftout", 2, -200, 0, 100, 100), false],
      ["leftedge", box("leftedge", 3, -100, 0, 100, 100), false],
      ["leftpart", box("leftpart", 4, -50, 0, 100, 100), true],
      ["rightout", box("rightout", 5, 9144000, 0, 100, 100), false],
      ["rightin", box("rightin", 6, 9143999, 0, 100, 100), true],
      ["topout", box("topout", 7, 0, -200, 100, 100), false],
      ["topedge", box("topedge", 8, 0, -100, 100, 100), false],
      ["toppart", box("toppart", 9, 0, -50, 100, 100), true],
      ["bottomout", box("bottomout", 10, 0, 6858000, 100, 100), false],
      ["bottomin", box("bottomin", 11, 0, 6857999, 100, 100), true],
      ["zeroedge", box("zeroedge", 12, 0, 0, 0, 0), true],
      ["zeroleft", box("zeroleft", 13, -1, 0, 0, 0), false],
      ["turned", box("turned", 14, -200, 0, 100, 100, ' rot="5400000"'), true],
    ];
    const sldSz = '<p:sldSz cx="9144000" cy="6858000"/>';
    const { doc } = await extractPptx(await deckOf([cases.map(([, shape]) => shape).join("")], { sldSz }));
    const text = JSON.stringify(doc.blocks);
    expect(cases.map(([name]) => [name, text.includes(name)])).toEqual(
      cases.map(([name, , published]) => [name, published]),
    );
  });

  it("H17 pptx: slides of a collapsed p14 section are published (sectionLst not read), no notice", async () => {
    const presExtra =
      '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
      '<p14:section name="Folded" id="{11111111-1111-1111-1111-111111111111}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>';
    const { doc, notices } = await extractPptx(
      await deckOf([sp("firstslide", 2), sp("foldedsectionslide", 3)], { presExtra }),
    );
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("firstslide");
    expect(text).toContain("foldedsectionslide");
    expect(text).not.toContain("Folded");
    expect(hiddenOf(notices)).toEqual([]);
  });

  it("H9 pptx slide comments (comments part linked from the slide): the comment text is not published", async () => {
    const { doc, notices } = await extractPptx(await deckOf([sp("slidetextkept", 2)], { withComments: true }));
    const text = JSON.stringify(doc.blocks);
    expect(text).toContain("slidetextkept");
    expect(text).not.toContain("slidecommentsecret");
    expect(hiddenOf(notices)).toEqual([]);
  });
});

describe("Google Sheet — display format, row/column groups and cell notes are not read", () => {
  it("H1/H8/H9 Google Sheet: a value is published as the API gives it; dimension groups and cell notes in the metadata change nothing, no notice", async () => {
    const sheet = {
      properties: { title: "S", index: 0 },
      rowMetadata: [{}, {}],
      rowGroups: [{ range: { startIndex: 1, endIndex: 2 }, depth: 1, collapsed: true }],
      columnGroups: [{ range: { startIndex: 1, endIndex: 2 }, depth: 1, collapsed: true }],
      data: [{ rowData: [{ values: [{ note: "cellnotesecret" }] }] }],
    };
    const meta = { sheets: [sheet] } as unknown as GoogleSheetMeta;
    const { doc, notices } = await extractGoogleSheet(
      meta,
      new Map([
        [
          "S",
          [
            ["h1", "h2"],
            ["groupedvalue", "g2"],
          ],
        ],
      ]),
    );
    const text = JSON.stringify(doc.blocks);
    for (const t of ["h2", "groupedvalue", "g2"]) expect(text).toContain(t);
    expect(text).not.toContain("cellnotesecret");
    expect(hiddenOf(notices)).toEqual([]);
  });
});
