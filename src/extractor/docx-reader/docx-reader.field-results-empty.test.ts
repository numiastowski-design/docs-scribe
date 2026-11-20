/** Matrix « dates / nombres » (docs/audit-2026-10-07/SUIVI.md), decision (09/10/2026): a Word field that carries NO
 *  computed result (a complex field begin/instr/end without « separate », or an empty w:fldSimple) shows nothing in Word and
 *  is published empty, as today — but the document now SAYS so on the dashboard, as an .xlsx does for a formula without a
 *  cached value (`formula-cells-empty`). One notice per document, its count = such fields in the body.
 *  Run: npx vitest run --project=main src/extractor/docx-reader/docx-reader.field-results-empty.test.ts */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractDocx } from "../docx";
import { PAGE_SIZED_LIMITS } from "../page-caps.test.support";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL}"`;
const docxOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
  zip.file("word/_rels/document.xml.rels", RELS(""));
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "uint8array" });
};

/** A complex field with no « separate »: nothing to show. */
const noResultField =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> = 2 + 2 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
/** A complex field WITH a result: shown, no notice. */
const resultField =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> = 3 + 3 </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>6</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
const notices = async (body: string) => {
  const { doc, notices: all } = await extractDocx(await docxOf(body), PAGE_SIZED_LIMITS);
  return { text: JSON.stringify(doc.blocks), empty: (all ?? []).filter((n) => n.kind === "field-results-empty") };
};

describe("Word: a field without a computed value is published empty AND said", () => {
  it("two fields without result (one complex, one empty fldSimple): one notice, count 2; the text around is published", async () => {
    const r = await notices(
      `<w:p><w:r><w:t>Total :</w:t></w:r>${noResultField}</w:p>` +
        '<w:p><w:r><w:t>Date :</w:t></w:r><w:fldSimple w:instr=" DATE "/></w:p>',
    );
    expect(r.text).toContain("Total :");
    expect(r.text).toContain("Date :");
    // Failure value: [] (published empty, nothing said).
    expect(r.empty).toEqual([{ kind: "field-results-empty", count: 2 }]);
  });

  it("a field with its result: the result is published, no notice", async () => {
    const r = await notices(`<w:p><w:r><w:t>Total :</w:t></w:r>${resultField}</w:p>`);
    expect(r.text).toContain("6");
    expect(r.empty).toEqual([]);
  });

  // 09/10/2026, final high review of the day.
  it("a field without result inside a table that is not published (over 21 columns) is not counted", async () => {
    const cells = Array.from(
      { length: 30 },
      (_, c) => `<w:tc><w:p>${c === 0 ? noResultField : `<w:r><w:t>c${c}</w:t></w:r>`}</w:p></w:tc>`,
    ).join("");
    const r = await notices(`<w:p><w:r><w:t>intro</w:t></w:r></w:p><w:tbl><w:tr>${cells}</w:tr></w:tbl>`);
    // Failure value: [{ kind: "field-results-empty", count: 1 }] for a field the page never publishes.
    expect(r.empty).toEqual([]);
  });

  it("fields that never have a result by nature (XE index entry, TC contents entry, PRIVATE: Word's « cold » fields) are not counted", async () => {
    const cold = (instr: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const r = await notices(
      `<w:p><w:r><w:t>Mot</w:t></w:r>${cold('XE "Mot"')}${cold('TC "Chapitre"')}${cold("PRIVATE data")}</w:p>`,
    );
    // Source: Microsoft, WdFieldKind.wdFieldKindCold « a field that doesn't have a result, for example XE, TC or Private ».
    // Failure value: [{ kind: "field-results-empty", count: 3 }].
    expect(r.empty).toEqual([]);
  });

  it("TA (table-of-authorities entry, ECMA-376 §17.16.5) is cold like XE; a field merely STARTING with TA is not", async () => {
    const cold = (instr: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const ta = await notices(`<w:p><w:r><w:t>Cas</w:t></w:r>${cold('TA \\l "Smith v. Jones" \\s "Smith"')}</w:p>`);
    // Failure value: [{ kind: "field-results-empty", count: 1 }].
    expect(ta.empty).toEqual([]);
    const lookalikes = await notices(`<w:p><w:r><w:t>Cas</w:t></w:r>${cold("TARGET x")}${cold("TAB y")}</w:p>`);
    // The guard: a name that only starts with TA keeps being reported (failure value: []).
    expect(lookalikes.empty).toEqual([{ kind: "field-results-empty", count: 2 }]);
  });

  it("the cold names are matched without case, and so is the letter after them: « xe », « Private » are cold, « Tab », « xEnd » are not", async () => {
    const field = (instr: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const cold = await notices(`<w:p><w:r><w:t>Mot</w:t></w:r>${field('xe "Mot"')}${field("Private data")}</w:p>`);
    // Failure value: [{ kind: "field-results-empty", count: 2 }].
    expect(cold.empty).toEqual([]);
    const lookalikes = await notices(`<w:p><w:r><w:t>Mot</w:t></w:r>${field("Tab y")}${field("xEnd z")}</w:p>`);
    // Failure value: [] — a lower-case letter after the name must keep it a lookalike.
    expect(lookalikes.empty).toEqual([{ kind: "field-results-empty", count: 2 }]);
  });
});
