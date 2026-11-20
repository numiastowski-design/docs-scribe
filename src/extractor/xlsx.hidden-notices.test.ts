import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractXlsx } from "./xlsx";

// 08/10/2026: rows hidden by an autoFilter (hidden, non-empty, inside the filter range, below its header row) stay out
// of the table and the document says so, once, with a count only: { kind: "hidden-dropped", what: "filtered-rows", count }.
// Run: npx vitest run --project=main src/extractor/xlsx.hidden-notices.test.ts

type Sheet = { name: string; state?: "hidden"; rows: string; autoFilter?: string };
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const workbookOf = async (sheets: readonly Sheet[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      `xmlns:r="${REL}"><sheets>` +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"${s.state ? ` state="${s.state}"` : ""}/>`,
        )
        .join("") +
      "</sheets></workbook>",
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
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData>${s.rows}</sheetData>${s.autoFilter !== undefined ? `<autoFilter ref="${s.autoFilter}"/>` : ""}</worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const cell = (ref: string, text: string): string => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
/** Row n with two filled cells "<tag>n-a" / "<tag>n-b"; `hidden` hides the row; `empty` leaves it without values. */
const row = (n: number, tag: string, opts: { hidden?: boolean; empty?: boolean } = {}): string =>
  `<row r="${n}"${opts.hidden ? ' hidden="1"' : ""}>${opts.empty ? "" : cell(`A${n}`, `${tag}${n}a`) + cell(`B${n}`, `${tag}${n}b`)}</row>`;
/** Header on row 1, then rows 2..n of "d" data. `hide` lists hidden row numbers, `empty` the valueless ones. */
const rowsOf = (n: number, hide: readonly number[] = [], empty: readonly number[] = []): string =>
  row(1, "h", { hidden: hide.includes(1) }) +
  Array.from({ length: n - 1 }, (_, i) =>
    row(i + 2, "d", { hidden: hide.includes(i + 2), empty: empty.includes(i + 2) }),
  ).join("");

const run = async (sheets: readonly Sheet[]): Promise<{ hidden: unknown[]; text: string }> => {
  const { doc, notices } = await extractXlsx(await workbookOf(sheets));
  return { hidden: (notices ?? []).filter((n) => n.kind === "hidden-dropped"), text: JSON.stringify(doc.blocks) };
};
const filtered = (count: number): unknown => ({ kind: "hidden-dropped", what: "filtered-rows", count });

describe("xlsx — rows hidden by an autoFilter are counted in one hidden-dropped/filtered-rows remark", () => {
  it("🔴 three hidden non-empty rows inside the filter range: count 3", async () => {
    const r = await run([{ name: "S", rows: rowsOf(8, [3, 4, 6]), autoFilter: "A1:B8" }]);
    expect(r.hidden).toEqual([filtered(3)]);
  });

  it("a hidden EMPTY row inside the range does not count", async () => {
    const r = await run([{ name: "S", rows: rowsOf(6, [3], [3]), autoFilter: "A1:B6" }]);
    expect(r.hidden).toEqual([]);
  });

  it("a row hidden by hand OUTSIDE any filter range: no remark", async () => {
    const r = await run([{ name: "S", rows: rowsOf(10, [9]), autoFilter: "A1:B5" }]);
    expect(r.hidden).toEqual([]);
  });

  it("a sheet without autoFilter and with hidden rows: no remark", async () => {
    const r = await run([{ name: "S", rows: rowsOf(6, [2, 3]) }]);
    expect(r.hidden).toEqual([]);
  });

  it("🔴 two visible sheets with a filter: counts are added in ONE entry", async () => {
    const r = await run([
      { name: "One", rows: rowsOf(6, [2, 3]), autoFilter: "A1:B6" },
      { name: "Two", rows: rowsOf(5, [4]), autoFilter: "A1:B5" },
    ]);
    expect(r.hidden).toEqual([filtered(3)]);
  });

  it("a HIDDEN sheet (state=hidden) with filtered rows does not count", async () => {
    const r = await run([
      { name: "Shown", rows: rowsOf(4) },
      { name: "Secret", state: "hidden", rows: rowsOf(6, [2, 3]), autoFilter: "A1:B6" },
    ]);
    expect(r.hidden).toEqual([]);
  });

  it("the first row of the range (the filter header) hidden does not count", async () => {
    const r = await run([{ name: "S", rows: rowsOf(5, [1]), autoFilter: "A1:B5" }]);
    expect(r.hidden).toEqual([]);
  });

  it("filtered rows are still absent from the published table, the others present (non-regression)", async () => {
    const r = await run([{ name: "S", rows: rowsOf(6, [3, 5]), autoFilter: "A1:B6" }]);
    for (const kept of ["d2a", "d4a", "d6b"]) expect(r.text).toContain(kept);
    for (const gone of ["d3a", "d3b", "d5a"]) expect(r.text).not.toContain(gone);
  });

  it("nothing filtered: no hidden-dropped entry", async () => {
    const r = await run([{ name: "S", rows: rowsOf(5), autoFilter: "A1:B5" }]);
    expect(r.hidden).toEqual([]);
  });
});
