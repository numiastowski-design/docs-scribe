import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_NOTICE_SHEET_NAME_CHARS, MAX_NOTICE_SHEET_NAMES } from "doc-ir/internal";
import type { IRBlock } from "doc-ir";
import { extractXlsx } from "./xlsx";

// Decision of 08/10/2026: a missing piece is said, the rest is published. Case: xl/workbook.xml names a
// sheet, its relationship exists, but the worksheet PART is absent from the archive. Said as one `sheet-unreadable`
// notice (same shape and bounds as `workbook-cap`); the other sheets are published intact.
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/xlsx.missing-sheet.test.ts
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

interface Sheet {
  name: string;
  word: string;
  /** The part is not written in the archive (workbook and relationship still name it). */
  absent?: boolean;
  hidden?: boolean;
}

const workbookOf = async (sheets: readonly Sheet[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      "</Types>",
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets>` +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${s.name}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`,
        )
        .join("") +
      "</sheets></workbook>",
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `${XML_DECL}<Relationships xmlns="${PKG_REL}">` +
      sheets
        .map(
          (_, i) => `<Relationship Id="rId${i + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join("") +
      "</Relationships>",
  );
  sheets.forEach((s, i) => {
    if (s.absent) return;
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${s.word}</t></is></c></row></sheetData></worksheet>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

interface Unreadable {
  kind: "sheet-unreadable";
  sheets: string[];
  more: number;
}
const unreadable = (notices: readonly { kind: string }[] | undefined): Unreadable[] =>
  (notices ?? []).filter((n) => n.kind === "sheet-unreadable") as unknown as Unreadable[];

const tabTitles = (blocks: readonly IRBlock[]): string[] => {
  const tabs = blocks[0] as Extract<IRBlock, { kind: "tabs" }> | undefined;
  return tabs?.kind === "tabs" ? tabs.tabs.map((t) => t.title) : [];
};

describe("xlsx — a sheet whose part is absent from the archive is said, the other sheets are published", () => {
  it("non-regression: an intact workbook has no sheet-unreadable notice (X5)", async () => {
    const { notices, doc } = await extractXlsx(
      await workbookOf([
        { name: "Alpha", word: "alphaword" },
        { name: "Beta", word: "betaword" },
      ]),
    );
    expect(unreadable(notices)).toEqual([]);
    expect(tabTitles(doc.blocks)).toEqual(["Alpha", "Beta"]);
  });

  it("non-regression: the other sheets of a workbook with one absent part are published intact (tabs and values), nothing throws (X1)", async () => {
    const bytes = await workbookOf([
      { name: "Alpha", word: "alphaword" },
      { name: "Gone", word: "goneword", absent: true },
      { name: "Gamma", word: "gammaword" },
    ]);
    const { doc } = await extractXlsx(bytes);
    expect(tabTitles(doc.blocks)).toEqual(["Alpha", "Gamma"]);
    expect(JSON.stringify(doc)).toContain("alphaword");
    expect(JSON.stringify(doc)).toContain("gammaword");
  });

  it("🔴 X1: one absent part gives no tab for it and exactly one sheet-unreadable notice naming it", async () => {
    const bytes = await workbookOf([
      { name: "Alpha", word: "alphaword" },
      { name: "Gone", word: "goneword", absent: true },
      { name: "Gamma", word: "gammaword" },
    ]);
    const { notices, doc } = await extractXlsx(bytes);
    expect(tabTitles(doc.blocks)).not.toContain("Gone");
    // Failure value (before): no notice at all (the tab is skipped in silence).
    expect(unreadable(notices)).toEqual([{ kind: "sheet-unreadable", sheets: ["Gone"], more: 0 }]);
  });

  it("🔴 X1: the absent sheet being the first one still says it", async () => {
    const { notices } = await extractXlsx(
      await workbookOf([
        { name: "Gone", word: "goneword", absent: true },
        { name: "Beta", word: "betaword" },
      ]),
    );
    expect(unreadable(notices)).toEqual([{ kind: "sheet-unreadable", sheets: ["Gone"], more: 0 }]);
  });

  it("🔴 X2: two absent parts give ONE entry with both names, in workbook order", async () => {
    const bytes = await workbookOf([
      { name: "Zulu", word: "zuluword", absent: true },
      { name: "Alpha", word: "alphaword" },
      { name: "Mike", word: "mikeword", absent: true },
    ]);
    const { notices, doc } = await extractXlsx(bytes);
    expect(tabTitles(doc.blocks)).toEqual(["Alpha"]);
    expect(unreadable(notices)).toEqual([{ kind: "sheet-unreadable", sheets: ["Zulu", "Mike"], more: 0 }]);
  });

  it("🔴 X3: more absent parts than MAX_NOTICE_SHEET_NAMES give the first names then `more` = the rest", async () => {
    const absentCount = MAX_NOTICE_SHEET_NAMES + 3;
    const names = Array.from({ length: absentCount }, (_, i) => `Gone${i + 1}`);
    const bytes = await workbookOf([
      { name: "Alpha", word: "alphaword" },
      ...names.map((name) => ({ name, word: "w", absent: true })),
    ]);
    const { notices } = await extractXlsx(bytes);
    const entries = unreadable(notices);
    expect(entries).toHaveLength(1);
    expect(entries[0]!.sheets).toEqual(names.slice(0, MAX_NOTICE_SHEET_NAMES));
    expect(entries[0]!.more).toBe(absentCount - MAX_NOTICE_SHEET_NAMES);
  });

  it("🔴 X4: a very long sheet name is cut like workbook-cap's, and the notice stays small", async () => {
    const long = `Q${"n".repeat(99)}`;
    const { notices } = await extractXlsx(
      await workbookOf([
        { name: "Alpha", word: "alphaword" },
        { name: long, word: "w", absent: true },
      ]),
    );
    const entries = unreadable(notices);
    expect(entries).toHaveLength(1);
    expect(entries[0]!.sheets).toHaveLength(1);
    expect(entries[0]!.sheets[0]!.length).toBeLessThanOrEqual(MAX_NOTICE_SHEET_NAME_CHARS + 1); // + the ellipsis
    expect(entries[0]!.sheets[0]!.startsWith("Qnnn")).toBe(true);
    expect(entries[0]!.more).toBe(0);
    expect(JSON.stringify(entries[0]).length).toBeLessThan(1000);
  });

  it("non-regression: an absent part of a HIDDEN sheet is never said (a hidden sheet is never read) (X5)", async () => {
    const bytes = await workbookOf([
      { name: "Alpha", word: "alphaword" },
      { name: "Secret", word: "secretword", absent: true, hidden: true },
    ]);
    const { notices, doc } = await extractXlsx(bytes);
    expect(tabTitles(doc.blocks)).toEqual(["Alpha"]);
    expect(unreadable(notices)).toEqual([]);
  });

  it("🔴 X5 control: a visible absent sheet next to a hidden absent one names only the visible one", async () => {
    const bytes = await workbookOf([
      { name: "Alpha", word: "alphaword" },
      { name: "Secret", word: "secretword", absent: true, hidden: true },
      { name: "Gone", word: "goneword", absent: true },
    ]);
    const { notices } = await extractXlsx(bytes);
    expect(unreadable(notices)).toEqual([{ kind: "sheet-unreadable", sheets: ["Gone"], more: 0 }]);
  });
});
