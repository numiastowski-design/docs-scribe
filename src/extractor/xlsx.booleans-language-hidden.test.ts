/** Product rule « Rien d'invisible » (STATUS 00 ter): what is hidden (sheet, row, column) must not influence what is published.
 *  The language that writes the booleans (VRAI/FAUX vs TRUE/FALSE) is detected from the VISIBLE strings only.
 *  Strings are shared strings (the usual writer), the only form that was read whole, hidden or not.
 *  Run: npx vitest run --project=main src/extractor/xlsx.booleans-language-hidden.test.ts */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractXlsx } from "./xlsx";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const escapeXml = (t: string): string => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const FRENCH_TEXTS = [
  "Le présent tableau récapitule les dépenses de l'équipe commerciale pour le premier trimestre de l'année.",
  "Chaque ligne décrit un déplacement chez un client, avec la date, le motif et le montant remboursé.",
  "Les justificatifs sont conservés par le service comptable et peuvent être consultés sur demande.",
  "Merci de vérifier les montants avant la clôture mensuelle, car aucune correction ne sera acceptée ensuite.",
];
const ENGLISH_TEXTS = [
  "This table summarises the expenses of the sales team for the first quarter of the year.",
  "Each line describes a trip to a customer, with the date, the reason and the amount that was refunded.",
  "The receipts are kept by the accounting department and can be consulted on request.",
  "Please check the amounts before the monthly closing, because no correction will be accepted afterwards.",
];
/** The hidden side outweighs the visible one (three copies), so a detector that sees it decides for it. */
const heavy = (texts: string[]): string[] => [
  ...texts,
  ...texts.map((t) => `${t} (2)`),
  ...texts.map((t) => `${t} (3)`),
];

type Where = "sheet" | "rows" | "cols";

/** Visible text in column A rows 1-4 and booleans on a visible row; the hidden text sits where `where` says. */
const build = async (visible: string[], hidden: string[], where: Where): Promise<Uint8Array> => {
  const strings = [...visible, ...hidden];
  const sst = (i: number): string => `<c r="@" t="s"><v>${i}</v></c>`;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/>` +
      `<sheet name="H" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>` +
      `<Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  );
  zip.file(
    "xl/sharedStrings.xml",
    `${XML_DECL}<sst xmlns="${MAIN_NS}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((t) => `<si><t>${escapeXml(t)}</t></si>`).join("")}</sst>`,
  );
  const total = where === "cols" ? Math.max(visible.length, hidden.length) : visible.length;
  const rows: string[] = Array.from({ length: total }, (_, i) => {
    const a = i < visible.length ? sst(i).replace("@", `A${i + 1}`) : "";
    const b = where === "cols" && i < hidden.length ? sst(visible.length + i).replace("@", `B${i + 1}`) : "";
    return `<row r="${i + 1}">${a}${b}</row>`;
  });
  if (where === "rows")
    hidden.forEach((_, i) =>
      rows.push(
        `<row r="${visible.length + 1 + i}" hidden="1">${sst(visible.length + i).replace("@", `A${visible.length + 1 + i}`)}</row>`,
      ),
    );
  const boolRow = rows.length + 1;
  rows.push(`<row r="${boolRow}"><c r="A${boolRow}" t="b"><v>1</v></c><c r="C${boolRow}" t="b"><v>0</v></c></row>`);
  const cols = where === "cols" ? '<cols><col min="2" max="2" hidden="1" width="9" customWidth="1"/></cols>' : "";
  zip.file(
    "xl/worksheets/sheet1.xml",
    `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${cols}<sheetData>${rows.join("")}</sheetData></worksheet>`,
  );
  const hiddenSheetRows =
    where === "sheet"
      ? hidden.map((_, i) => `<row r="${i + 1}">${sst(visible.length + i).replace("@", `A${i + 1}`)}</row>`)
      : [];
  zip.file(
    "xl/worksheets/sheet2.xml",
    `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>${hiddenSheetRows.join("")}</sheetData></worksheet>`,
  );
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
};

const textsOf = (node: unknown, out: string[] = []): string[] => {
  if (Array.isArray(node)) node.forEach((n) => textsOf(n, out));
  else if (node !== null && typeof node === "object") {
    for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
      if (k === "text" && typeof v === "string") out.push(v);
      else textsOf(v, out);
    }
  }
  return out;
};
const published = async (bytes: Uint8Array, page: "fr" | "en"): Promise<string[]> =>
  textsOf((await extractXlsx(bytes, page)).doc.blocks);

describe("xlsx boolean language ignores what is hidden", () => {
  for (const where of ["sheet", "rows", "cols"] as const) {
    it(`visible English, hidden French (${where}): TRUE / FALSE`, async () => {
      // The page language is French, so only the document's visible text can give TRUE. Failure value: "VRAI".
      const texts = await published(await build(ENGLISH_TEXTS, heavy(FRENCH_TEXTS), where), "fr");
      expect(texts).toContain("TRUE");
      expect(texts).toContain("FALSE");
      expect(texts).not.toContain("VRAI");
    });

    it(`visible French, hidden English (${where}): VRAI / FAUX`, async () => {
      // Failure value: "TRUE".
      const texts = await published(await build(FRENCH_TEXTS, heavy(ENGLISH_TEXTS), where), "en");
      expect(texts).toContain("VRAI");
      expect(texts).toContain("FAUX");
      expect(texts).not.toContain("TRUE");
    });
  }
});
