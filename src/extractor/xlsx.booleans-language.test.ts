/** Matrix « dates / nombres » (docs/audit-2026-10-07/SUIVI.md), decision (09/10/2026): an .xlsx boolean is written in
 *  the language of the DOCUMENT, as its numbers and dates already are (08/10/2026): Excel shows VRAI/FAUX to a French reader,
 *  TRUE/FALSE to an English one; the file stores only 1/0 (`t="b"`). Same priority as numbers: the document's detected
 *  language (French or English), else the page language.
 *  Run: npx vitest run --project=main src/extractor/xlsx.booleans-language.test.ts */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractXlsx } from "./xlsx";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const escapeXml = (t: string): string => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// 15 March 2024 (a Friday), serial 1900 system.
const DATE_SERIAL = 45366;
const NUMBER = 1234.5;

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
const GERMAN_TEXTS = [
  "Diese Tabelle fasst die Ausgaben des Vertriebsteams für das erste Quartal des Jahres zusammen.",
  "Jede Zeile beschreibt eine Reise zu einem Kunden, mit dem Datum, dem Grund und dem erstatteten Betrag.",
  "Die Belege werden von der Buchhaltung aufbewahrt und können auf Anfrage eingesehen werden.",
  "Bitte überprüfen Sie die Beträge vor dem Monatsabschluss, da danach keine Korrektur mehr angenommen wird.",
];

/** Style indexes: 1 = built-in 4 (#,##0.00), 2 = built-in 14 (short date), 3 = custom 164. */
const stylesXml = (custom: { id: number; code: string }[]): string =>
  `${XML_DECL}<styleSheet xmlns="${MAIN_NS}">` +
  (custom.length > 0
    ? `<numFmts count="${custom.length}">${custom.map((c) => `<numFmt numFmtId="${c.id}" formatCode="${escapeXml(c.code)}"/>`).join("")}</numFmts>`
    : "") +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>';

interface Spec {
  texts: string[];
  custom?: { id: number; code: string }[];
  /** Extra date cells (style 3 = the custom format) placed after the number and the built-in date. */
  namedDateCells?: number;
}

const build = async (spec: Spec): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.file("xl/styles.xml", stylesXml(spec.custom ?? []));
  const rows: string[] = spec.texts.map(
    (t, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${escapeXml(t)}</t></is></c></row>`,
  );
  const r = spec.texts.length + 1;
  let cells = `<c r="A${r}" s="1"><v>${NUMBER}</v></c><c r="B${r}" s="2"><v>${DATE_SERIAL}</v></c><c r="Z${r}" t="b"><v>1</v></c><c r="Y${r}" t="b"><v>0</v></c>`;
  for (let i = 0; i < (spec.namedDateCells ?? 0); i++)
    cells += `<c r="${String.fromCharCode(67 + i)}${r}" s="3"><v>${DATE_SERIAL}</v></c>`;
  rows.push(`<row r="${r}">${cells}</row>`);
  zip.file(
    "xl/worksheets/sheet1.xml",
    `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>${rows.join("")}</sheetData></worksheet>`,
  );
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
};

/** Every published text string of the document, in order (all `text` fields of the IR). */
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
const published = async (bytes: Uint8Array, lang: "fr" | "en"): Promise<string[]> =>
  textsOf((await extractXlsx(bytes, lang)).doc.blocks);

describe("xlsx booleans follow the document's language", () => {
  it("a French document: VRAI / FAUX", async () => {
    const texts = await published(await build({ texts: FRENCH_TEXTS }), "en");
    // Failure value: "TRUE" / "FALSE".
    expect(texts).toContain("VRAI");
    expect(texts).toContain("FAUX");
    expect(texts).not.toContain("TRUE");
  });

  it("an English document: TRUE / FALSE", async () => {
    const texts = await published(await build({ texts: ENGLISH_TEXTS }), "fr");
    expect(texts).toContain("TRUE");
    expect(texts).toContain("FALSE");
    expect(texts).not.toContain("VRAI");
  });

  it("a language we do not detect (German): the page language decides", async () => {
    expect(await published(await build({ texts: GERMAN_TEXTS }), "fr")).toContain("VRAI");
    expect(await published(await build({ texts: GERMAN_TEXTS }), "en")).toContain("TRUE");
  });
});
