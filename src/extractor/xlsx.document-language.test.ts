import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractXlsx } from "./xlsx";

// Decision of 08/10/2026: "keep the Drive document's format for everything". An .xlsx stores no display language,
// so the decimal/thousands separators, month names and the built-in short date depend on a language we must pick.
// Priority: (1) a format code that NAMES its language ([$-40C]...) keeps it; (2) else the language OF THE DOCUMENT,
// detected from its cell texts (French or English only); (3) else (numbers only, undetected, abstention) the page language.
// Google Sheets is out of scope. Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/xlsx.document-language.test.ts
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
  let cells = `<c r="A${r}" s="1"><v>${NUMBER}</v></c><c r="B${r}" s="2"><v>${DATE_SERIAL}</v></c>`;
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

const FR_NUMBER = /^1[\s  ]234,50$/;
const EN_NUMBER = "1,234.50";
const FR_DATE = "15/03/2024";
const EN_DATE = "3/15/24";

const expectFrench = (texts: string[]): void => {
  expect(texts.some((t) => FR_NUMBER.test(t))).toBe(true);
  expect(texts).toContain(FR_DATE);
  expect(texts).not.toContain(EN_NUMBER);
  expect(texts).not.toContain(EN_DATE);
};
const expectEnglish = (texts: string[]): void => {
  expect(texts).toContain(EN_NUMBER);
  expect(texts).toContain(EN_DATE);
  expect(texts.some((t) => FR_NUMBER.test(t))).toBe(false);
  expect(texts).not.toContain(FR_DATE);
};

describe("xlsx — numbers and dates follow the language of the document, then the page", () => {
  it("🔴 L1: a French-texted workbook is published the French way even when the page is English", async () => {
    // Failure value (before): "1,234.50" and "3/15/24" (page language).
    expectFrench(await published(await build({ texts: FRENCH_TEXTS }), "en"));
  });

  it("🔴 L2: an English-texted workbook is published the English way even when the page is French", async () => {
    // Failure value (before): "1 234,50" and "15/03/2024" (page language).
    expectEnglish(await published(await build({ texts: ENGLISH_TEXTS }), "fr"));
  });

  it("non-regression: document and page sharing the language give that language", async () => {
    expectFrench(await published(await build({ texts: FRENCH_TEXTS }), "fr"));
    expectEnglish(await published(await build({ texts: ENGLISH_TEXTS }), "en"));
  });

  it("non-regression L3: a workbook of numbers only follows the page language, fr and en each", async () => {
    const bytes = await build({ texts: [] });
    expectFrench(await published(bytes, "fr"));
    expectEnglish(await published(bytes, "en"));
  });

  it("non-regression L4: [$-40C] keeps French month names in an English workbook, whatever the page", async () => {
    const bytes = await build({
      texts: ENGLISH_TEXTS,
      custom: [{ id: 164, code: "[$-40C]dddd d mmmm yyyy" }],
      namedDateCells: 1,
    });
    for (const lang of ["fr", "en"] as const) expect(await published(bytes, lang)).toContain("vendredi 15 mars 2024");
  });

  it("non-regression L4: [$-409] keeps English month names in a French workbook, whatever the page", async () => {
    const bytes = await build({
      texts: FRENCH_TEXTS,
      custom: [{ id: 164, code: "[$-409]dddd d mmmm yyyy" }],
      namedDateCells: 1,
    });
    for (const lang of ["fr", "en"] as const) expect(await published(bytes, lang)).toContain("Friday 15 March 2024");
  });

  it("🔴 L5: a workbook dominated by French with a few English words takes French", async () => {
    // Failure value (before): English formatting under an "en" page.
    expectFrench(await published(await build({ texts: [...FRENCH_TEXTS, "Total", "Budget", "Meeting notes"] }), "en"));
  });

  it("🔴 L6: on a French workbook the output is identical for page 'en' and page 'fr'", async () => {
    const bytes = await build({ texts: FRENCH_TEXTS });
    const en = await published(bytes, "en");
    expect(en.some((t) => FR_NUMBER.test(t))).toBe(true); // the equality is not between two raw outputs
    expect(en).toEqual(await published(bytes, "fr"));
  });

  it("🔴 L6: same on an English workbook, page 'fr' versus page 'en'", async () => {
    const bytes = await build({ texts: ENGLISH_TEXTS });
    const fr = await published(bytes, "fr");
    expect(fr).toContain(EN_NUMBER);
    expect(fr).toEqual(await published(bytes, "en"));
  });

  it("non-regression L7: a workbook in another language (German) falls back to the page language", async () => {
    const bytes = await build({ texts: GERMAN_TEXTS });
    expectFrench(await published(bytes, "fr"));
    expectEnglish(await published(bytes, "en"));
  });

  it("non-regression L8: detection adds and loses no word, the text cells are published as written", async () => {
    for (const [texts, page] of [
      [FRENCH_TEXTS, "en"],
      [ENGLISH_TEXTS, "fr"],
      [GERMAN_TEXTS, "en"],
    ] as const) {
      const out = await published(await build({ texts }), page);
      for (const t of texts) expect(out).toContain(t);
      expect(out).toHaveLength(texts.length + 2); // the texts, plus the formatted number and date
    }
  });
});
