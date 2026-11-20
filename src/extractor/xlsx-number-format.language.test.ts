/**
 * The page language on a spreadsheet's numbers and dates
 * (features/page-language.feature, 30/09/2026). The table
 * below IS the design's §3 table: code, value, the English page, the French
 * page. A locale written INSIDE a code (`[$-40C]`) governs the WORDS only
 * (month and day names); separators always follow the page language.
 */
import { describe, expect, it } from "vitest";

import { formatGeneralNumber, formatNumberCell } from "./xlsx-number-format";

const NBSP = " ";

// [code, value, English page, French page] — §3 of the design, row for row.
const TABLE: [code: string, value: number, en: string, fr: string][] = [
  ["#,##0.00", 1234.5, "1,234.50", `1${NBSP}234,50`],
  ["#,##0.00", -1234.5, "-1,234.50", `-1${NBSP}234,50`],
  ["#,##0.00 [$€-40C]", 1234.5, "1,234.50 €", `1${NBSP}234,50 €`],
  ["[$€-40C]#,##0.00", 1234.5, "€1,234.50", `€1${NBSP}234,50`],
  ["0.0%", 0.125, "12.5%", "12,5%"],
  ["dddd d mmmm yyyy", 45000, "Wednesday 15 March 2023", "mercredi 15 mars 2023"],
  ["d-mmm-yy", 45000, "15-Mar-23", "15-mars-23"],
  ["[$-409]d mmmm yyyy", 45000, "15 March 2023", "15 March 2023"],
  ["[$-40C]d mmmm yyyy", 45000, "15 mars 2023", "15 mars 2023"],
  ["hh:mm", 0.75, "18:00", "18:00"],
  ["h:mm AM/PM", 0.75, "6:00 PM", "6:00 PM"],
  ["0.00E+00", 12345, "1.23E+04", "1,23E+04"],
];

describe("formatNumberCell — the design's §3 table, English page and French page", () => {
  for (const [code, value, en, fr] of TABLE) {
    it(`${JSON.stringify(code)} on ${value} → EN "${en}"`, () => {
      expect(formatNumberCell(value, code, false, "en")).toBe(en);
    });
    it(`${JSON.stringify(code)} on ${value} → FR "${fr}"`, () => {
      expect(formatNumberCell(value, code, false, "fr")).toBe(fr);
    });
  }

  it("General: 1/3 → 0.333333333 (EN) and 0,333333333 (FR)", () => {
    expect(formatNumberCell(1 / 3, "General", false, "en")).toBe("0.333333333");
    expect(formatNumberCell(1 / 3, "General", false, "fr")).toBe("0,333333333");
    expect(formatNumberCell(1 / 3, undefined, false, "fr")).toBe("0,333333333");
    expect(formatGeneralNumber(123456789012, "fr")).toBe("1,23457E+11");
    expect(formatGeneralNumber(42, "fr")).toBe("42");
  });

  it("built-in id 14 is read as the page language's short date (the caller passes the code)", () => {
    expect(formatNumberCell(46053, "m/d/yy", false, "en")).toBe("1/31/26");
    expect(formatNumberCell(46053, "dd/mm/yyyy", false, "fr")).toBe("31/01/2026");
  });

  it("the language defaults to English when omitted", () => {
    expect(formatNumberCell(1234.5, "#,##0.00", false)).toBe("1,234.50");
  });

  it("a hide-in-place code shows nothing whatever the language", () => {
    for (const language of ["en", "fr"] as const) {
      expect(formatNumberCell(5, ";;;", false, language)).toBe("");
      expect(formatNumberCell(0, "0.00;;", false, language)).toBe("");
      expect(formatNumberCell(-5, "0.00;;", false, language)).toBe("");
    }
  });

  it("the French abbreviations are CLDR's (lowercase, points included)", () => {
    const months = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."];
    // First of each month of 2023: serials 44927, 44958, 44986, 45017, 45047, 45078, 45108, 45139, 45170, 45200, 45231, 45261.
    const firsts = [44927, 44958, 44986, 45017, 45047, 45078, 45108, 45139, 45170, 45200, 45231, 45261];
    firsts.forEach((serial, i) => expect(formatNumberCell(serial, "mmm", false, "fr")).toBe(months[i]));
    // Monday 2 January 2023 (44928) … Sunday 8 January 2023.
    const days = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."];
    days.forEach((day, i) => expect(formatNumberCell(44928 + i, "ddd", false, "fr")).toBe(day));
    expect(formatNumberCell(44927, "mmmm", false, "fr")).toBe("janvier");
    expect(formatNumberCell(44928, "dddd", false, "fr")).toBe("lundi");
    expect(formatNumberCell(44927, "mmmmm", false, "fr")).toBe("J");
  });

  it("fractions of a second follow the page's decimal mark", () => {
    expect(formatNumberCell(0.0000605, "ss.00", false, "en")).toBe("05.23");
    expect(formatNumberCell(0.0000605, "ss.00", false, "fr")).toBe("05,23");
  });
});

describe("a locale written in the format governs the WORDS only (§2 rule 1, decided Q2)", () => {
  // The fifteen French language ids of the design, plus the labels.
  const FRENCH_IDS = [
    "040C",
    "080C",
    "0C0C",
    "100C",
    "140C",
    "180C",
    "1C0C",
    "200C",
    "240C",
    "280C",
    "2C0C",
    "300C",
    "340C",
    "380C",
    "3C0C",
  ];
  for (const id of FRENCH_IDS) {
    it(`[$-${id}] gives French words on an English page`, () => {
      expect(formatNumberCell(45000, `[$-${id}]d mmmm yyyy`, false, "en")).toBe("15 mars 2023");
    });
  }
  it("a lowercase or short hex id counts as well: [$-40c], [$€-40C]", () => {
    expect(formatNumberCell(45000, "[$-40c]mmmm", false, "en")).toBe("mars");
    expect(formatNumberCell(45000, "[$€-40C]mmmm", false, "en")).toBe("€mars"); // the symbol shows, the words are French
  });
  it("a French label — fr, fr-FR, fr-CA — gives French words", () => {
    for (const label of ["fr", "fr-FR", "fr-CA"]) {
      expect(formatNumberCell(45000, `[$-${label}]mmmm`, false, "en")).toBe("mars");
    }
  });
  it("any other locale gives English words, even on a French page", () => {
    expect(formatNumberCell(45000, "[$-409]mmmm", false, "fr")).toBe("March");
    expect(formatNumberCell(45000, "[$-407]mmmm", false, "fr")).toBe("March");
    expect(formatNumberCell(45000, "[$-en-US]mmmm", false, "fr")).toBe("March");
    expect(formatNumberCell(45000, "[$-de-DE]mmmm", false, "fr")).toBe("March");
  });
  it("[$-x-sysdate] stays undefined: the system date extension is not ours", () => {
    expect(formatNumberCell(45000, "[$-x-sysdate]dddd, mmmm dd, yyyy", false, "fr")).toBeUndefined();
  });
  // 06/11/2026, reversing the 29/09 refusal: [$-F800] is the LONG DATE of the machine opening the file, which we
  // cannot know; the page language stands in. A built cell, no real specimen in fixtures/ (21 such cells in the 68-workbook corpus).
  it("[$-F800] writes the long date in the PAGE language, never the serial number", () => {
    expect(formatNumberCell(45000, "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", false, "en")).toBe(
      "Wednesday, March 15, 2023",
    );
    expect(formatNumberCell(45000, "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", false, "fr")).toBe("mercredi 15 mars 2023");
    expect(formatNumberCell(45000.5, "[$-F800]dddd, mmmm dd, yyyy", false, "en")).toBe("Wednesday, March 15, 2023");
  });
  // 06/11/2026: a literal space escaped inside the digit placeholders, as Excel's screen shows it. Built cells.
  it("#\\ ##0 writes the space where the code puts it", () => {
    expect(formatNumberCell(1234, "#\\ ##0", false, "en")).toBe("1 234");
    expect(formatNumberCell(1234, "#\\ ##0", false, "fr")).toBe("1 234");
    expect(formatNumberCell(1234567, "#\\ ##0", false, "en")).toBe("1234 567");
    expect(formatNumberCell(-1234, "#\\ ##0", false, "en")).toBe("-1 234");
    expect(formatNumberCell(0, "#\\ ##0", false, "en")).toBe(" 0");
  });
  it("separators keep following the PAGE language, whatever the locale in the code", () => {
    expect(formatNumberCell(1234.5, "[$-409]#,##0.00", false, "fr")).toBe(`1${NBSP}234,50`);
    expect(formatNumberCell(1234.5, "[$-40C]#,##0.00", false, "en")).toBe("1,234.50");
  });
  it("the currency symbol keeps the code's own place", () => {
    expect(formatNumberCell(1234.5, "#,##0.00 [$€-40C]", false, "fr")).toBe(`1${NBSP}234,50 €`);
    expect(formatNumberCell(1234.5, "[$€-40C]#,##0.00", false, "fr")).toBe(`€1${NBSP}234,50`);
  });
});
