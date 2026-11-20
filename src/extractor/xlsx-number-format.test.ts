/**
 * OUR number formatter (`xlsx-number-format.ts`), the only one since SheetJS
 * was removed (29/09/2026). `CASES` are (raw value, code) pairs harvested
 * from the repository's workbooks with the text SheetJS displayed then
 * (28/09/2026), kept byte-identical. The later blocks (HIGH review of 29/09,
 * General, engineering, fractional seconds, accounting, digit placeholders)
 * are pinned on Excel's own documentation (ISO 29500, Microsoft Learn and
 * support pages, named in each `describe`), not on SheetJS. The US shape is
 * TEMPORARY (decision (c)), but it is the shipped behavior.
 */
import { describe, expect, it } from "vitest";
import { formatNumberCell, isHiddenTextFormatCode } from "./xlsx-number-format";

const CASES: [code: string, value: number, expected: string][] = [
  // Dates (serial 1900). Harvested: fixture 11 B2, a real expense sheet (B9), fixture 11 B6 (builtin 14 → m/d/yy).
  ["dd/mm/yyyy", 46053, "31/01/2026"],
  ["d/m/yyyy", 44852, "18/10/2022"],
  ["m/d/yy", 46173, "5/31/26"],
  // Percentages. Harvested: a real comparison sheet (B3), fixture 11 D2.
  ["0%", 2, "200%"],
  ["0.0%", 0.215, "21.5%"],
  // Currency and grouping. Harvested: fixture 11 C2, the comparison sheet (B7/B24), the expense sheet (D3/D4, G9).
  ['#,##0.00 "€"', 18400.5, "18,400.50 €"],
  ['#,##0\\ "€";[Red]\\-#,##0\\ "€"', 70, "70 €"],
  ['#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"', 322.5, "322.50 €"],
  ["#,##0.0", 2608, "2,608.0"],
  ["#,##0.000", 0.603, "0.603"],
  ["0", 50, "50"],
];

describe("formatNumberCell — the harvested (value, code) pairs render byte-identical (étape 5)", () => {
  for (const [code, value, expected] of CASES) {
    it(`${JSON.stringify(code)} on ${value} → "${expected}"`, () => {
      expect(formatNumberCell(value, code, false)).toBe(expected);
    });
  }

  it("a negative value takes the SECOND section, with its own literal minus", () => {
    expect(formatNumberCell(-206.01, '#,##0.00\\ "€";\\-#,##0.00\\ "€"', false)).toBe("-206.01 €");
  });

  it("date1904: the same date sits 1 462 serials lower", () => {
    expect(formatNumberCell(46053 - 1462, "dd/mm/yyyy", true)).toBe("31/01/2026");
  });

  // Times and scientific were the examples here until the revue du 29/09
  // (option c): we write them now — the still-unknown shapes take over.
  it("an unknown code answers undefined — the caller falls back to the raw value, never invents", () => {
    expect(formatNumberCell(3.14, "[DBNum1]0", false)).toBeUndefined();
    expect(formatNumberCell(45000, "[$-x-sysdate]dddd, mmmm dd, yyyy", false)).toBeUndefined(); // accounting was here: written since 29/09 evening
  });
});

describe("formatNumberCell — hidden values, scaling commas, datetimes (constats 3, 4, 7)", () => {
  it("🔴 « ;;; » (the classic hide-in-place format) publishes EMPTY, never the raw value", () => {
    expect(formatNumberCell(123456, ";;;", false)).toBe("");
    expect(formatNumberCell(-5, "0.00;;", false)).toBe(""); // negative hits the empty second section
  });

  it("🔴 a trailing scaling comma (« #,##0, » = thousands) DIVIDES — never a figure 1000× off", () => {
    expect(formatNumberCell(1234567, "#,##0,", false)).toBe("1,235");
    expect(formatNumberCell(1234567, "0.00,,", false)).toBe("1.23");
    // Never the undivided number.
    expect(formatNumberCell(1234567, "#,##0,", false)).not.toBe("1,234,567");
  });

  it("🔴 a datetime under a pure DATE format renders the date — never the raw serial", () => {
    expect(formatNumberCell(45123.75, "dd/mm/yyyy", false)).toBe("16/07/2023");
  });
});

describe("isHiddenTextFormatCode — the TEXT half of the hide-in-place formats (constat 3)", () => {
  it("🔴 a 4-section code with an empty text section hides text; anything else does not", () => {
    expect(isHiddenTextFormatCode(";;;")).toBe(true);
    expect(isHiddenTextFormatCode('0;-0;0;"x"')).toBe(false);
    expect(isHiddenTextFormatCode("@")).toBe(false);
    expect(isHiddenTextFormatCode(undefined)).toBe(false);
    expect(isHiddenTextFormatCode("0.00")).toBe(false);
  });
});

/** Excel's « General », not `String(v)` (0.30000000000000004, 1e+21): at most 11 characters, trailing zeros removed,
 *  scientific notation for large and small magnitudes. */
describe("formatNumberCell — Excel « General » (revue 29/09, constat 1)", () => {
  it("🔴 1/3 and 0.1+0.2 read as Excel shows them, never the binary tail", () => {
    expect(formatNumberCell(1 / 3, "General", false)).toBe("0.333333333");
    expect(formatNumberCell(0.1 + 0.2, "General", false)).toBe("0.3");
    expect(formatNumberCell(-1 / 3, "General", false)).toBe("-0.333333333");
  });

  it("🔴 very large and very small magnitudes switch to scientific notation", () => {
    expect(formatNumberCell(1e21, "General", false)).toBe("1E+21");
    expect(formatNumberCell(123456789012, "General", false)).toBe("1.23457E+11");
    expect(formatNumberCell(1e-10, "General", false)).toBe("1E-10");
  });

  it("integers up to 11 characters are unchanged; an absent code IS General", () => {
    expect(formatNumberCell(42, "General", false)).toBe("42");
    expect(formatNumberCell(-7, undefined, false)).toBe("-7");
    expect(formatNumberCell(12345678901, "General", false)).toBe("12345678901");
    expect(formatNumberCell(0, "General", false)).toBe("0");
    expect(formatNumberCell(2608.5, "@", false)).toBe("2608.5"); // text format on a number: General
  });
});

describe("formatNumberCell — zero section, decimal rounding, '#' on zero (revue 29/09, constat 7)", () => {
  it("🔴 a zero value takes the THIRD section", () => {
    expect(formatNumberCell(0, '0.00;-0.00;"zéro"', false)).toBe("zéro");
    expect(formatNumberCell(0, "0.00;-0.00;\-", false)).toBe("-");
    expect(formatNumberCell(5, '0.00;-0.00;"zéro"', false)).toBe("5.00");
  });

  it("🔴 rounding is half away from zero on the DECIMAL value, as Excel", () => {
    expect(formatNumberCell(1.005, "0.00", false)).toBe("1.01");
    expect(formatNumberCell(2.675, "#,##0.00", false)).toBe("2.68");
    expect(formatNumberCell(-1.005, "0.00", false)).toBe("-1.01");
    expect(formatNumberCell(0.5, "0", false)).toBe("1");
    expect(formatNumberCell(1234.5, "#,##0", false)).toBe("1,235");
  });

  it("🔴 a '#'-only integer mask renders nothing for a zero integer part", () => {
    expect(formatNumberCell(0, "#", false)).toBe("");
    expect(formatNumberCell(0, "#,###", false)).toBe("");
    expect(formatNumberCell(0, "#,##0", false)).toBe("0");
  });
});

/** Month and day names are English: the US format is temporary. A code still unknown gives `undefined`, the raw
 *  value. */
describe("formatNumberCell — the missing families, written by us (revue 29/09, constats 2 et 6, option c)", () => {
  it("🔴 dates with month and day NAMES (English)", () => {
    expect(formatNumberCell(45000, "d-mmm-yy", false)).toBe("15-Mar-23");
    expect(formatNumberCell(45000, "dddd, mmmm d, yyyy", false)).toBe("Wednesday, March 15, 2023");
    expect(formatNumberCell(45000, "ddd dd mmm", false)).toBe("Wed 15 Mar");
    expect(formatNumberCell(45000, "mmmmm", false)).toBe("M");
  });

  it("🔴 times: h, hh, minutes after h or before s, ss, AM/PM", () => {
    expect(formatNumberCell(0.75, "hh:mm", false)).toBe("18:00");
    expect(formatNumberCell(0.75, "h:mm AM/PM", false)).toBe("6:00 PM");
    expect(formatNumberCell(0.25, "h:mm A/P", false)).toBe("6:00 A");
    expect(formatNumberCell(3661 / 86400, "hh:mm:ss", false)).toBe("01:01:01");
    expect(formatNumberCell(3661 / 86400, "mm:ss", false)).toBe("01:01");
    expect(formatNumberCell(45123.75, "dd/mm/yyyy hh:mm", false)).toBe("16/07/2023 18:00");
  });

  it("🔴 elapsed time: [h], [mm], [ss]", () => {
    expect(formatNumberCell(1.5, "[h]:mm", false)).toBe("36:00");
    expect(formatNumberCell(3661 / 86400, "[mm]:ss", false)).toBe("61:01");
    expect(formatNumberCell(3661 / 86400, "[ss]", false)).toBe("3661");
  });

  it("🔴 scientific: 0.00E+00", () => {
    expect(formatNumberCell(12345, "0.00E+00", false)).toBe("1.23E+04");
    expect(formatNumberCell(0.000123, "0.00E+00", false)).toBe("1.23E-04");
    expect(formatNumberCell(-12345, "0.00E+00", false)).toBe("-1.23E+04");
    expect(formatNumberCell(0, "0.00E+00", false)).toBe("0.00E+00");
    expect(formatNumberCell(99999, "0.0E+0", false)).toBe("1.0E+5");
  });

  it("🔴 fractions: # ?/?, # ??/??, a fixed denominator", () => {
    expect(formatNumberCell(1.5, "# ?/?", false)).toBe("1 1/2");
    expect(formatNumberCell(3.14159, "# ??/??", false)).toBe("3  1/7 ");
    expect(formatNumberCell(0.75, "?/?", false)).toBe("3/4");
    expect(formatNumberCell(2.5, "# ?/4", false)).toBe("2 2/4");
  });

  it("🔴 currency/locale brackets: the symbol shows, the locale id does not", () => {
    expect(formatNumberCell(1234.5, "[$€-40C]#,##0.00", false)).toBe("€1,234.50");
    expect(formatNumberCell(1234.5, "#,##0.00 [$€-40C]", false)).toBe("1,234.50 €");
    expect(formatNumberCell(1.5, "[$USD] 0.00", false)).toBe("USD 1.50");
    expect(formatNumberCell(46053, "[$-409]d/m/yyyy", false)).toBe("31/1/2026");
  });

  it("🔴 conditions select the section", () => {
    expect(formatNumberCell(1500, "[>=1000]#,##0;0.00", false)).toBe("1,500");
    expect(formatNumberCell(5, "[>=1000]#,##0;0.00", false)).toBe("5.00");
    expect(formatNumberCell(-1, '[<0]"neg";[>=100]"big";0', false)).toBe("neg");
    expect(formatNumberCell(150, '[<0]"neg";[>=100]"big";0', false)).toBe("big");
    expect(formatNumberCell(5, '[<0]"neg";[>=100]"big";0', false)).toBe("5");
  });

  it("colour brackets stay styling, never content", () => {
    expect(formatNumberCell(1.5, "[Red]0.00", false)).toBe("1.50");
    expect(formatNumberCell(1.5, "[Color10]0.00", false)).toBe("1.50");
  });

  it("🔴 any OTHER bracket, and anything still unknown, answers undefined — the raw value, never a guess", () => {
    expect(formatNumberCell(1.5, "[DBNum1]0", false)).toBeUndefined();
    expect(formatNumberCell(-0.5, "hh:mm", false)).toBeUndefined(); // a negative time has no display
  });
});

/** Sources: ISO/IEC 29500-1 §18.8.30 numFmt; Microsoft Learn « Format Cells dialog box in Excel » (table of symbols);
 *  Microsoft support « Number format codes », « Review guidelines for customizing a number format », « How to
 *  correct a ##### error ». Each expectation cites its source. */
describe("formatNumberCell — General at the extremes (ISO 29500 §18.8.30 + support « 12 or more digits »)", () => {
  // §18.8.30 : « max overall length for cell display is 11, not including
  // negative sign, but includes leading zeros and decimal separator » ;
  // exponential only when the float form has no room. 0.00001 FITS (7).
  it("small magnitudes stay fixed while they fit in 11 characters", () => {
    expect(formatNumberCell(0.00001, "General", false)).toBe("0.00001");
    expect(formatNumberCell(0.000001, "General", false)).toBe("0.000001");
    expect(formatNumberCell(-0.000001, "General", false)).toBe("-0.000001");
  });

  it("past 11 characters: exponential (1E-10, 1.23457E-05); 12 digits and more: exponential", () => {
    expect(formatNumberCell(1e-10, "General", false)).toBe("1E-10");
    expect(formatNumberCell(0.000012345678, "General", false)).toBe("1.23457E-05");
    expect(formatNumberCell(12345678901, "General", false)).toBe("12345678901");
    expect(formatNumberCell(123456789012, "General", false)).toBe("1.23457E+11");
    expect(formatNumberCell(1e21, "General", false)).toBe("1E+21");
  });
});

describe("formatNumberCell — engineering notation (Learn: #0.0E+0 on 12,200,000 → 12.2E+6; ISO builtin 48 ##0.0E+0)", () => {
  it("🔴 the exponent is a multiple of the integer placeholder count", () => {
    expect(formatNumberCell(12_200_000, "#0.0E+0", false)).toBe("12.2E+6");
    expect(formatNumberCell(12345, "##0.0E+0", false)).toBe("12.3E+3");
    expect(formatNumberCell(123456, "##0.0E+0", false)).toBe("123.5E+3");
    expect(formatNumberCell(1234, "##0.0E+0", false)).toBe("1.2E+3");
    expect(formatNumberCell(0.0001234, "##0.0E+0", false)).toBe("123.4E-6");
    expect(formatNumberCell(-12345, "##0.0E+0", false)).toBe("-12.3E+3");
    expect(formatNumberCell(999_960, "##0.0E+0", false)).toBe("1.0E+6");
    expect(formatNumberCell(0, "##0.0E+0", false)).toBe("0.0E+0");
  });
});

describe("formatNumberCell — fractional seconds (support « Number format codes »: [ss].00 → 3735.80, h:mm:ss.00 → 4:36:03.75 PM)", () => {
  it("🔴 ss.00, [ss].000, mm:ss.0", () => {
    expect(formatNumberCell(3735.8 / 86400, "[ss].00", false)).toBe("3735.80");
    expect(formatNumberCell((16 * 3600 + 36 * 60 + 3.75) / 86400, "h:mm:ss.00 AM/PM", false)).toBe("4:36:03.75 PM");
    expect(formatNumberCell((16 * 3600 + 36 * 60 + 3.75) / 86400, "hh:mm:ss.00", false)).toBe("16:36:03.75");
    expect(formatNumberCell(3.1416 / 86400, "[ss].000", false)).toBe("03.142"); // [ss] pads to its width
    expect(formatNumberCell(62.25 / 86400, "mm:ss.0", false)).toBe("01:02.3");
    expect(formatNumberCell(59.996 / 86400, "mm:ss.00", false)).toBe("01:00.00");
  });

  it("more than three decimals of a second is not a code Excel writes: raw value", () => {
    expect(formatNumberCell(1 / 86400, "ss.0000", false)).toBeUndefined();
  });
});

describe("formatNumberCell — a negative time (support « How to correct a ##### error »: 1900 system → #####)", () => {
  it("Excel shows no value: undefined, the raw value is published", () => {
    expect(formatNumberCell(-0.5, "hh:mm", false)).toBeUndefined();
    expect(formatNumberCell(-1.5, "[h]:mm", false)).toBeUndefined();
    expect(formatNumberCell(-0.5, "[ss].00", false)).toBeUndefined();
  });
});

/** Accounting: `_x` is a space as wide as x (Learn: « Skip the width of the next character »), published as nothing;
 *  `*x` fills the column (Learn: « Repeat the next character … to fill the column »), nothing either; `?` is a space
 *  for a non-significant digit (Learn); `$ ( )` show without quotes (Learn: « $-+/():space Displays the symbol »).
 *  Thousands separator: the US comma, temporary. */
describe("formatNumberCell — accounting formats (FR and US)", () => {
  const FR = '_-* #,##0 "€"_-;-* #,##0 "€"_-;_-* "-" "€"_-;_-@_-';
  const US = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

  it("🔴 French accounting: positive, negative, zero", () => {
    expect(formatNumberCell(1234, FR, false)).toBe("1,234 €");
    expect(formatNumberCell(-1234, FR, false)).toBe("-1,234 €");
    expect(formatNumberCell(0, FR, false)).toBe("- €");
  });

  it("🔴 US accounting: positive, negative in parentheses, zero dash with ?? alignment spaces", () => {
    expect(formatNumberCell(1234.5, US, false)).toBe("$1,234.50");
    expect(formatNumberCell(-1234.5, US, false)).toBe("$(1,234.50)");
    expect(formatNumberCell(0, US, false)).toBe("$-  ");
  });

  it("the text section does not hide a text cell (xlsx.ts publishes its text as-is)", () => {
    expect(isHiddenTextFormatCode(FR)).toBe(false);
    expect(isHiddenTextFormatCode(US)).toBe(false);
  });

  it("🔴 '?' pads insignificant digits with a space (Learn: 0.0? aligns 8.9 and 88.99)", () => {
    expect(formatNumberCell(8.9, "0.0?", false)).toBe("8.9 ");
    expect(formatNumberCell(88.99, "0.0?", false)).toBe("88.99");
    expect(formatNumberCell(5, "??0", false)).toBe("  5");
  });
});

describe("formatNumberCell — digit placeholders pad as Learn's table says", () => {
  it("🔴 '0' pads with zeros ((000) on 12 → (012)); '#' drops insignificant decimals (#.## on 8.9 → 8.9)", () => {
    expect(formatNumberCell(12, "(000)", false)).toBe("(012)");
    expect(formatNumberCell(8.9, "#.##", false)).toBe("8.9");
    expect(formatNumberCell(8.9, "#.00", false)).toBe("8.90");
  });
});

describe("formatNumberCell — a trailing « @ » text section is not a number section (review finding 5)", () => {
  // Microsoft, « How to control and understand settings in the Format Cells
  // dialog box in Excel »: up to three sections for numbers, a FOURTH for text;
  // « @ » is the text placeholder.
  it("two sections, the second is @: a negative uses the first section", () => {
    expect(formatNumberCell(-1.5, "0.00;@", false)).toBe("-1.50");
    expect(formatNumberCell(1.5, "0.00;@", false)).toBe("1.50");
  });
  it("an EMPTY number section still hides when the original code had a text section", () => {
    expect(formatNumberCell(5, ";@", false)).toBe("");
    expect(formatNumberCell(5, ";;@", false)).toBe("");
    expect(formatNumberCell(-1.5, "0.00;@", false)).toBe("-1.50");
  });
  it("three sections, the third is @: zero uses the first section", () => {
    expect(formatNumberCell(0, "#,##0.00;(#,##0.00);@", false)).toBe("0.00");
    expect(formatNumberCell(-1234.5, "#,##0.00;(#,##0.00);@", false)).toBe("(1,234.50)");
  });
});

describe("formatNumberCell — the parsed code is memoised per code (review finding 13)", () => {
  // The memo is a cost, not a result: what repeated calls on ONE code must keep is the same answer for the same value, in either section.
  // The count of tokenizations the former test read is gone with its instrumentation (no behaviour of the output can tell a memo from a re-parse).
  it("the same code answers the same on every call, positive then negative section, in any order", () => {
    const code = '#,##0.00\ "€";\-#,##0.00\ "€"';
    const first = [1234.5, -1234.5];
    const expected = first.map((v) => formatNumberCell(v, code, false));
    expect(expected.every((text) => text !== undefined)).toBe(true);
    expect(expected[0]).not.toBe(expected[1]);
    for (let i = 0; i < 50; i++) {
      expect(formatNumberCell(first[i % 2]!, code, false)).toBe(expected[i % 2]);
      expect(formatNumberCell(first[(i + 1) % 2]!, code, false)).toBe(expected[(i + 1) % 2]);
    }
  });
});

describe("formatNumberCell — pathological decimals stay bounded (review finding 19)", () => {
  // Microsoft, « Display more digits in trendline equation coefficients »:
  // « Decimal places setting to 30 or less ».
  const fast = (run: () => unknown): void => {
    const start = performance.now();
    run();
    expect(performance.now() - start).toBeLessThan(1500);
  };
  it("a 5M-decimal mask is refused fast", () => {
    fast(() => expect(formatNumberCell(1234.5, `0.${"0".repeat(5_000_000)}`, false)).toBeUndefined());
  });
  it("a 5M-zero integer mask is refused fast (code longer than MAX_FORMAT_CODE_CHARS)", () => {
    fast(() => expect(formatNumberCell(1234.5, "0".repeat(5_000_000), false)).toBeUndefined());
  });
  it("a 5M-wide elapsed bracket is refused fast", () => {
    fast(() => expect(formatNumberCell(1, `[${"h".repeat(5_000_000)}]`, false)).toBeUndefined());
  });
  it("a code of exactly 1024 characters is still read, 1025 is refused", () => {
    expect(formatNumberCell(5, `0${" ".repeat(1023)}`, false)).toBe(`5${" ".repeat(1023)}`);
    expect(formatNumberCell(5, `0${" ".repeat(1024)}`, false)).toBeUndefined();
  });
  it("31 decimals are refused, 30 are written", () => {
    expect(formatNumberCell(1.5, `0.${"0".repeat(31)}`, false)).toBeUndefined();
    expect(formatNumberCell(1.5, `0.${"0".repeat(30)}`, false)).toBe(`1.5${"0".repeat(29)}`);
  });
  it("grouping a 301-digit integer is linear and exact", () => {
    fast(() => expect(formatNumberCell(1e300, "#,##0", false)).toMatch(/^1(,000){100}$/));
  });
});

/** Scaling commas: a comma after the last digit placeholder divides by 1,000, once per comma (Microsoft, « Number
 *  format codes »: « Display numbers in thousands »). Rendered as ordinary grouping it would publish numbers 1,000
 *  times too large. */
describe("formatNumberCell — la virgule d'échelle divise par mille (30/10/2026)", () => {
  it("« #,##0, » affiche les milliers, groupés", () => {
    expect(formatNumberCell(1234567, "#,##0,", false)).toBe("1,235");
  });

  it("« #,, » affiche les millions", () => {
    expect(formatNumberCell(1234567890, "#,,", false)).toBe("1235");
  });

  it("une virgule SUIVIE d'un emplacement de chiffre reste un groupement, elle ne divise rien", () => {
    expect(formatNumberCell(1234567, "#,##0", false)).toBe("1,234,567");
  });

  it("le signe et l'arrondi suivent l'échelle, pas la valeur d'origine", () => {
    expect(formatNumberCell(-1500, "#,##0,", false)).toBe("-2");
    expect(formatNumberCell(499, "#,##0,", false)).toBe("0");
  });
});

/** A locale prefix before « General » (`[$-1010409]General`): the bracket names a reading locale, which the formatter
 *  does not use, and `General` is the default format. */
describe("formatNumberCell — « General » précédé d'un préfixe de langue (30/10/2026)", () => {
  it("se rend comme General", () => {
    expect(formatNumberCell(1234.5, "[$-1010409]General", false)).toBe(formatNumberCell(1234.5, "General", false));
  });

  it("et en français aussi, avec sa virgule décimale", () => {
    expect(formatNumberCell(1234.5, "[$-1010409]General", false, "fr")).toBe(
      formatNumberCell(1234.5, "General", false, "fr"),
    );
  });
});

/** A non-finite value is not written. A format code is text the document chooses: « % » repeated two hundred times
 *  overflows the multiplication to `Infinity`, and decimal rounding would turn it into the string « NaN ». */
describe("formatNumberCell — rien d'infini ne devient du texte (30/10/2026)", () => {
  it("un pourcentage qui déborde répond undefined, jamais « NaN »", () => {
    const out = formatNumberCell(1e300, `${"%".repeat(200)}0`, false);
    expect(out).toBeUndefined();
  });

  it("et la notation scientifique aussi", () => {
    expect(formatNumberCell(1e300, `${"%".repeat(200)}0.00E+00`, false)).toBeUndefined();
  });
});
