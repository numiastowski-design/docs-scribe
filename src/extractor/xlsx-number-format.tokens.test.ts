/**
 * Characterization of `xlsx-number-format.ts` at the seam of its token readers
 * (`tokenize`) and its date/time writer (`writeDateTime`): one row per
 * shape each step decides on. Every expected text was produced by the code
 * BEFORE the 10C2 split (audit de factorisation, lot 10) and is pinned as is,
 * oddities included (« mm:[ss] », « [$ -409]d »): this file guards that the
 * split changed nothing, it does not judge Excel fidelity.
 */
import { describe, expect, it } from "vitest";

import { formatNumberCell } from "./xlsx-number-format";

/** [code, value, page language, 1904 system, expected text (undefined: the raw value is kept)] */
type Row = [string, number, "en" | "fr", boolean, string | undefined];

const ROWS: Row[] = [
  ['0"abc', 45000.7243, "en", false, undefined],
  ['"abc', 45000.7243, "en", false, undefined],
  ["0\\", 45000.7243, "en", false, undefined],
  ["\\", 45000.7243, "en", false, undefined],
  ["[h", 45000.7243, "en", false, undefined],
  ["[hh", 45000.7243, "en", false, undefined],
  ["[s", 45000.7243, "en", false, undefined],
  ["[$-409]", 45000.7243, "en", false, undefined],
  ["0_", 45000.7243, "en", false, undefined],
  ["0*", 45000.7243, "en", false, undefined],
  ["E+00", 45000.7243, "en", false, undefined],
  ["0 E+00", 45000.7243, "en", false, undefined],
  ["d.0", 45000.7243, "en", false, undefined],
  ["h.00", 45000.7243, "en", false, undefined],
  ["0 ,", 45000.7243, "en", false, undefined],
  ["0 .", 45000.7243, "en", false, undefined],
  ["hhh:mm", 45000.7243, "en", false, undefined],
  ["sss", 45000.7243, "en", false, undefined],
  ["h:mm A/", 45000.7243, "en", false, undefined],
  ["h:mm AM/P", 45000.7243, "en", false, undefined],
  ["h a", 45000.7243, "en", false, undefined],
  ["h:mm@", 45000.7243, "en", false, undefined],
  ["[H]:mm", 45000.7243, "en", false, "1080017:23"],
  ["[M]", 45000.7243, "en", false, "64801043"],
  ["[S]", 45000.7243, "en", false, "3888062580"],
  ["[Hh]:mm", 45000.7243, "en", false, "1080017:23"],
  ["[h]:mm", 45000.7243, "en", false, "1080017:23"],
  ["[$-409]mmm", 45000.7243, "en", false, "Mar"],
  ["[$-F800]mmm", 45000.7243, "en", false, "Wednesday, March 15, 2023"],
  ["[$-x-sysdate]mmm", 45000.7243, "en", false, undefined],
  ["[$€]0.00", 45000.7243, "en", false, "€45000.72"],
  ["[$€-409]0.00", 45000.7243, "en", false, "€45000.72"],
  ["[$ -409]d", 45000.7243, "en", false, " 15"],
  ["[Red]d", 45000.7243, "en", false, "15"],
  ["[>5]d", 45000.7243, "en", false, "15"],
  ["[Whatever]d", 45000.7243, "en", false, undefined],
  ["0.0E-00", 12345, "en", false, "1.2E04"],
  ["0.0E-00", 0.00012345, "en", false, "1.2E-04"],
  ["0.0E+00", 12345, "en", false, "1.2E+04"],
  ["0.0E+00", 0.00012345, "en", false, "1.2E-04"],
  ["0.00e+00", 12345, "en", false, "1.23E+04"],
  ["0.00e+00", 0.00012345, "en", false, "1.23E-04"],
  ["0.0e-0", 12345, "en", false, "1.2E4"],
  ["0.0e-0", 0.00012345, "en", false, "1.2E-4"],
  ["0E+0", 12345, "en", false, "1E+4"],
  ["0E+0", 0.00012345, "en", false, "1E-4"],
  ["E+0 0", 12345, "en", false, undefined],
  ["E+0 0", 0.00012345, "en", false, undefined],
  ["YYYY-MM-DD", 45000.7243, "en", false, "2023-03-15"],
  ["HH:MM:SS", 45000.7243, "en", false, "17:23:00"],
  ["DDDD", 45000.7243, "en", false, "Wednesday"],
  ["DDD", 45000.7243, "en", false, "Wed"],
  ["D/M/YY", 45000.7243, "en", false, "15/3/23"],
  ["H:M", 45000.7243, "en", false, "17:23"],
  ["h:m", 45000.7243, "en", false, "17:23"],
  ["m:ss", 45000.7243, "en", false, "23:00"],
  ["mm:[ss]", 45000.7243, "en", false, "23:3888062580"],
  ["m[ss]", 45000.7243, "en", false, "233888062580"],
  ["h[$-409]m", 45000.7243, "en", false, "1723"],
  ["m[$-409]s", 45000.7243, "en", false, "230"],
  ['h"x"m', 45000.7243, "en", false, "17x23"],
  ["h m", 45000.7243, "en", false, "17 23"],
  ["m m", 45000.7243, "en", false, "3 3"],
  ["mm h", 45000.7243, "en", false, "03 17"],
  ["yyyy mm", 45000.7243, "en", false, "2023 03"],
  ["h:mm AM/PM", 45000.7243, "en", false, "5:23 PM"],
  ["h:mm am/pm", 45000.7243, "en", false, "5:23 PM"],
  ["h a/p", 45000.7243, "en", false, "5 p"],
  ["h A/P", 45000.7243, "en", false, "5 P"],
  ["h:mm a/p", 45000.7243, "en", false, "5:23 p"],
  ["hh:mm:ss a/p", 45000.7243, "en", false, "05:23:00 p"],
  ["h:mm A/P", 45000.7243, "en", false, "5:23 P"],
  ["[hh]:mm", 45000.7243, "en", false, "1080017:23"],
  ["[mm]:ss", 45000.7243, "en", false, "64801043:00"],
  ["[ss]", 45000.7243, "en", false, "3888062580"],
  ["[s]", 45000.7243, "en", false, "3888062580"],
  ["mm:ss.00", 45000.7243, "en", false, "22:59.52"],
  ["ss.000", 45000.7243, "en", false, "59.520"],
  ["ss.", 45000.7243, "en", false, "00."],
  ["ss.x", 45000.7243, "en", false, undefined],
  ["[ss].0", 45000.7243, "en", false, "3888062579.5"],
  ["m.00", 45000.7243, "en", false, undefined],
  ["mm:ss.0000", 45000.7243, "en", false, undefined],
  ["h:mm AM/PM", 0.25, "en", false, "6:00 AM"],
  ["h:mm AM/PM", 0.5, "en", false, "12:00 PM"],
  ["h:mm AM/PM", 0.999, "en", false, "11:58 PM"],
  ["h a/p", 0.25, "en", false, "6 a"],
  ["h a/p", 0.5, "en", false, "12 p"],
  ["h a/p", 0.999, "en", false, "11 p"],
  ["d/m", 45000.7243, "en", false, "15/3"],
  ["d m", 45000.7243, "en", false, "15 3"],
  ["d-m", 45000.7243, "en", false, "15-3"],
  ["d:m", 45000.7243, "en", false, "15:3"],
  ["d$m", 45000.7243, "en", false, "15$3"],
  ["d+m", 45000.7243, "en", false, "15+3"],
  ["d(m", 45000.7243, "en", false, "15(3"],
  ["d)m", 45000.7243, "en", false, "15)3"],
  ["d, mmm", 45000.7243, "en", false, "15, Mar"],
  ["d. mmmm", 45000.7243, "en", false, "15. March"],
  ["d,m", 45000.7243, "en", false, "15,3"],
  ["d.m", 45000.7243, "en", false, "15.3"],
  ["(d)", 45000.7243, "en", false, "(15)"],
  ["$d", 45000.7243, "en", false, "$15"],
  ["+d", 45000.7243, "en", false, "+15"],
  ["d - m", 45000.7243, "en", false, "15 - 3"],
  ["0%%", 0.5, "en", false, "5000%%"],
  ["0%", 0.5, "en", false, "50%"],
  ["%0", 0.5, "en", false, "%50"],
  ["0 %", 0.5, "en", false, "50 %"],
  ["#0?", 0.5, "en", false, "01"],
  ["0.0#", 0.5, "en", false, "0.5"],
  ["##0", 0.5, "en", false, "1"],
  ["#,##0.00", 0.5, "en", false, "0.50"],
  ["?0", 0.5, "en", false, " 1"],
  ["0?", 0.5, "en", false, "01"],
  ["0_)", 5, "en", false, "5"],
  ["0*x", 5, "en", false, "5"],
  ["_(0", 5, "en", false, "5"],
  ["0 _x*y", 5, "en", false, "5 "],
  ["0*-", 5, "en", false, "5"],
  ["dddd d mmmm yyyy", 45000.7243, "fr", false, "mercredi 15 mars 2023"],
  ["[$-409]dddd d mmmm yyyy", 45000.7243, "fr", false, "Wednesday 15 March 2023"],
  ["[$-40C]dddd d mmmm yyyy", 45000.7243, "en", false, "mercredi 15 mars 2023"],
  ["[$fr-FR]dddd", 45000.7243, "en", false, "frmercredi"],
  ["d/m/yyyy", 1, "en", true, "2/1/1904"],
  ["h:mm", 1.5, "en", true, "12:00"],
  ["d/m/yyyy h:mm", 1.75, "en", true, "2/1/1904 18:00"],
  ["dd/mm/yyyy hh:mm:ss", 36525.5, "en", true, "01/01/2004 12:00:00"],
  ["dddd h:mm", 100.25, "en", true, "Sunday 6:00"],
  ["d/m/yyyy h:mm", 1.75, "en", false, "1/1/1900 18:00"],
  // Past 9999-12-31 Excel shows #####: raw value kept (decision n° 40; the year no longer wraps).
  ["yyyy", 2958466, "en", false, undefined],
  ["yyyy", 2958900, "en", false, undefined],
  ["yy", 2958900, "en", false, undefined],
  ["[h]", 1.7, "en", false, "40"],
  ["d/m/yyyy h:mm", 45000.99999999, "en", false, "16/3/2023 0:00"],
  ["d/m/yyyy h:mm", 45000.99999999, "en", true, "17/3/2027 0:00"],
  ["d/m/yyyy", 45000.99999999, "en", false, "15/3/2023"],
  ["d/m/yyyy", 0.5, "en", false, undefined],
  ["d/m/yyyy", 0, "en", false, undefined],
  ["d h:mm", 0.5, "en", false, undefined],
  ["d/m/yyyy", -0.5, "en", false, undefined],
  ["[$-409]d [$-40C]mmmm", 45000.7243, "en", false, "15 mars"],
  ["[$-40C]d [$-409]mmmm", 45000.7243, "fr", false, "15 March"],
  // A serial past the JS date range: no display, raw value kept (decision n° 40).
  ["d/m/yyyy", 3000000000, "en", false, undefined],
  ["[mm]", 0.0108, "en", false, "15"],
  ["mmmmm", 45000.7243, "en", false, "M"],
  ["mmmmm", 1, "en", false, "J"],
  ["mmmmm", 59, "en", false, "F"],
  ["mmmmm", 60, "en", false, "F"],
  ["mmmmm", 61, "en", false, "M"],
  ["mmmm", 45000.7243, "en", false, "March"],
  ["mmmm", 1, "en", false, "January"],
  ["mmmm", 59, "en", false, "February"],
  ["mmmm", 60, "en", false, "February"],
  ["mmmm", 61, "en", false, "March"],
  ["mmm", 45000.7243, "en", false, "Mar"],
  ["mmm", 1, "en", false, "Jan"],
  ["mmm", 59, "en", false, "Feb"],
  ["mmm", 60, "en", false, "Feb"],
  ["mmm", 61, "en", false, "Mar"],
  ["mm", 45000.7243, "en", false, "03"],
  ["mm", 1, "en", false, "01"],
  ["mm", 59, "en", false, "02"],
  ["mm", 60, "en", false, "02"],
  ["mm", 61, "en", false, "03"],
  ["m", 45000.7243, "en", false, "3"],
  ["m", 1, "en", false, "1"],
  ["m", 59, "en", false, "2"],
  ["m", 60, "en", false, "2"],
  ["m", 61, "en", false, "3"],
  ["yyyyy", 45000.7243, "en", false, "2023"],
  ["yyyyy", 1, "en", false, "1900"],
  ["yyyyy", 59, "en", false, "1900"],
  ["yyyyy", 60, "en", false, "1900"],
  ["yyyyy", 61, "en", false, "1900"],
  ["yyy", 45000.7243, "en", false, "2023"],
  ["yyy", 1, "en", false, "1900"],
  ["yyy", 59, "en", false, "1900"],
  ["yyy", 60, "en", false, "1900"],
  ["yyy", 61, "en", false, "1900"],
  ["yy", 45000.7243, "en", false, "23"],
  ["yy", 1, "en", false, "00"],
  ["yy", 59, "en", false, "00"],
  ["yy", 60, "en", false, "00"],
  ["yy", 61, "en", false, "00"],
  ["y", 45000.7243, "en", false, "23"],
  ["y", 1, "en", false, "00"],
  ["y", 59, "en", false, "00"],
  ["y", 60, "en", false, "00"],
  ["y", 61, "en", false, "00"],
  ["dddd", 45000.7243, "en", false, "Wednesday"],
  ["dddd", 1, "en", false, "Sunday"],
  ["dddd", 59, "en", false, "Tuesday"],
  ["dddd", 60, "en", false, "Wednesday"],
  ["dddd", 61, "en", false, "Thursday"],
  ["ddd", 45000.7243, "en", false, "Wed"],
  ["ddd", 1, "en", false, "Sun"],
  ["ddd", 59, "en", false, "Tue"],
  ["ddd", 60, "en", false, "Wed"],
  ["ddd", 61, "en", false, "Thu"],
  ["dd", 45000.7243, "en", false, "15"],
  ["dd", 1, "en", false, "01"],
  ["dd", 59, "en", false, "28"],
  ["dd", 60, "en", false, "28"],
  ["dd", 61, "en", false, "01"],
  ["d", 45000.7243, "en", false, "15"],
  ["d", 1, "en", false, "1"],
  ["d", 59, "en", false, "28"],
  ["d", 60, "en", false, "28"],
  ["d", 61, "en", false, "1"],
];

describe("number-format tokens and date/time writer, pinned shapes", () => {
  it.each(ROWS)("%j at %d (%s, 1904: %s)", (code, value, language, date1904, expected) => {
    expect(formatNumberCell(value, code, date1904, language)).toBe(expected);
  });

  // Decision n° 40 (15/10/2026): a serial past the JS date range is a date with
  // no display, like one before day 1: the raw value is kept, never NaN, never a throw.
  it.each(["d-mmm-yy", "mmmmm", "yyyy-mm-dd", "dddd", "[h]:mm"])(
    "keeps the raw value of 1e21, -1e21, Infinity and NaN under %j",
    (code) => {
      for (const value of [1e21, -1e21, Infinity, -Infinity, NaN]) {
        expect(formatNumberCell(value, code, false, "en")).toBeUndefined();
      }
    },
  );

  // Excel's last displayable date is 9999-12-31 (serial 2958465); above it Excel shows #####.
  it("writes 9999-12-31 at serial 2958465 and keeps the raw value above it", () => {
    expect(formatNumberCell(2958465, "yyyy-mm-dd", false, "en")).toBe("9999-12-31");
    expect(formatNumberCell(2958466, "yyyy-mm-dd", false, "en")).toBeUndefined();
    expect(formatNumberCell(100_000_000, "yyyy-mm-dd", false, "en")).toBeUndefined();
    expect(formatNumberCell(2958466, "mmmmm", false, "en")).toBeUndefined();
  });

  it("holds the same last date in the 1904 system (offset 1462)", () => {
    expect(formatNumberCell(2958465 - 1462, "yyyy-mm-dd", true, "en")).toBe("9999-12-31");
    expect(formatNumberCell(2958466 - 1462, "yyyy-mm-dd", true, "en")).toBeUndefined();
  });

  it("keeps the raw value when the time rounds up into the day after the last date", () => {
    expect(formatNumberCell(2958465.99999999, "yyyy-mm-dd hh:mm:ss", false, "en")).toBeUndefined();
    expect(formatNumberCell(2958465.5, "yyyy-mm-dd hh:mm:ss", false, "en")).toBe("9999-12-31 12:00:00");
  });

  it("pins enough distinct shapes to guard each step", () => {
    expect(ROWS.length).toBeGreaterThan(150);
    expect(new Set(ROWS.map(([, , , , out]) => out === undefined)).size).toBe(2);
  });
});
