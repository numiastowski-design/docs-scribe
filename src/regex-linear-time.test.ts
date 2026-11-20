/**
 * Hostile inputs for regexes that were quadratic (sonarjs/super-linear-regex triage, 14/10/2026). Each case asserts
 * the RESULT and a loose time bound; the time measured before the fix is given per case.
 * The cases of this package's own functions: until 19/11/2026 they sat with the publisher's app's cases, in one file.
 * Run: npx vitest run --project=pdf-serial packages/docs-scribe/src/regex-linear-time.test.ts
 */
import { describe, expect, it } from "vitest";

import { partNumberOf } from "./extractor/smartart";
import { formatNumberCell } from "./extractor/xlsx-number-format";

/** Starting value to calibrate (10/11/2026): both cases run in a few milliseconds on healthy code, and the defect took
 *  4 600 to 6 400 ms (14/10/2026, not re-measured). 2 000 ms is the bound the three files of these cases share. */
const LIMIT_MS = 2_000;
const timed = <T>(fn: () => T): { value: T; ms: number } => {
  const start = performance.now();
  const value = fn();
  return { value, ms: performance.now() - start };
};

describe("linear-time regexes on hostile input", () => {
  it("smartart part number: a 64k-digit target", () => {
    // before the fix: 4.6 to 6.4 s
    const { value, ms } = timed(() => partNumberOf(`${"9".repeat(64_000)}x.xml`));
    expect(value).toBe(Number.MAX_SAFE_INTEGER);
    expect(partNumberOf("ppt/diagrams/drawing12.xml")).toBe(12);
    expect(ms).toBeLessThan(LIMIT_MS);
  });

  it("xlsx number format: many '[' and no ']' (code capped at 1024 chars: bounded, not red before the fix)", () => {
    const { ms } = timed(() => formatNumberCell(1, `0"${"[".repeat(1_000)}"`, false));
    expect(formatNumberCell(1234.5, "[Red]#,##0.00", false)).toBe("1,234.50");
    expect(ms).toBeLessThan(LIMIT_MS);
  });
});
