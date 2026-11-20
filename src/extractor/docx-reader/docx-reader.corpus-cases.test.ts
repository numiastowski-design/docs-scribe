/** Real Word documents of the large local corpus (`fixtures/docx-corpus-cases/README.md` gives each one's source line).
 *  The 200 file is the SSHRC report with its embedded workbooks and its pictures taken out (469 KB for 14 MB); the words and the
 *  32 chart parts are the original's. Rule held here: a chart's data table shows what the chart DISPLAYS (19/10/2026;
 *  the rule itself, with its precedence, is `extractor/pptx-reader/chart.format.test.ts`). */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "../docx";

vi.setConfig({ testTimeout: 60_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../../fixtures/docx-corpus-cases", name)));

const chartsOf = (
  value: unknown,
  out: Extract<IRBlock, { kind: "chart" }>[] = [],
): Extract<IRBlock, { kind: "chart" }>[] => {
  if (Array.isArray(value)) for (const v of value) chartsOf(v, out);
  else if (typeof value === "object" && value !== null) {
    if ((value as { kind?: string }).kind === "chart") out.push(value as Extract<IRBlock, { kind: "chart" }>);
    for (const v of Object.values(value)) chartsOf(v, out);
  }
  return out;
};

describe("200-2022-23-sshrcs-report-on-competition.docx: 32 charts, 0% / 0.0% / General caches", () => {
  it("🔴 chart13 « Language of correspondence of survey respondents (2022-23) » shows 81% and 19%, not 0.80861244019138756 and 0.19138755980861244", async () => {
    const charts = chartsOf((await extractDocx(fixture("200-2022-23-sshrcs-report-on-competition.docx"))).doc.blocks);
    expect(charts.length).toBe(32);
    const chart13 = charts.find(
      (c) =>
        c.title?.startsWith("Language of correspondence of survey respondents") === true && c.data.rows.length === 2,
    );
    expect(chart13?.data.rows.flat()).toEqual(expect.arrayContaining(["English", "81%", "French", "19%"]));
  });

  it("🔴 no cell of any of the 32 charts is a stored double (7 decimals or more, or an exponent): Word displays none of them", async () => {
    const charts = chartsOf((await extractDocx(fixture("200-2022-23-sshrcs-report-on-competition.docx"))).doc.blocks);
    const stored = charts.flatMap((c) => c.data.rows.flat()).filter((cell) => /\d\.\d{7,}|\dE-?\d/i.test(cell));
    // Red today: 133 cells (« 0.80861244019138756 », « 2.5477707006369428E-2 », « 37.799999999999997 »).
    expect(stored).toEqual([]);
  });
});
