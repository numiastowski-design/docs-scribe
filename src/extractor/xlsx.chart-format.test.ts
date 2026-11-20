/** The charts of a workbook go through the same reader as Word's and PowerPoint's (`pptx-reader/chart.ts`), so they carry the
 *  same defect: a chart's data table shows the cache's stored double (« 0.22830667532314231 ») where the chart displays « 22.83% »
 *  (`pptx-reader/chart.format.test.ts` holds the rule). Witness: fixtures/xlsx-heading-reel/34-barnet-london-performance-2015.xlsx,
 *  14 charts under « 0% », « 0.00% », « 0.00 », « #,##0.00 ». */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractXlsx } from "./xlsx";

vi.setConfig({ testTimeout: 60_000 });

const here = dirname(fileURLToPath(import.meta.url));

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

describe("xlsx charts show displayed values", () => {
  it("🔴 34-barnet-london-performance-2015: no chart cell is a stored double (7 decimals or more, or an exponent)", async () => {
    const bytes = new Uint8Array(
      readFileSync(join(here, "../../fixtures/xlsx-heading-reel/34-barnet-london-performance-2015.xlsx")),
    );
    const charts = chartsOf((await extractXlsx(bytes, "en")).doc.blocks);
    expect(charts.length).toBeGreaterThan(0);
    const stored = charts.flatMap((c) => c.data.rows.flat()).filter((cell) => /\d\.\d{7,}|\dE-?\d/i.test(cell));
    expect(stored).toEqual([]);
  });
});
