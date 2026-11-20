/**
 * Review of 12/10/2026 (performance): the language of the booleans is sampled from the VISIBLE cells of the sheets read ahead
 * (`xlsx.ts`). Those cells were then built a SECOND time when the sheet's tab opened. The cells of a sheet are built once.
 * Own file: the `vi.mock` holds for the whole file.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it, vi } from "vitest";
import { extractXlsx } from "./xlsx";
import { anchoredPictureWorkbook } from "./xlsx-anchored-picture.test.support";

const built = new Map<string, number>();

vi.mock("./xlsx-cells", async (importOriginal) => {
  const real = await importOriginal<typeof import("./xlsx-cells")>();
  return {
    ...real,
    readWorksheetCells: (...args: Parameters<typeof real.readWorksheetCells>) => {
      const path = args[1];
      built.set(path, (built.get(path) ?? 0) + 1);
      return real.readWorksheetCells(...args);
    },
  };
});

const here = dirname(fileURLToPath(import.meta.url));

/** Fixture 11 (several sheets, a chart) and the made-up workbook of the support (one sheet, a picture). */
const WORKBOOKS: readonly (readonly [string, () => Promise<Uint8Array>])[] = [
  [
    "xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx",
    () =>
      Promise.resolve(
        new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx/11-FABRIQUE-Ventes-Graphique-Liens.xlsx"))),
      ),
  ],
  ["a made-up workbook with one anchored picture", anchoredPictureWorkbook],
];

describe("extractXlsx — the cells of a sheet are built once, the language sample included", () => {
  for (const [name, bytesOf] of WORKBOOKS) {
    it(`🔴 ${name}: no sheet's cells built twice`, async () => {
      const bytes = await bytesOf();
      built.clear();
      await extractXlsx(bytes, "en");
      expect(built.size, "fixture sanity: some sheet was read").toBeGreaterThan(0);
      // Failure value: every sheet read ahead for the language sample counts 2.
      expect([...built].filter(([, n]) => n > 1)).toEqual([]);
    });
  }
});
