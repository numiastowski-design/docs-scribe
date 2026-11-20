/**
 * Each part is read once: a sheet's drawing part serves both its charts and its pictures, and a second read would
 * count twice towards `PACKAGE_MAX_BYTES`. The package is spied on (a file apart: `vi.mock` holds for the whole file).
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it, vi } from "vitest";
import { extractXlsx } from "./xlsx";
import { anchoredPictureWorkbook } from "./xlsx-anchored-picture.test.support";

const reads = new Map<string, number>();

vi.mock("./ooxml/package", async (importOriginal) => {
  const real = await importOriginal<typeof import("./ooxml/package")>();
  return {
    ...real,
    openPackage: async (...args: Parameters<typeof real.openPackage>) => {
      const pkg = await real.openPackage(...args);
      const count = (path: string): void => void reads.set(path, (reads.get(path) ?? 0) + 1);
      return {
        ...pkg,
        readXml: (path: string) => (count(path), pkg.readXml(path)),
        readBytes: (path: string, maxBytes?: number) => (count(path), pkg.readBytes(path, maxBytes)),
      };
    },
  };
});

const here = dirname(fileURLToPath(import.meta.url));

/** A drawing that carries a CHART (fixture 11) and one that carries a PICTURE (the made-up workbook of the support). */
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

describe("extractXlsx — every part is inflated once (revue 29/09, constat 10)", () => {
  for (const [name, bytesOf] of WORKBOOKS) {
    it(`🔴 ${name}: no part read twice, and a drawing part was read at all`, async () => {
      const bytes = await bytesOf();
      reads.clear();
      await extractXlsx(bytes);
      expect([...reads.keys()].some((p) => p.startsWith("xl/drawings/") && !p.includes("_rels"))).toBe(true);
      expect([...reads].filter(([, n]) => n > 1)).toEqual([]);
    });
  }
});
