/** The premise of drawing a chart from its sheet (docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md): where Excel wrote the cached values
 *  of a chart, the values read from the sheet that holds it ARE those values. Witnesses: the charts of
 *  fixtures/xlsx-heading-reel/03-… and 34-… (Excel's, cache filled), same-sheet references only. Compared point by point:
 *  the occupied indices, the stored numbers (strictly), the format code the cell carries, and the cell's text for a text point. */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import { attributeOf, findChild, findChildren, findDescendants, textContentOf, type XNode } from "./ooxml/xml";
import type { ChartResolver } from "./xlsx-chart-reference";
import { chartResolverOf, keptGridOf } from "./xlsx";
import { drawingPartsOf, readWorkbookSheets } from "./xlsx-parts";
import { formatNumberCell } from "./xlsx-number-format";
import { readSharedStringsAndStyles, readWorksheetCells } from "./xlsx-cells";

vi.setConfig({ testTimeout: 60_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx-heading-reel", name)));

/** The references whose cell code is NOT the cache's, as measured (docs/CONCEPTION-GRAPHIQUE-SANS-CACHE.md, question 5 names the first: chart 1
 *  of fixture 03, whose cache holds a localized code; the design wrote it `#\ ##0`, the real code is the whole accounting one below). The two
 *  of fixture 34 are not in the design: found by this test. Every point's code is compared (`cellCode` lists the distinct codes of the range, in
 *  order): the two of fixture 34 mix codes, and its chart 10 holds General cells among `0.00%` ones (found when the comparison stopped reading
 *  only the first point). For these, only the numbers are compared (the displayed text is compared where the codes are equal); the list is
 *  closed, so a new difference turns the test red. */
const NAMED_CODE_DIFFERENCES = [
  {
    where: "Graphiques/chart 1 Graphiques!$B$2:$B$7",
    cacheCode: String.raw`_-* #\ ##0_-;\-* #\ ##0_-;_-* "-"??_-;_-@_-`,
    cellCode: String.raw`_-* #,##0_-;\-* #,##0_-;_-* "-"??_-;_-@_-`,
  },
  {
    where: "Base LA Summary Data/chart 1 'Base LA Summary Data'!$C$10:$G$10",
    cacheCode: "0%",
    cellCode: "##0.00% | 0%",
  },
  {
    where: "Base LA Summary Data/chart 2 'Base LA Summary Data'!$C$9:$G$9",
    cacheCode: "0.00",
    cellCode: "#,##0.00 | 0.00",
  },
  {
    where: "2013-14 Kerbside split/chart 10 '2013-14 Kerbside split'!$H$4:$H$37",
    cacheCode: "0.00%",
    cellCode: "0.00% | General",
  },
] as const;

interface Mismatch {
  where: string;
  what: string;
}
interface Tally {
  references: number;
  points: number;
  mismatches: Mismatch[];
  codeDifferences: { where: string; cacheCode: string; cellCode: string | undefined }[];
  skippedOtherSheet: number;
}

const cachePoints = (cache: XNode): { index: number; value: string }[] =>
  findChildren(cache, "c:pt").map((pt) => ({
    index: Number(attributeOf(pt, "idx")),
    value: textContentOf(findChild(pt, "c:v") ?? {}),
  }));

const compareReference = (
  ref: XNode,
  where: string,
  resolver: ChartResolver,
  tally: Tally,
  perturb: ((value: number) => number) | undefined,
): void => {
  const tag = findChild(ref, "c:numCache") ? "c:numCache" : "c:strCache";
  const cache = findChild(ref, tag);
  const formula = textContentOf(findChild(ref, "c:f") ?? {});
  if (cache === undefined || cachePoints(cache).length === 0) return;
  const resolved = resolver(formula);
  if (resolved === undefined) {
    tally.skippedOtherSheet++;
    return;
  }
  tally.references++;
  const numeric = tag === "c:numCache";
  const expected = cachePoints(cache);
  const got = [...resolved].filter((point) => !numeric || point.stored !== undefined);
  const missing = (cause: string): void => void tally.mismatches.push({ where: `${where} ${formula}`, what: cause });
  if (expected.map((p) => p.index).join() !== got.map((p) => p.index).join())
    return missing(`indices cache [${expected.map((p) => p.index).join()}] sheet [${got.map((p) => p.index).join()}]`);
  const cacheCode = textContentOf(findChild(cache, "c:formatCode") ?? {});
  // Every point's own code, not the first one's: a range may mix formats.
  const cellCodes = [...new Set(got.map((point) => point.formatCode ?? "General"))];
  const sameCode = !numeric || cacheCode === "" || cellCodes.every((code) => code === cacheCode);
  if (!sameCode)
    tally.codeDifferences.push({ where: `${where} ${formula}`, cacheCode, cellCode: cellCodes.join(" | ") });
  for (const [position, point] of expected.entries()) {
    tally.points++;
    const cell = got[position]!;
    if (!numeric) {
      // A text reference shows the grid's text, a numeric cell's included (`pptx-reader/chart.ts#textCells`): compared to the cache's text.
      if (cell.text !== point.value) missing(`idx ${point.index}: cache "${point.value}" sheet "${cell.text}"`);
      continue;
    }
    const value = perturb === undefined || position !== 0 ? Number(cell.stored) : perturb(Number(cell.stored));
    if (value !== Number(point.value)) missing(`idx ${point.index}: cache ${point.value} sheet ${String(value)}`);
    else if (
      sameCode &&
      formatNumberCell(value, cell.formatCode, false) !==
        formatNumberCell(Number(point.value), cacheCode === "" ? undefined : cacheCode, false)
    )
      missing(`idx ${point.index}: displayed text differs`);
  }
};

const parity = async (file: string, perturb?: (value: number) => number): Promise<Tally> => {
  const { pkg, sheets, workbookXml } = await readWorkbookSheets(fixture(file));
  const workbook = await readSharedStringsAndStyles(pkg, workbookXml);
  const tally: Tally = { references: 0, points: 0, mismatches: [], codeDifferences: [], skippedOtherSheet: 0 };
  let chartOrdinal = 0;
  for (const info of sheets) {
    if (info.kind !== "worksheet" || info.state !== "visible") continue;
    const sheetXml = (await pkg.readXml(info.partPath))!;
    const cells = await readWorksheetCells(pkg, info.partPath, sheetXml, workbook);
    const resolver = chartResolverOf(info.name, cells, keptGridOf(cells, workbook.date1904, "en"));
    for (const part of (await drawingPartsOf(pkg, info.partPath, sheetXml)).charts) {
      chartOrdinal++;
      const root = (await pkg.readXml(part))!;
      for (const ser of findDescendants(root, "c:ser")) {
        for (const ref of [...findDescendants(ser, "c:numRef"), ...findDescendants(ser, "c:strRef")])
          compareReference(ref, `${info.name}/chart ${chartOrdinal}`, resolver, tally, perturb);
      }
    }
  }
  return tally;
};

describe("a chart's cached values equal the sheet's values (Excel's charts)", () => {
  for (const file of ["03-bilan-recolement-paca.xlsx", "34-barnet-london-performance-2015.xlsx"]) {
    it(`🔴 ${file}: every same-sheet reference reads the cache's points from the sheet`, async () => {
      const tally = await parity(file);
      expect(tally.mismatches).toEqual([]);
      // The fixture exercises the comparison: there are references and points to compare, not an empty pass.
      expect(tally.references).toBeGreaterThan(0);
      expect(tally.points).toBeGreaterThan(0);
    });
  }

  it("🔴 the references whose cell code differs from the cache's are exactly the named ones", async () => {
    const found = [
      ...(await parity("03-bilan-recolement-paca.xlsx")).codeDifferences,
      ...(await parity("34-barnet-london-performance-2015.xlsx")).codeDifferences,
    ];
    expect(found).toEqual(NAMED_CODE_DIFFERENCES.map((named) => ({ ...named })));
  });

  it("🔴 the comparison bites: one cell value changed in memory is reported", async () => {
    const tally = await parity("34-barnet-london-performance-2015.xlsx", (value) => value + 1);
    expect(tally.mismatches.length).toBeGreaterThan(0);
    expect(tally.mismatches[0]!.what).toMatch(/cache .* sheet /);
  });
});
