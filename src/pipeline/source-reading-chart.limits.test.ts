/** The chart bounds are a PARAMETER (`ReaderLimits.chartPoints`, `chartSeries`, `chartLevels`, `chartCells`, plan step 5): the extractor and the
 *  control reading, given the SAME reduced limits, withhold exactly the same charts (« the extractor keeps the chart ⇔ the source expects its
 *  words »), on every edge of `source-reading-chart.caps.test.ts`, which holds the same rule by reducing the constants; and, given none, they keep
 *  the product's values. Run: npx vitest run --project=main src/pipeline/source-reading-chart.limits.test.ts */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractDocx } from "../extractor/docx";
import type { ReaderLimits } from "doc-ir";
import { readSource, wordsOf } from "./source-reading-docx";

const REDUCED = { MAX_CHART_POINTS: 20, MAX_CHART_CELLS: 200, MAX_CHART_SERIES: 12, MAX_CHART_LEVELS: 3 };
const LIMITS: ReaderLimits = { chartPoints: 20, chartCells: 200, chartSeries: 12, chartLevels: 3 };

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"`;
const drawing = `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;

/** The marker word every chart's title carries: on the page iff the chart is kept. */
const TITLE = "zebrafishtitle";

type Pts = readonly (readonly [number, string])[];
const pts = (entries: Pts): string => entries.map(([i, v]) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
const countOf = (entries: Pts): number => Math.max(...entries.map((e) => e[0]), 0) + 1;

interface SerSpec {
  cats?: Pts;
  levels?: Pts[];
  vals?: Pts;
}
const serXml = (s: SerSpec, i: number): string => {
  const cat = s.levels
    ? `<c:cat><c:multiLvlStrRef><c:f>S!$A$2</c:f><c:multiLvlStrCache><c:ptCount val="1"/>${s.levels.map((l) => `<c:lvl>${pts(l)}</c:lvl>`).join("")}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`
    : s.cats
      ? `<c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="${countOf(s.cats)}"/>${pts(s.cats)}</c:strCache></c:strRef></c:cat>`
      : "";
  const val = s.vals
    ? `<c:val><c:numRef><c:f>S!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${countOf(s.vals)}"/>${pts(s.vals)}</c:numCache></c:numRef></c:val>`
    : "";
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${cat}${val}</c:ser>`;
};
const chartOf = (series: readonly SerSpec[]): string =>
  `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>` +
  `<c:title><c:tx><c:rich><a:p><a:r><a:t>${TITLE}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart>${series.map(serXml).join("")}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;

const docxOf = async (chart: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="${REL}/chart" Target="charts/chart1.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document ${NS}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r>${drawing}</w:r></w:p></w:body></w:document>`,
  );
  zip.file("word/charts/chart1.xml", chart);
  return zip.generateAsync({ type: "uint8array" });
};

const dense = (n: number, v: string): Pts => Array.from({ length: n }, (_, i) => [i, v] as const);
const times = (count: number, spec: SerSpec): SerSpec[] => Array.from({ length: count }, () => spec);
const levelsOf = (count: number): Pts[] => Array.from({ length: count }, () => [[0, "Lvl"] as const]);

const keptBy = async (bytes: Uint8Array, limits?: ReaderLimits): Promise<{ kept: boolean; expected: boolean }> => ({
  kept: (await extractDocx(bytes, limits)).doc.blocks.some((b) => b.kind === "chart"),
  expected: wordsOf((await readSource(bytes, undefined, limits)).text).includes(TITLE),
});

type Case = { name: string; series: SerSpec[] };

const checkCase = async (c: Case): Promise<void> => {
  const { kept, expected } = await keptBy(await docxOf(chartOf(c.series)), LIMITS);
  expect(
    expected,
    `the extractor ${kept ? "keeps" : "withholds"} the chart, the source ${expected ? "expects" : "does not expect"} its words`,
  ).toBe(kept);
};

const P = REDUCED.MAX_CHART_POINTS;
const S = REDUCED.MAX_CHART_SERIES;
const L = REDUCED.MAX_CHART_LEVELS;
/** Series count: P rows x (1 label + KEPT) = the cell cap exactly; OVER = one more, over it. */
const KEPT = REDUCED.MAX_CHART_CELLS / P - 1;
const OVER = KEPT + 1;
const HALF = P / 2;

const cases: readonly Case[] = [
  { name: "one small series", series: [{ cats: [[0, "A"]], vals: [[0, "1"]] }] },
  // Cap = rows x (label columns + series); at most P rows (point cap).
  {
    name: `${P} rows x (1 label + ${KEPT} series) = the cell cap: kept (the cap is « over »)`,
    series: times(KEPT, { cats: dense(P, "C"), vals: dense(P, "1") }),
  },
  {
    name: `${P} rows x (1 label + ${OVER} series) = one series over the cell cap: withheld`,
    series: times(OVER, { cats: dense(P, "C"), vals: dense(P, "1") }),
  },
  {
    name: `sparse: one value at idx ${P - 1} and ${OVER} series (1 occupied row, array length ${P} x ${OVER + 1} columns: the extractor withholds on array length)`,
    series: times(OVER, { cats: [[0, "A"]], vals: [[P - 1, "1"]] }),
  },
  {
    name: `sparse, no c:cat: one value at idx ${P - 1} and ${OVER} series (array length ${P} x ${OVER + 1} columns)`,
    series: times(OVER, { vals: [[P - 1, "1"]] }),
  },
  {
    name: `no c:cat: the label column is empty but still a column (${KEPT} series x ${P} rows + 1 = the cell cap)`,
    series: times(KEPT, { vals: dense(P, "1") }),
  },
  {
    name: `no c:cat: ${OVER} series x ${P} rows + the empty label column = over the cell cap`,
    series: times(OVER, { vals: dense(P, "1") }),
  },
  {
    name: `empty values occupy nothing: no c:cat, ${P - 1} rows of empty values, ${S - 1} series (array length over the cell cap)`,
    series: times(S - 1, { vals: dense(P - 1, "") }),
  },
  {
    name: "empty values occupy nothing: non-empty categories on the FIRST series, empty values elsewhere",
    series: [{ cats: dense(P - 2, "C"), vals: dense(P - 2, "") }, ...times(KEPT, { vals: dense(HALF, "") })],
  },
  {
    name: "categories come from the first series that has some: later series' categories are ignored",
    series: [
      { cats: dense(HALF, "C"), vals: dense(HALF, "1") },
      ...times(KEPT, { cats: dense(P - 1, "C"), vals: dense(HALF, "1") }),
    ],
  },
  {
    name: `categories come from the first series that has some: the first has none, the second holds ${P}`,
    series: [
      { vals: dense(HALF, "1") },
      { cats: dense(P, "C"), vals: dense(HALF, "1") },
      ...times(KEPT - 1, { vals: dense(HALF, "1") }),
    ],
  },
  { name: `point cap: a value at idx ${P} withholds the chart`, series: [{ cats: [[0, "A"]], vals: [[P, "1"]] }] },
  { name: `point cap: a value at idx ${P - 1} is kept`, series: [{ cats: [[0, "A"]], vals: [[P - 1, "1"]] }] },
  {
    name: `point cap on categories: a category at idx ${P} withholds the chart`,
    series: [{ cats: [[P, "A"]], vals: [[0, "1"]] }],
  },
  {
    name: "point cap on a LATER series' categories, which the extractor never reads",
    series: [
      { cats: [[0, "A"]], vals: [[0, "1"]] },
      { cats: [[P, "A"]], vals: [[0, "1"]] },
    ],
  },
  { name: `level cap: ${L} category levels are kept`, series: [{ levels: levelsOf(L), vals: [[0, "1"]] }] },
  {
    name: `level cap: ${L + 1} category levels withhold the chart`,
    series: [{ levels: levelsOf(L + 1), vals: [[0, "1"]] }],
  },
  { name: `${S} series (one point each) are kept`, series: times(S, { cats: [[0, "A"]], vals: [[0, "1"]] }) },
  { name: `${S + 1} series are withheld`, series: times(S + 1, { cats: [[0, "A"]], vals: [[0, "1"]] }) },
];

describe("chart bounds given as limits: the extractor and the control reading withhold the same charts", () => {
  for (const c of cases) it(c.name, () => checkCase(c), 60_000);

  it("the table is not vacuous: under the limits half its cases are withheld, and none is under the product's values", async () => {
    const outcomes = await Promise.all(cases.map(async (c) => keptBy(await docxOf(chartOf(c.series)), LIMITS)));
    const byDefault = await Promise.all(cases.map(async (c) => keptBy(await docxOf(chartOf(c.series)), undefined)));
    // FAILS WITH (limits ignored): expected 0 to be greater than 6.
    expect(outcomes.filter((o) => !o.kept).length).toBeGreaterThan(6);
    expect(byDefault.filter((o) => !o.kept)).toEqual([]);
  }, 120_000);
});

describe("chart bounds: absent limits cap NOTHING, and each bound bites alone when it is given", () => {
  const twelve = chartOf(times(S + 1, { cats: [[0, "A"]], vals: [[0, "1"]] }));
  it("13 series are kept by default and withheld under chartSeries: 12, by both readings", async () => {
    const bytes = await docxOf(twelve);
    // FAILS WITH (limits ignored): expected true to be false (the chart is still kept under the lowered bound).
    expect(await keptBy(bytes, { chartSeries: 12 })).toEqual({ kept: false, expected: false });
    expect(await keptBy(bytes, undefined)).toEqual({ kept: true, expected: true });
    expect(await keptBy(bytes, { chartPoints: 20 })).toEqual({ kept: true, expected: true });
  }, 60_000);
  it("each other bound bites alone: points, levels, cells", async () => {
    const point = await docxOf(chartOf([{ cats: [[0, "A"]], vals: [[P, "1"]] }]));
    expect(await keptBy(point, { chartPoints: P })).toEqual({ kept: false, expected: false });
    expect(await keptBy(point)).toEqual({ kept: true, expected: true });
    const lvl = await docxOf(chartOf([{ levels: levelsOf(L + 1), vals: [[0, "1"]] }]));
    expect(await keptBy(lvl, { chartLevels: L })).toEqual({ kept: false, expected: false });
    expect(await keptBy(lvl)).toEqual({ kept: true, expected: true });
    const cells = await docxOf(chartOf(times(4, { cats: dense(10, "C"), vals: dense(10, "1") })));
    expect(await keptBy(cells, { chartCells: 40 })).toEqual({ kept: false, expected: false });
    expect(await keptBy(cells, { chartCells: 50 })).toEqual({ kept: true, expected: true });
    expect(await keptBy(cells)).toEqual({ kept: true, expected: true });
  }, 60_000);
});
