/** The control reading must withhold EXACTLY the charts the extractor withholds (`extractor/pptx-reader/chart.ts#tableOf` against
 *  `pipeline/source-reading-chart.ts#chartIsWithheld`): for every case, « the extractor keeps the chart ⇔ the source expects its words ».
 *  Each case is one `c:chartSpace` written here; the extractor side is `extractDocx` on a one-chart docx (the published page), the source
 *  side is `readSource`. Cases sit on the edges: occupied cells (non-empty values), categories of the first series that has some, the
 *  label column when there is no `c:cat`, the sparse index, the caps on cells, points and category levels.
 *  Run: npx vitest run --project=main src/pipeline/source-reading-chart.caps.test.ts */
import JSZip from "jszip";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { wordsOf } from "./source-reading-docx";

/** Both readings import the SAME `lib/reader-limits`, so the rule under test (« kept ⇔ expected ») is the same at any cap. The table below
 *  runs on REDUCED caps (the real ones cost 6 to 14 s a case: 10 000 rows x 20 series); ONE case, the guard, runs on the real values.
 *  The reduced values keep every edge reachable: ROWS x (1 label + SERIES_KEPT) = CELLS, and SERIES_KEPT < SERIES. `vi.mock` takes in the
 *  `main` project only (the `pdf` project has `isolate: false`). */
const { REDUCED, mode } = vi.hoisted(() => ({
  REDUCED: { MAX_CHART_POINTS: 20, MAX_CHART_CELLS: 200, MAX_CHART_SERIES: 12, MAX_CHART_LEVELS: 3 },
  mode: { reduced: true },
}));
vi.mock("doc-ir/internal", async (importOriginal) => {
  const actual = await importOriginal<typeof import("doc-ir/internal")>();
  // Getters, not a snapshot: vitest keeps the factory's result across `resetModules`, the mode changes between the two describes.
  const caps = { ...actual };
  for (const [name, value] of Object.entries(REDUCED))
    Object.defineProperty(caps, name, {
      enumerable: true,
      get: () => (mode.reduced ? value : (actual as Record<string, unknown>)[name]),
    });
  return caps;
});

type ExtractDocx = typeof import("../extractor/docx").extractDocx;
type ReadSource = typeof import("./source-reading-docx").readSource;
let extractDocx: ExtractDocx;
let readSource: ReadSource;
/** Fresh modules under the caps of the given mode: the extractor and the control reading re-import `lib/reader-limits`. */
const load = async (reduced: boolean): Promise<void> => {
  mode.reduced = reduced;
  vi.resetModules();
  ({ extractDocx } = await import("../extractor/docx"));
  ({ readSource } = await import("./source-reading-docx"));
};

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

type Case = { name: string; series: SerSpec[] };

const checkCase = async (c: Case): Promise<void> => {
  const bytes = await docxOf(chartOf(c.series));
  // Kept = a chart block in the IR (a table of 200 000 cells would hit the page limit (`CONTENT_LIMIT_BYTES`), which is another withholding).
  const kept = (await extractDocx(bytes)).doc.blocks.some((b) => b.kind === "chart");
  const expected = wordsOf((await readSource(bytes)).text).includes(TITLE);
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

describe("chart caps (reduced caps): the control reading expects the words of a chart iff the extractor publishes them", () => {
  beforeAll(() => load(true));
  for (const c of cases) it(c.name, () => checkCase(c), 60_000);
});

describe("chart points with an unreadable idx", () => {
  beforeAll(() => load(true));
  it('a `c:pt idx="abc"` is ignored by BOTH readings (no false lost word)', async () => {
    const ghost = "ghostwordxyz";
    const bytes = await docxOf(
      chartOf([
        {
          cats: [
            [0, "A"],
            ["abc" as unknown as number, ghost],
          ],
          vals: [[0, "1"]],
        },
      ]),
    );
    const expected = wordsOf((await readSource(bytes)).text).includes(ghost);
    const published = JSON.stringify((await extractDocx(bytes)).doc).includes(ghost);
    // Failure value (before): the control reading expects the word (true), the extractor drops it (false).
    expect(expected).toBe(published);
  }, 60_000);

  it("an idx of 400 digits (parseInt gives Infinity) still withholds the chart, as any idx past the bound", async () => {
    const bytes = await docxOf(
      chartOf([
        {
          cats: [
            [0, "A"],
            ["9".repeat(400) as unknown as number, "B"],
          ],
          vals: [[0, "1"]],
        },
      ]),
    );
    // Failure value: the point skipped as « unreadable », the chart published without it (a silent loss).
    expect(JSON.stringify((await extractDocx(bytes)).doc)).not.toContain('"kind":"chart"');
  }, 60_000);
});

describe("chart caps (real values, the guard): the same rule at the real edges", () => {
  beforeAll(() => load(false));
  // The cheapest case that pins the real numbers: 256 series, one past Excel's limit (MAX_CHART_SERIES = 255).
  it("256 series are withheld, 255 are kept", async () => {
    expect((await import("doc-ir/internal")).MAX_CHART_SERIES).toBe(255);
    await checkCase({ name: "256", series: times(256, { cats: [[0, "A"]], vals: [[0, "1"]] }) });
    await checkCase({ name: "255", series: times(255, { cats: [[0, "A"]], vals: [[0, "1"]] }) });
  }, 60_000);
});
