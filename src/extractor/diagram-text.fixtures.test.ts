import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { Inline, IRBlock } from "doc-ir";
import { extractDocx } from "./docx";
import { extractPptx } from "./pptx";
import { extractXlsx } from "./xlsx";
import { readPptx } from "./pptx-reader";

// Lost-text defects of 01/10/2026: the words of a SmartArt (06 lost all 9 of
// its texts, 07 28 of 32), and a native Word chart (the docx reader dropped it
// whole, title included). The expected values are read here from the file's own XML, by an
// independent regex: the reader's output must carry every one of them.
vi.setConfig({ testTimeout: 30_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));

const unescape = (s: string): string =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

const words = (text: string): string[] => text.split(/\s+/).filter((w) => w !== "");

/** The words of a SmartArt data model: one string per `a:p` (runs joined,
 *  as a reader does), split on blanks. */
const dataModelWords = (xml: string): string[] =>
  (xml.match(/<a:p>.*?<\/a:p>/gs) ?? []).flatMap((p) =>
    words(unescape([...p.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(""))),
  );

const count = (list: readonly string[]): Map<string, number> => {
  const m = new Map<string, number>();
  for (const w of list) m.set(w, (m.get(w) ?? 0) + 1);
  return m;
};

/** What the words of `expected` lack in `actual`, by count. */
const missing = (expected: readonly string[], actual: readonly string[]): string[] => {
  const have = count(actual);
  const out: string[] = [];
  for (const w of expected) {
    const n = have.get(w) ?? 0;
    if (n === 0) out.push(w);
    else have.set(w, n - 1);
  }
  return out;
};

const inlineText = (inlines: readonly Inline[]): string => inlines.map((i) => i.text).join("");

const diagramWords = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((b) =>
    b.kind === "diagramText" ? b.sections.flatMap((s) => words(inlineText(s.inlines).replace(/ · /g, " "))) : [],
  );

describe("SmartArt words (milestone 0)", () => {
  for (const [file, parts] of [
    ["diagrammes-reel/06-fr-la-veille.pptx", ["data1", "data2"]],
    ["diagrammes-reel/07-fr-trouver-une-image.pptx", ["data1", "data2", "data3", "data4"]],
  ] as const) {
    it(`${file}: every word of the data models is in the diagram text`, async () => {
      const zip = await JSZip.loadAsync(fixture(file));
      const expected: string[] = [];
      for (const p of parts) expected.push(...dataModelWords(await zip.file(`ppt/diagrams/${p}.xml`)!.async("string")));
      const { doc } = await extractPptx(fixture(file));
      const lost = missing(expected, diagramWords(doc.blocks));
      // Failure value (before): the whole data model lost (06: 0 of 45 words kept).
      expect(`${lost.length} of ${expected.length} words lost`).toBe(`0 of ${expected.length} words lost`);
      // Conservation both ways: what the reader hands on is EXACTLY the data
      // models' words (nothing added, nothing repeated). All slides, no cap.
      const read = await readPptx(fixture(file));
      const given = read.slides.flatMap((s) =>
        s.shapes.flatMap((sh) => (sh.kind === "graphic" && sh.words ? sh.words.flatMap(words) : [])),
      );
      expect([...given].sort()).toEqual([...expected].sort());
    });
  }
});

describe("native Word chart", () => {
  for (const file of [
    "g-fr-graphique-barres.docx",
    "g-en-bar-chart.docx",
    "h-fr-graphique-secteurs.docx",
    "h-en-pie-chart.docx",
  ]) {
    it(`${file}: title, series, categories and values of the chart cache are published as a chart block`, async () => {
      const zip = await JSZip.loadAsync(fixture(`diagrammes-generes/${file}`));
      const xml = unescape(await zip.file("word/charts/chart1.xml")!.async("string"));
      const title = [...xml.match(/<c:title>.*?<\/c:title>/s)![0].matchAll(/<a:t>([^<]*)<\/a:t>/g)]
        .map((m) => m[1]!)
        .join("");
      const series = [...xml.matchAll(/<c:ser>.*?<\/c:ser>/gs)].flatMap((s) =>
        [...s[0].match(/<c:tx>.*?<\/c:tx>/s)![0].matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((m) => m[1]!),
      );
      const first = xml.match(/<c:ser>.*?<\/c:ser>/s)![0];
      const categories = [...first.match(/<c:cat>.*?<\/c:cat>/s)![0].matchAll(/<c:v>([^<]*)<\/c:v>/g)].map(
        (m) => m[1]!,
      );
      const values = [...xml.matchAll(/<c:val>.*?<\/c:val>/gs)].flatMap((v) =>
        [...v[0].matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((m) => m[1]!),
      );
      const { doc } = await extractDocx(fixture(`diagrammes-generes/${file}`));
      const charts = doc.blocks.filter((b) => b.kind === "chart");
      // Failure value (before): 0 chart blocks, the chart gone without trace.
      expect(`${charts.length} chart block(s)`).toBe("1 chart block(s)");
      const chart = charts[0]!;
      if (chart.kind !== "chart") throw new Error("unreachable");
      const published = [chart.title ?? "", ...(chart.data.headers ?? []), ...chart.data.rows.flat()];
      const expected = [title, ...series, ...categories, ...values];
      const lost = expected.filter((e) => !published.includes(e));
      expect(lost).toEqual([]);
      expect(expected.length).toBeGreaterThanOrEqual(8);
    });
  }
});

const chartBlocks = (
  value: unknown,
  out: Extract<IRBlock, { kind: "chart" }>[] = [],
): Extract<IRBlock, { kind: "chart" }>[] => {
  if (Array.isArray(value)) for (const v of value) chartBlocks(v, out);
  else if (typeof value === "object" && value !== null) {
    if ((value as { kind?: string }).kind === "chart") out.push(value as Extract<IRBlock, { kind: "chart" }>);
    for (const v of Object.values(value)) chartBlocks(v, out);
  }
  return out;
};

describe("categories on two levels (c:multiLvlStrRef): every word of every level is published", () => {
  // Generated by us from the g-fr file: leaves Janvier..Avril, outer level T1 (idx 0) and T2 (idx 2).
  const EXPECTED = ["Janvier", "Février", "Mars", "Avril", "T1", "T2"];
  for (const [ext, run] of [
    ["docx", async (b: Uint8Array) => (await extractDocx(b)).doc],
    ["pptx", async (b: Uint8Array) => (await extractPptx(b)).doc],
    ["xlsx", async (b: Uint8Array) => (await extractXlsx(b)).doc],
  ] as const) {
    it(`.${ext}: the chart block carries both levels`, async () => {
      const doc = await run(fixture(`diagrammes-generes/i-fr-categories-deux-niveaux.${ext}`));
      const charts = chartBlocks(doc.blocks);
      expect(charts.length).toBeGreaterThan(0);
      const cells = charts[0]!.data.rows.flat();
      // Failure value (before): the outer level T1, T2 lost.
      const lost = EXPECTED.filter((w) => !cells.includes(w));
      expect(lost).toEqual([]);
    });
  }
});

// A declared table stays a table on a slide the layout calls a diagram
// (02/10/2026): 03-pvi slide 7 has one a:tbl (88 cells) and five connectors.
// Its cells were flattened into the diagram paragraph, structure lost.
describe("a declared table on a diagram slide", () => {
  const file = "diagrammes-reel/03-pvi-decision-trees.pptx";
  const tabOf = (doc: { blocks: readonly IRBlock[] }, n: number) => {
    const tabs = doc.blocks[0]?.kind === "tabs" ? doc.blocks[0].tabs : [];
    return tabs.find((t) => t.title.split(" ").pop() === String(n))!;
  };
  const tableTexts = (blocks: readonly IRBlock[]): string[] =>
    blocks.flatMap((b) =>
      b.kind === "table"
        ? b.rows.flatMap((r) =>
            r.cells.map((c) =>
              c.blocks
                .flatMap((x) => (x.kind === "paragraph" ? [inlineText(x.inlines)] : []))
                .join(" ")
                .replace(/\s+/g, " ")
                .trim(),
            ),
          )
        : [],
    );

  it("slide 7: the a:tbl cells are a table block, not diagram text", async () => {
    const zip = await JSZip.loadAsync(fixture(file));
    const xml = await zip.file("ppt/slides/slide7.xml")!.async("string");
    const tbl = xml.match(/<a:tbl>.*<\/a:tbl>/s)![0];
    const expected = [...tbl.matchAll(/<a:tc[ >].*?<\/a:tc>/gs)]
      .map((tc) =>
        unescape(
          [...tc[0].matchAll(/<a:p>.*?<\/a:p>/gs)]
            .map((para) => [...para[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(""))
            .join(" "),
        )
          .replace(/\s+/g, " ")
          .trim(),
      )
      .filter((t) => t !== "");
    expect(expected.length).toBe(46);
    const { doc } = await extractPptx(fixture(file));
    const published = tableTexts(tabOf(doc, 7).blocks);
    // Failure value (before): 0 table blocks on the slide, every cell flattened into the diagram text.
    expect(expected.filter((t) => !published.includes(t))).toEqual([]);
    expect(`${expected.filter((t) => published.includes(t)).length} of ${expected.length} cells in a table block`).toBe(
      `${expected.length} of ${expected.length} cells in a table block`,
    );
    const section = doc.blocks
      .flatMap((b) => (b.kind === "diagramText" ? b.sections : []))
      .find((s) => s.label === "Slide 7")!;
    const flat = inlineText(section.inlines);
    // A cell whose text a shape outside the table also carries proves nothing.
    const outside = unescape(xml.replace(tbl, "")).replace(/<[^>]*>/g, "");
    const flattened = expected.filter((t) => t.includes(" ") && !outside.includes(t) && flat.includes(t));
    expect(flattened).toEqual([]);
  });

  it("slide 10 (a flow slide with a table) is unchanged", async () => {
    const { doc } = await extractPptx(fixture(file));
    const tab = tabOf(doc, 10);
    expect(tab.blocks.map((b) => b.kind)).toEqual(["heading", "table"]);
    const sections = doc.blocks.flatMap((b) => (b.kind === "diagramText" ? b.sections : []));
    expect(sections.some((s) => s.label === "Slide 10")).toBe(false);
  });
});
