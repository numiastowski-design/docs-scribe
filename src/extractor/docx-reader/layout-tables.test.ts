import { describe, expect, it } from "vitest";

import { wordsOf } from "../../lib/words";
import type { IRBlock } from "doc-ir";
import { LAYOUT_TABLE_MIN_WORDS, columnsBackToTables, layoutTablesToColumns } from "./layout-tables";

// A pure pass over the BODY blocks turns a one-row table of 2 or 3
// cells, each prose of at least LAYOUT_TABLE_MIN_WORDS words or a picture-only cell, into a `columns` block. Anything else is returned as is.
// Tables are written by hand in the IR; each case names its failure value (the table the reader gives today).

type Cell = { blocks: IRBlock[]; colspan?: number; rowspan?: number; header?: boolean };
type Table = Extract<IRBlock, { kind: "table" }>;

const p = (text: string): IRBlock => ({ kind: "paragraph", inlines: [{ text }] });
const img = (sha = "a"): IRBlock => ({ kind: "image", ref: { sha256: sha, mimeType: "image/png" } });
/** `n` distinct words for the words tokenizer. */
const words = (n: number, tag = "mot"): string =>
  Array.from({ length: n }, (_, i) => `${tag}${String.fromCharCode(97 + (i % 26))}${Math.floor(i / 26)}`).join(" ");
const prose = (n: number, tag = "mot"): Cell => ({ blocks: [p(words(n, tag))] });
const picture = (sha = "a"): Cell => ({ blocks: [img(sha)] });
const oneRow = (...cells: Cell[]): Table => ({ kind: "table", rows: [{ cells }] });
const only = (blocks: IRBlock[]): IRBlock => {
  expect(blocks).toHaveLength(1);
  return blocks[0]!;
};

describe("the constant and the tokenizer it counts with", () => {
  it("LAYOUT_TABLE_MIN_WORDS is 10 (a starting value to calibrate, measured on 369 Word documents)", () => {
    expect(LAYOUT_TABLE_MIN_WORDS).toBe(10);
  });

  it("the helper writes the number of words the app's tokenizer reads (fixture sanity)", () => {
    expect(wordsOf(words(10))).toHaveLength(10);
    expect(wordsOf(words(9))).toHaveLength(9);
  });
});

describe("a prose table becomes columns", () => {
  it("1 row x 2 cells of prose: `columns` two_equal, each cell's blocks unchanged and in order", () => {
    const left: IRBlock[] = [{ kind: "heading", level: 2, inlines: [{ text: "Titre de gauche" }] }, p(words(12, "g"))];
    const right: IRBlock[] = [p(words(11, "d")), { kind: "list", ordered: false, items: [[p("un"), p("deux")]] }];
    // Failure value: the same `table`.
    expect(only(layoutTablesToColumns([oneRow({ blocks: left }, { blocks: right })]))).toEqual({
      kind: "columns",
      layout: "two_equal",
      cells: [left, right],
    });
  });

  it("1 row x 3 cells of prose: three_equal", () => {
    const out = only(layoutTablesToColumns([oneRow(prose(10, "a"), prose(20, "b"), prose(30, "c"))]));
    expect(out).toMatchObject({ kind: "columns", layout: "three_equal" });
    expect((out as Extract<IRBlock, { kind: "columns" }>).cells).toHaveLength(3);
  });

  it("image | prose: two_left_sidebar (the picture cell keeps its blocks)", () => {
    expect(only(layoutTablesToColumns([oneRow(picture("x"), prose(15))]))).toEqual({
      kind: "columns",
      layout: "two_left_sidebar",
      cells: [[img("x")], [p(words(15))]],
    });
  });

  it("prose | image: two_right_sidebar", () => {
    expect(only(layoutTablesToColumns([oneRow(prose(15), picture("y"))]))).toMatchObject({
      kind: "columns",
      layout: "two_right_sidebar",
    });
  });

  it("image | prose | image: three_with_sidebars", () => {
    expect(only(layoutTablesToColumns([oneRow(picture("l"), prose(15), picture("r"))]))).toMatchObject({
      kind: "columns",
      layout: "three_with_sidebars",
    });
  });

  it("a picture cell may hold several pictures (052: prose | two pictures)", () => {
    const two: Cell = { blocks: [img("1"), img("2")] };
    expect(only(layoutTablesToColumns([oneRow(prose(59), two)]))).toMatchObject({
      kind: "columns",
      layout: "two_right_sidebar",
    });
  });

  it("the other blocks of the page pass through, in order, and only the table is replaced", () => {
    const before = p("avant");
    const after = p("apres");
    const out = layoutTablesToColumns([before, oneRow(prose(10), prose(10)), after]);
    expect(out).toHaveLength(3);
    expect(out[0]).toBe(before);
    expect(out[1]).toMatchObject({ kind: "columns" });
    expect(out[2]).toBe(after);
  });

  it("the input is not mutated (pure function)", () => {
    const input = [oneRow(prose(10), prose(10))];
    const snapshot = JSON.stringify(input);
    layoutTablesToColumns(input);
    expect(JSON.stringify(input)).toBe(snapshot);
  });
});

describe("what stays a table (the very same object comes back)", () => {
  const stays = (table: Table): void => {
    const out = layoutTablesToColumns([table]);
    expect(out).toHaveLength(1);
    expect(out[0]).toBe(table);
  };

  it("a form: 2 words and 1 word", () => stays(oneRow(prose(2), prose(1))));
  it("a form label | value where one cell is long prose and the other one word", () =>
    stays(oneRow(prose(40), prose(1))));
  it("boundary: 10 and 9 words stays a table", () => stays(oneRow(prose(10), prose(9))));
  it("boundary: 9 and 9 words stays a table", () => stays(oneRow(prose(9), prose(9))));
  it("boundary: 10 and 10 words becomes columns (not an overshoot of the previous cases)", () => {
    expect(only(layoutTablesToColumns([oneRow(prose(10), prose(10))]))).toMatchObject({ kind: "columns" });
  });
  it("three cells, the middle one at 9 words", () => stays(oneRow(prose(30), prose(9), prose(30))));
  it("prose + an EMPTY cell (a spacer)", () => stays(oneRow(prose(30), { blocks: [] })));
  it("prose + a cell with one empty paragraph", () => stays(oneRow(prose(30), { blocks: [p("")] })));
  it("pictures only", () => stays(oneRow(picture("1"), picture("2"))));
  it("a picture and an empty cell", () => stays(oneRow(picture("1"), { blocks: [] })));
  it("a picture and a few words (4 words + image is neither prose nor a picture cell)", () =>
    stays(oneRow(prose(30), { blocks: [p(words(4)), img()] })));
  it("colspan on a cell", () => stays(oneRow({ ...prose(20), colspan: 2 }, prose(20))));
  it("rowspan on a cell", () => stays(oneRow({ ...prose(20), rowspan: 2 }, prose(20))));
  it("a single cell", () => stays(oneRow(prose(30))));
  it("four cells", () => stays(oneRow(prose(30), prose(30), prose(30), prose(30))));
  it("two rows", () =>
    stays({ kind: "table", rows: [{ cells: [prose(30), prose(30)] }, { cells: [prose(30), prose(30)] }] }));
  it("a cell holding a table", () => {
    const inner = oneRow(prose(1), prose(1));
    stays(oneRow({ blocks: [p(words(20)), inner] }, prose(20)));
  });
});

describe("only the top level is looked at", () => {
  it("a layout-looking table INSIDE a cell of a data table is never converted", () => {
    const inner = oneRow(prose(15, "x"), prose(15, "y"));
    const outer: Table = { kind: "table", rows: [{ cells: [{ blocks: [inner] }] }, { cells: [prose(3)] }] };
    const out = layoutTablesToColumns([outer]);
    expect(out[0]).toBe(outer);
    expect(outer.rows[0]!.cells[0]!.blocks[0]).toBe(inner);
  });

  it("a layout-looking table inside a list item stays a table", () => {
    const inner = oneRow(prose(15), prose(15));
    const list: IRBlock = { kind: "list", ordered: false, items: [[p("item"), inner]] };
    const out = layoutTablesToColumns([list]);
    expect(out[0]).toBe(list);
    expect((out[0] as Extract<IRBlock, { kind: "list" }>).items[0]![1]).toBe(inner);
  });

  it("a layout-looking table inside a footnote block or a section stays a table", () => {
    const note: IRBlock = { kind: "footnote", id: "1", blocks: [oneRow(prose(15), prose(15))] };
    const section: IRBlock = { kind: "section", title: "S", blocks: [oneRow(prose(15), prose(15))] };
    const out = layoutTablesToColumns([note, section]);
    expect(out[0]).toBe(note);
    expect(out[1]).toBe(section);
  });
});

describe("other languages (the tokenizer decides)", () => {
  const de =
    "Die Geschäftsleitung hat in der letzten Sitzung beschlossen, dass alle Mitarbeiter bis Ende Januar geschult werden.";
  const fr =
    "Le comité de pilotage a décidé, lors de sa dernière réunion, que tous les agents seraient formés avant la fin du mois de janvier.";

  it("sanity: both sentences have at least 10 words, and a short form label has fewer", () => {
    expect(wordsOf(de).length).toBeGreaterThanOrEqual(10);
    expect(wordsOf(fr).length).toBeGreaterThanOrEqual(10);
  });

  it("German prose | French prose: columns", () => {
    expect(only(layoutTablesToColumns([oneRow({ blocks: [p(de)] }, { blocks: [p(fr)] })]))).toMatchObject({
      kind: "columns",
      layout: "two_equal",
    });
  });

  it("French form (label | value, 2 words and 3 words): table", () => {
    const table = oneRow({ blocks: [p("Nom prénom")] }, { blocks: [p("à compléter ici")] });
    expect(layoutTablesToColumns([table])[0]).toBe(table);
  });

  it("German form with accents (Straße | Hausnummer): table", () => {
    const table = oneRow({ blocks: [p("Straße")] }, { blocks: [p("Hausnummer 12")] });
    expect(layoutTablesToColumns([table])[0]).toBe(table);
  });

  it("a CJK cell of 10 or more tokens by the tokenizer counts as prose", () => {
    const cjk = "会議 の 議事録 を 作成 し ます 日本 語 で 書き ます";
    expect(wordsOf(cjk).length).toBeGreaterThanOrEqual(10);
    expect(only(layoutTablesToColumns([oneRow({ blocks: [p(cjk)] }, { blocks: [p(cjk)] })]))).toMatchObject({
      kind: "columns",
    });
  });

  it("a long CJK sentence written WITHOUT spaces is one token for the tokenizer: stays a table (the tokenizer decides, not the character count)", () => {
    const run = "国際会議における日本語の議事録を作成して関係者全員に配布することが決まりました";
    expect(wordsOf(run).length).toBeLessThan(10);
    const table = oneRow({ blocks: [p(run)] }, { blocks: [p(run)] });
    expect(layoutTablesToColumns([table])[0]).toBe(table);
  });
});

// Review of 24/10: the cells are classed BEFORE `images.resolve`, which may drop a picture (damaged part, image budget). A `columns` block
// whose picture cell lost its image would publish an empty sidebar; it goes back to the one-row table it was, as the reader gave it.
describe("columnsBackToTables — after the images are resolved", () => {
  it("🔴 a sidebar whose picture was dropped goes back to its one-row table", () => {
    const text = p(words(40, "t"));
    const emptied: IRBlock = { kind: "columns", layout: "two_left_sidebar", cells: [[], [text]] };
    // Failure value: the `columns` block, unchanged.
    expect(only(columnsBackToTables([emptied]))).toEqual(oneRow({ blocks: [] }, { blocks: [text] }));
  });

  it("🔴 three columns that lost one side picture go back to the table, the kept picture in place", () => {
    const text = p(words(40, "t"));
    const kept = img("b");
    const emptied: IRBlock = { kind: "columns", layout: "three_with_sidebars", cells: [[], [text], [kept]] };
    expect(only(columnsBackToTables([emptied]))).toEqual(
      oneRow({ blocks: [] }, { blocks: [text] }, { blocks: [kept] }),
    );
  });

  it("columns whose cells still qualify come back as the same object", () => {
    const whole: IRBlock = { kind: "columns", layout: "two_left_sidebar", cells: [[img()], [p(words(40))]] };
    expect(columnsBackToTables([whole])[0]).toBe(whole);
  });
});
