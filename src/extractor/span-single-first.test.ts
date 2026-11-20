/** The shared colspan rule, pinned directly (it is exported for exactly this): the hardening found by the 27/10/2026 evening review —
 *  the rule ASSUMED every rowspan-free row sums to the table's width (true of every current builder, `sheet-grid.ts#tableOfKept` emits
 *  one cell per kept column) without asserting it. A RAGGED row (shorter than the table) must ABSTAIN, never widen itself to the
 *  table's width: a discriminant the rule cannot read (why is the row short?) is a decision it must not take. */
import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { hasContent } from "doc-ir/internal";
import { spanSingleFirstRows } from "./span-single-first";

type Table = Extract<IRBlock, { kind: "table" }>;
const cell = (
  text: string,
  over: Partial<Table["rows"][number]["cells"][number]> = {},
): Table["rows"][number]["cells"][number] => ({
  blocks: text === "" ? [] : [{ kind: "paragraph", inlines: [{ text }] }],
  ...over,
});
const table = (rows: Table["rows"]): Table => ({ kind: "table", rows });

describe("spanSingleFirstRows — direct pins", () => {
  it("spans a single-first row whose cells sum to the table's width", () => {
    const out = spanSingleFirstRows(
      table([{ cells: [cell("Title"), cell(""), cell("")] }, { cells: [cell("a"), cell("b"), cell("c")] }]),
    );
    expect(out.rows[0]!.cells).toHaveLength(1);
    expect(out.rows[0]!.cells[0]!.colspan).toBe(3);
  });

  it("ABSTAINS on a RAGGED single-first row (its cells sum under the table's width): the row is left exactly as it was, never widened", () => {
    // FAILS WITH (before the hardening): the two-cell row is replaced by one cell of colspan 3 — the rule widened a row it did not understand.
    const ragged = { cells: [cell("Note"), cell("")] };
    const out = spanSingleFirstRows(table([{ cells: [cell("a"), cell("b"), cell("c")] }, ragged]));
    expect(out.rows[1]!.cells).toHaveLength(2);
    expect(out.rows[1]!.cells[0]!.colspan).toBeUndefined();
  });

  it("a cell holding ONLY a list keeps its row: content without paragraph text is content (revue 27/10/2026, zero lost words)", () => {
    // FAILS WITH (before the fix): the list cell reads as empty text, the row collapses to one cell and the list's words vanish.
    const listCell: Table["rows"][number]["cells"][number] = {
      blocks: [
        {
          kind: "list",
          ordered: false,
          items: [[{ kind: "paragraph", inlines: [{ text: "point" }] }]],
        },
      ],
    };
    const out = spanSingleFirstRows(
      table([{ cells: [cell("Title"), listCell, cell("")] }, { cells: [cell("a"), cell("b"), cell("c")] }]),
    );
    expect(out.rows[0]!.cells).toHaveLength(3);
  });

  it("a cell holding ONLY an image keeps its row too", () => {
    const imageCell: Table["rows"][number]["cells"][number] = {
      blocks: [{ kind: "image", ref: "r1" } as unknown as IRBlock],
    };
    const out = spanSingleFirstRows(table([{ cells: [cell("Title"), imageCell] }, { cells: [cell("a"), cell("b")] }]));
    expect(out.rows[0]!.cells).toHaveLength(2);
  });

  it("🔴 a cell whose only inline is an ANCHOR target (a Word bookmark, empty text) keeps its row: dropping it would break every #name link (revue 27/10/2026, tour 8)", () => {
    // FAILS WITH (before the fix): holdsContent read the inline TEXT only, the bookmark cell counted as empty, the row
    // collapsed and the anchor macro was never rendered — the internal links pointing at it dangle.
    const bookmarkCell: Table["rows"][number]["cells"][number] = {
      blocks: [{ kind: "paragraph", inlines: [{ text: "", anchor: "chapitre-2" }] }],
    };
    const out = spanSingleFirstRows(
      table([{ cells: [cell("Title"), bookmarkCell] }, { cells: [cell("a"), cell("b")] }]),
    );
    expect(out.rows[0]!.cells).toHaveLength(2);
  });

  it("🔴 a cell whose only inline carries an INLINE IMAGE (empty text) keeps its row too", () => {
    const pictureCell: Table["rows"][number]["cells"][number] = {
      blocks: [{ kind: "paragraph", inlines: [{ text: "", image: { ref: "r9" } }] } as unknown as IRBlock],
    };
    const out = spanSingleFirstRows(
      table([{ cells: [cell("Title"), pictureCell] }, { cells: [cell("a"), cell("b")] }]),
    );
    expect(out.rows[0]!.cells).toHaveLength(2);
  });
  it("still spans a single-first row whose width comes from a horizontal merge (anchor colspan + trailing empty)", () => {
    const out = spanSingleFirstRows(
      table([
        { cells: [cell("Merged note", { colspan: 4 }), cell("")] },
        { cells: [cell("a"), cell("b"), cell("c"), cell("d"), cell("e")] },
      ]),
    );
    expect(out.rows[0]!.cells).toHaveLength(1);
    expect(out.rows[0]!.cells[0]!.colspan).toBe(5);
  });
});

describe("the rule's notion of CONTENT and the renderer's are the same (revue 27/10/2026, tour 9)", () => {
  // The rule writes its own copy (the library may take no new edge to the app; step 2 of the library plan deletes it).
  // Pinned here, where a test may import the app: a cell the renderer would keep must never be collapsed over.
  const cases: { what: string; blocks: IRBlock[] }[] = [
    { what: "empty paragraph", blocks: [{ kind: "paragraph", inlines: [{ text: "  " }] }] },
    { what: "text", blocks: [{ kind: "paragraph", inlines: [{ text: "x" }] }] },
    { what: "anchor only", blocks: [{ kind: "paragraph", inlines: [{ text: "", anchor: "a" }] }] },
    {
      what: "inline image only",
      blocks: [{ kind: "paragraph", inlines: [{ text: "", image: { ref: "r" } }] } as unknown as IRBlock],
    },
    {
      what: "empty list",
      blocks: [
        {
          kind: "list",
          ordered: false,
          items: [[{ kind: "paragraph", inlines: [{ text: "" }] }]],
        },
      ],
    },
    {
      what: "list with a word",
      blocks: [
        {
          kind: "list",
          ordered: false,
          items: [[{ kind: "paragraph", inlines: [{ text: "p" }] }]],
        },
      ],
    },
    { what: "nested empty table", blocks: [{ kind: "table", rows: [{ cells: [{ blocks: [] }] }] }] },
    {
      what: "nested table with a word",
      blocks: [
        {
          kind: "table",
          rows: [{ cells: [{ blocks: [{ kind: "paragraph", inlines: [{ text: "n" }] }] }] }],
        },
      ],
    },
    { what: "image block", blocks: [{ kind: "image", ref: "r1" } as unknown as IRBlock] },
    { what: "no block at all", blocks: [] },
  ];

  for (const { what, blocks } of cases) {
    it(`${what}: the rule collapses over it exactly when the renderer drops it`, () => {
      const keptByRenderer = hasContent(blocks);
      const out = spanSingleFirstRows(
        table([{ cells: [cell("Title"), { blocks }] }, { cells: [cell("a"), cell("b")] }]),
      );
      const collapsed = out.rows[0]!.cells.length === 1;
      expect(collapsed).toBe(!keptByRenderer);
    });
  }
});
