import { describe, expect, it } from "vitest";

import { extractGoogleSheet } from "./google-sheet";
import type { GoogleSheetMeta } from "./google-sheet";

// Real API responses (spreadsheets.get + values:batchGet), captured via a
// throwaway probe against the actual Sheets API (C1 session log, STATUS.md)
// for the Google Sheets twin of fixtures/xlsx/04-Effectifs-Atria.xlsx (same
// corpus item, published as a native Sheet — same tab names, same merge at
// A9:D9 on "2026", same hidden "Archive-Interne" tab).
const META: GoogleSheetMeta = {
  sheets: [
    { properties: { title: "2025", index: 0 } },
    {
      properties: { title: "2026", index: 1 },
      merges: [{ startRowIndex: 8, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 4 }],
    },
    { properties: { title: "Grand-Onglet", index: 2 } },
    { properties: { title: "Archive-Interne", index: 3, hidden: true } },
  ],
};

const VALUES = new Map<string, string[][]>([
  [
    "2025",
    [
      ["Service", "Effectif janvier", "Effectif décembre", "Variation"],
      ["Direction Commerciale", "42", "48", "6"],
      ["Direction Opérations", "96", "101", "5"],
      ["Direction Finance", "24", "24", "0"],
      ["Support Informatique", "18", "22", "4"],
      ["Ressources Humaines", "11", "12", "1"],
      ["Total", "191", "207", "16"],
    ],
  ],
  [
    "2026",
    [
      ["Service", "Effectif janvier", "Effectif décembre", "Variation", "Contrôle"],
      ["Direction Commerciale", "48", "51", "3", "#DIV/0!"],
      ["Direction Opérations", "101", "104", "3"],
      ["Direction Finance", "24", "25", "1"],
      ["Support Informatique", "22", "24", "2"],
      ["Ressources Humaines", "12", "13", "1"],
      ["Total", "207", "217", "10"],
      [],
      ["Note : effectifs mesurés au dernier jour de chaque mois."],
    ],
  ],
]);

describe("extractGoogleSheet (real API data)", () => {
  it("produces one TAB per visible sheet, in index order, dropping the hidden one (tabs shape since étape 11)", async () => {
    const { doc } = await extractGoogleSheet(META, VALUES);
    const tabs = doc.blocks[0];
    if (tabs?.kind !== "tabs") throw new Error("expected tabs");
    expect(tabs.tabs.map((t) => t.title)).toEqual(["2025", "2026", "Grand-Onglet"]);
  });

  it("'2025': 7 rows x 4 cols, no merges, a formula-error cell reproduced verbatim would be here if present", async () => {
    const { doc } = await extractGoogleSheet(META, VALUES);
    const tabsBlock = doc.blocks[0];
    if (tabsBlock?.kind !== "tabs") throw new Error("expected tabs");
    const table = tabsBlock.tabs[0]!.blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    expect(table.rows).toHaveLength(7);
    expect(table.rows[0]?.cells).toHaveLength(4);
  });

  it("'2026': ragged rows padded with empty cells, #DIV/0! reproduced as plain text, merge at row 8 (A9:D9) as colspan 4", async () => {
    const { doc } = await extractGoogleSheet(META, VALUES);
    const tabsBlock = doc.blocks[0];
    if (tabsBlock?.kind !== "tabs") throw new Error("expected tabs");
    const table = tabsBlock.tabs[1]!.blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");

    // The blank row (index 7) drops since the revue HIGH du 29/09 (constat
    // 5) — the same retreat as the .xlsx twin's since étape 2; the note row
    // follows.
    expect(table.rows).toHaveLength(8);
    expect(table.rows[1]?.cells[4]?.blocks[0]).toEqual({ kind: "paragraph", inlines: [{ text: "#DIV/0!" }] });
    // A row with fewer values than the header (e.g. row 2, no "Contrôle") gets
    // an empty (not missing) trailing cell — same column count as every row.
    expect(table.rows[2]?.cells).toHaveLength(5);
    expect(table.rows[2]?.cells[4]?.blocks).toEqual([]);

    const noteRow = table.rows[7];
    // 27/10/2026: the A:D merge (colspan 4) + the empty column E (2 cells) became ONE cell over the full width: a row whose only
    // non-empty cell is the first one spans every column (`span-single-first.ts#spanSingleFirstRows`, shared with xlsx). Was 2 cells, colspan 4.
    expect(noteRow?.cells).toHaveLength(1);
    expect(noteRow?.cells[0]?.colspan).toBe(5);
    expect(noteRow?.cells[0]?.rowspan).toBeUndefined();
    expect(noteRow?.cells[0]?.blocks[0]).toEqual({
      kind: "paragraph",
      inlines: [{ text: "Note : effectifs mesurés au dernier jour de chaque mois." }],
    });
  });

  it("is deterministic and synchronous (no I/O — a pure fold over already-fetched data)", async () => {
    const a = await extractGoogleSheet(META, VALUES);
    const b = await extractGoogleSheet(META, VALUES);
    expect(a).toEqual(b);
  });

  // STATUS.md 31/08/2026 (Gemini consultation, data-exposure finding, same
  // as xlsx.ts's own): a hidden row/column within an otherwise-visible tab
  // must be ignored. Field names (`rowMetadata[].hiddenByUser`,
  // `columnMetadata[].hiddenByUser`) confirmed live against the real Sheets
  // API (c5-parent-probe.ts, sheets-hidden-field-probe), not assumed.
  it("ignores a hidden row and a hidden column, same treatment as a hidden tab", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [
        {
          properties: { title: "Payroll", index: 0 },
          rowMetadata: [{}, { hiddenByUser: true }, {}],
          columnMetadata: [{}, { hiddenByUser: true }, {}],
        },
      ],
    };
    const values = new Map([
      [
        "Payroll",
        [
          ["a1", "b1", "c1"],
          ["a2-secret", "b2-secret", "c2-secret"],
          ["a3", "b3", "c3"],
        ],
      ],
    ]);
    const { doc } = await extractGoogleSheet(meta, values);
    const tabsBlock = doc.blocks[0];
    if (tabsBlock?.kind !== "tabs") throw new Error("expected tabs");
    const table = tabsBlock.tabs[0]!.blocks[0];
    if (table?.kind !== "table") throw new Error("expected table");
    const allText = table.rows.flatMap((r) =>
      r.cells.flatMap((c) => c.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text) : []))),
    );
    expect(allText).not.toContain("a2-secret");
    expect(allText).not.toContain("b2-secret");
    expect(allText).not.toContain("c2-secret");
    expect(allText).not.toContain("b1");
    expect(allText).not.toContain("b3");
    expect(allText).toEqual(["a1", "c1", "a3", "c3"]);
  });

  it("an unhidden tab with no values entry becomes an empty table, not a crash", async () => {
    const { doc } = await extractGoogleSheet(META, new Map());
    const tabsBlock = doc.blocks[0];
    if (tabsBlock?.kind !== "tabs") throw new Error("expected tabs");
    expect(tabsBlock.tabs[2]!.blocks[0]).toEqual({ kind: "table", rows: [] });
  });
});

// Cell links of a Google Sheet (07/10/2026): the same as a .xlsx cell link — the cell's text carries an href; an
// internal link (to another tab, « #gid=… ») stays plain text, as a .xlsx's internal link does.
describe("extractGoogleSheet — cell links", () => {
  it("🔴 a linked cell's text carries the link; an internal link does not", async () => {
    const meta: GoogleSheetMeta = {
      sheets: [
        {
          properties: { title: "Liens", index: 0 },
          links: [
            { row: 1, col: 1, uri: "https://miksumia.com/" },
            { row: 2, col: 1, uri: "#gid=0" },
          ],
        },
      ],
    };
    const values = new Map([
      [
        "Liens",
        [
          ["Ce qui est testé", "Cellule"],
          ["Lien", "Site Miksumia"],
          ["Interne", "Onglet"],
        ],
      ],
    ]);
    const { doc } = await extractGoogleSheet(meta, values);
    const inlines = JSON.stringify(doc.blocks);
    expect(inlines).toContain('{"text":"Site Miksumia","href":"https://miksumia.com/"}');
    expect(inlines).toContain('{"text":"Onglet"}');
  });
});
