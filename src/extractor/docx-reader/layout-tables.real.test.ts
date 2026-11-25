import { readFileSync } from "node:fs";

import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "../docx";

// A real Word file (fixtures/docx-layout-tables/SOURCES.md): 097 is a form whose one-row tables have 0 to 4 words per cell.
// The words never change.
// The READER alone is held here; the published page and its controls on the same file are held by the
// publisher's application, not by this package.
vi.setConfig({ testTimeout: 120_000 });

const DIR = "fixtures/docx-layout-tables";
const load = (name: string): Uint8Array => new Uint8Array(readFileSync(`${DIR}/${name}`));
const F097 = "097-mlit-38bf.docx";

type Columns = Extract<IRBlock, { kind: "columns" }>;

const topColumns = (blocks: readonly IRBlock[]): Columns[] => blocks.filter((b): b is Columns => b.kind === "columns");

describe("097: a form with 0 to 4 words per cell stays tables", () => {
  it("no `columns` block at all, and its one-row tables are still there", async () => {
    const { doc } = await extractDocx(load(F097));
    expect(topColumns(doc.blocks)).toHaveLength(0);
    const oneRow = doc.blocks.filter((b) => b.kind === "table" && b.rows.length === 1 && b.rows[0]!.cells.length >= 2);
    expect(oneRow.length).toBeGreaterThan(30);
  });
});
