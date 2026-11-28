import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import type { IRBlock, IRDoc } from "doc-ir";
import { extractDocx } from "./docx";

// Two real Word documents read through the published entry (`extractDocx`): headings, tables and images, checked
// against each file's raw `document.xml` (heading styles, `w:tbl`, `w:drawing` counts, in document order) when the
// cases were written. The reader's own cases on the same files (`docx-reader/docx-reader.test.ts`) check the cells
// and marks; these check what a client of the package gets. Came from the publisher's application on 28/11/2026.
const here = dirname(fileURLToPath(import.meta.url));
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures/docx", name)));

/** How many blocks of `kind` sit inside `block` (the block itself not counted). */
const nestedCount = (block: IRBlock, kind: IRBlock["kind"]): number => {
  if ("blocks" in block) return countByKind(block.blocks, kind);
  if (block.kind === "list") return block.items.reduce((sum, item) => sum + countByKind(item, kind), 0);
  if (block.kind === "table")
    return block.rows.reduce(
      (sum, row) => sum + row.cells.reduce((cellSum, cell) => cellSum + countByKind(cell.blocks, kind), 0),
      0,
    );
  return 0;
};

const countByKind = (blocks: IRBlock[], kind: IRBlock["kind"]): number =>
  blocks.reduce((total, block) => total + (block.kind === kind ? 1 : 0) + nestedCount(block, kind), 0);

const headingLevels = (doc: IRDoc): number[] =>
  doc.blocks
    .filter((block): block is Extract<IRBlock, { kind: "heading" }> => block.kind === "heading")
    .map((heading) => heading.level);

describe("extractDocx (real Word documents)", () => {
  it("01-Guide-Onboarding-Atria.docx: headings, tables and images match the source XML", async () => {
    const { doc, images } = await extractDocx(fixture("01-Guide-Onboarding-Atria.docx"));

    expect(headingLevels(doc)).toEqual([1, 2, 2, 3, 3, 2, 3, 3, 2, 2]);
    expect(countByKind(doc.blocks, "table")).toBe(2);
    expect(countByKind(doc.blocks, "image")).toBe(2);
    // Two <w:drawing> in the XML, two distinct raw images collected.
    expect(images).toHaveLength(2);
    for (const image of images) {
      expect(image.mimeType).toBe("image/png");
      expect(image.sha256).toMatch(/^[0-9a-f]{64}$/);
      expect(image.bytes.length).toBeGreaterThan(0);
    }
    // Every image block references bytes actually collected — no dangling sha256.
    const shas = new Set(images.map((image) => image.sha256));
    for (const block of doc.blocks) {
      if (block.kind === "image") expect(shas.has(block.ref.sha256)).toBe(true);
    }
  });

  it("02-Procedure-Support-IT.docx: shorter doc, one image, no tables", async () => {
    const { doc, images } = await extractDocx(fixture("02-Procedure-Support-IT.docx"));

    expect(headingLevels(doc)).toEqual([1, 2, 2]);
    expect(countByKind(doc.blocks, "table")).toBe(0);
    expect(countByKind(doc.blocks, "image")).toBe(1);
    expect(images).toHaveLength(1);
  });

  it("is deterministic: extracting the same bytes twice yields the same IRDoc", async () => {
    const bytes = fixture("01-Guide-Onboarding-Atria.docx");
    const first = await extractDocx(bytes);
    const second = await extractDocx(bytes);
    expect(second.doc).toEqual(first.doc);
    const bySha = (left: string, right: string): number => (left < right ? -1 : Number(left > right));
    expect(second.images.map((image) => image.sha256).sort(bySha)).toEqual(
      first.images.map((image) => image.sha256).sort(bySha),
    );
  });
});
