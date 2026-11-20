/**
 * Word's truth for the wrapped URLs of the jvmalin guide: in `02-jvmalin-guide-api.docx` each of the two WebService URLs
 * is ONE word. The PDF reader's test on Word's own export of the same document (`pdf-hidden-repaints.fixtures.test.ts`)
 * requires both to be published whole; this half holds that the requirement is the document's, read by the Word reader.
 * NOT here: anything about the PDF export.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "./docx";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "../../fixtures/docx-reel");
const bytesOf = (name: string): Uint8Array => new Uint8Array(readFileSync(join(fixturesDir, name)));

const textOf = (blocks: readonly IRBlock[]): string =>
  blocks
    .map((b): string => {
      if (b.kind === "table")
        return b.rows.map((row) => row.cells.map((cell) => textOf(cell.blocks)).join(" ")).join(" ");
      if (b.kind === "list") return b.items.map((item) => textOf(item)).join(" ");
      if (b.kind === "paragraph" || b.kind === "heading") return b.inlines.map((i) => i.text).join("");
      return "";
    })
    .join(" ");

describe("02 (jvmalin guide API, Word) — the WebService URLs as Word writes them", () => {
  it("reads « schedules » in one piece in BOTH URLs of the Word document", async () => {
    const truth = textOf((await extractDocx(bytesOf("02-jvmalin-guide-api.docx"))).doc.blocks);
    expect(truth).toContain("{toStopAreaId}/schedules");
    expect(truth).toContain("{toStopPointId}/schedules");
  });
});
