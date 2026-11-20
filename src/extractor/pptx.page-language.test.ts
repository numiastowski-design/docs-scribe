/**
 * A presentation's tab labels follow the page language
 * (features/page-language.feature): `Slide N` or `Diapositive N`,
 * and the independent second reading's unit labels follow the SAME function —
 * a unit label is never compared with the page, but it is what the dashboard
 * names when words are missing.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import type { IRBlock } from "doc-ir";
import { readPptxConservationSource } from "../pipeline/source-reading-pptx";
import { extractPptx } from "./pptx";

const here = dirname(fileURLToPath(import.meta.url));
const deck = new Uint8Array(readFileSync(join(here, "../../fixtures/pptx/03-Presentation-Resultats-T1.pptx")));

const tabTitles = (blocks: readonly IRBlock[]): string[] =>
  blocks.flatMap((b) => (b.kind === "tabs" ? b.tabs.map((t) => t.title) : []));

describe("a presentation in the page language", () => {
  it("French: « Diapositive N » tabs", async () => {
    const { doc } = await extractPptx(deck, undefined, "fr");
    expect(tabTitles(doc.blocks)).toEqual([
      "Diapositive 1",
      "Diapositive 2",
      "Diapositive 3",
      "Diapositive 4",
      "Diapositive 5",
      "Diapositive 6",
    ]);
  });

  it("English by default", async () => {
    const { doc } = await extractPptx(deck);
    expect(tabTitles(doc.blocks)).toEqual(["Slide 1", "Slide 2", "Slide 3", "Slide 4", "Slide 5", "Slide 6"]);
  });

  it("the second reading's unit labels follow the same language", async () => {
    const fr = await readPptxConservationSource(deck, "fr");
    const en = await readPptxConservationSource(deck);
    if (!("units" in fr) || !("units" in en)) throw new Error("expected readings");
    expect(fr.units.map((u) => u.label)).toEqual([
      "Diapositive 1",
      "Diapositive 2",
      "Diapositive 3",
      "Diapositive 4",
      "Diapositive 5",
      "Diapositive 6",
    ]);
    expect(en.units.map((u) => u.label)).toEqual(["Slide 1", "Slide 2", "Slide 3", "Slide 4", "Slide 5", "Slide 6"]);
  });
});
