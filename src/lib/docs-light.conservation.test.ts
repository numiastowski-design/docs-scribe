/**
 * The word check of the declared formats, through the published light door (`docs-scribe/light`): `pageWords`,
 * `missingWordGroups` and `conservationFromMissingWordGroups`, as a client of the package calls them — a source
 * written by hand, a page given as TEXT block by block.
 * The publisher's application holds the same rules on ITS page (the markup of a Confluence page turned into these
 * blocks, the numeric cells found in its tables); nothing of that is read here.
 * WHAT THE CLIENT COMPOSES, written on purpose: the record is `conservationFromMissingWordGroups(missingWordGroups(…))`.
 */
import { describe, expect, it } from "vitest";

import {
  conservationFromMissingWordGroups,
  missingWordGroups,
  pageWords,
  type DeclaredFormatConservationSource,
  type PublishedPage,
} from "./docs-light";

const page = (...blocks: string[]): PublishedPage => ({ blocks });

const source = (
  words: string[],
  opts: { deliberateRemovals?: string[]; excludedUnits?: number[] } = {},
): DeclaredFormatConservationSource => ({
  units: [{ label: "", tokens: words.map((text) => ({ text })) }],
  deliberateRemovals: opts.deliberateRemovals ?? [],
  excludedUnits: new Set(opts.excludedUnits ?? []),
});

const conservationOf = (
  src: DeclaredFormatConservationSource,
  published: PublishedPage,
): ReturnType<typeof conservationFromMissingWordGroups> =>
  conservationFromMissingWordGroups(missingWordGroups(src, published));

const OK = { decision: "ok", lostLetters: 0, gluedWords: 0, missingWords: 0, units: [] };

describe("pageWords — the published page's own words", () => {
  it("tokenises the text of the blocks like the source side: lower case, no punctuation", () => {
    expect(pageWords(page("Le chat dort & rêve."))).toEqual(["le", "chat", "dort", "rêve"]);
  });

  it("a block boundary ends a word: two blocks never weld their last and first words", () => {
    // Failure value: ["premiersecond"], the blocks joined with nothing between them.
    expect(pageWords(page("premier", "second"))).toEqual(["premier", "second"]);
  });

  it("an empty page has no word", () => {
    expect(pageWords(page())).toEqual([]);
    expect(pageWords(page(""))).toEqual([]);
  });

  it("U+200B is not a break: « 4 » + zero-width space + « ème » is the one word « 4ème »", () => {
    // Failure value: ["4", "ème"], one missing source word on every ordinal a page writes with a superscript.
    expect(pageWords(page("4​ème end"))).toEqual(["4ème", "end"]);
  });
});

describe("missingWordGroups then conservationFromMissingWordGroups — missing words only, no repair", () => {
  it("every source word found on the page: ok, zero missing", () => {
    expect(conservationOf(source(["chat", "dort"]), page("Le chat dort."))).toEqual(OK);
  });

  it("a source word the page never carries is reported by count and cause, never by name", () => {
    const record = conservationOf(source(["chat", "dort"]), page("Le chat."));
    expect(record).toEqual({
      decision: "missing-words",
      lostLetters: 0,
      gluedWords: 0,
      missingWords: 1,
      units: [],
      causeCounts: [{ cause: "text", count: 1 }],
    });
    expect(JSON.stringify(record)).not.toContain("dort");
  });

  it("missingWordGroups gives the words themselves, grouped by cause in first-seen order", () => {
    const withCauses: DeclaredFormatConservationSource = {
      units: [
        {
          label: "",
          tokens: [
            { text: "chat" },
            { text: "confidentiel", cause: "content-control" },
            { text: "billet", cause: "footnote" },
            { text: "salaire", cause: "content-control" },
          ],
        },
      ],
      deliberateRemovals: [],
      excludedUnits: new Set(),
    };
    const groups = missingWordGroups(withCauses, page("Le chat."));
    expect(groups).toEqual([
      { cause: "content-control", words: ["confidentiel", "salaire"] },
      { cause: "footnote", words: ["billet"] },
    ]);
    const record = conservationFromMissingWordGroups(groups);
    expect(record.missingWords).toBe(3);
    expect(record.causeCounts).toEqual([
      { cause: "content-control", count: 2 },
      { cause: "footnote", count: 1 },
    ]);
    expect(JSON.stringify(record)).not.toMatch(/confidentiel|salaire|billet/);
  });

  it("a source word the page SPLITS in two is not lost when both halves are on the page", () => {
    // A heading and the next paragraph held in one source paragraph read « competitionthere »; the page publishes a
    // heading then a paragraph. Not a letter is missing: reported apart (`split-by-page`), never counted as a loss.
    const groups = missingWordGroups(source(["competitionthere"]), page("…Competition", "There are no changes."));
    expect(groups).toEqual([{ cause: "split-by-page", words: ["competitionthere"] }]);
    expect(conservationFromMissingWordGroups(groups)).toMatchObject({
      decision: "ok",
      missingWords: 0,
      causeCounts: [{ cause: "split-by-page", count: 1 }],
    });
  });

  it("… but only when BOTH halves are on the page: one half alone leaves the word lost", () => {
    expect(conservationOf(source(["competitionthere"]), page("There are no changes.")).missingWords).toBe(1);
  });

  it("counts a word missing as many times as it is short on the page", () => {
    expect(conservationOf(source(["chat", "chat", "chat"]), page("Le chat dort.")).missingWords).toBe(2);
  });

  it("a page word absent from the source is NEVER reported (missing only, never extra)", () => {
    expect(conservationOf(source(["chat"]), page("Le grand chat noir dort."))).toEqual(OK);
  });

  it("an excluded unit's words never count as missing", () => {
    expect(conservationOf(source(["chat", "dort"], { excludedUnits: [0] }), page("rien à voir ici"))).toEqual(OK);
  });

  it("a declared removal absorbs a leftover source word instead of counting it missing", () => {
    const withRemoval = source(["chat", "dort", "confidentiel"], { deliberateRemovals: ["confidentiel"] });
    expect(conservationOf(withRemoval, page("Le chat dort."))).toEqual(OK);
    // Control: without the declared removal the same word is missing.
    expect(conservationOf(source(["chat", "dort", "confidentiel"]), page("Le chat dort.")).missingWords).toBe(1);
  });

  it("however many words are missing, the count stays exact and the record holds none of them", () => {
    const many = Array.from({ length: 40 }, (_, index) => `mot${index}`);
    const record = conservationOf(source(many), page(""));
    expect(record.missingWords).toBe(many.length);
    expect(record.causeCounts).toEqual([{ cause: "text", count: many.length }]);
    expect(JSON.stringify(record)).not.toContain("mot0");
  });
});

describe("missingWordGroups — the numeric cells the caller found missing join the same cause groups", () => {
  it("a missing numeric cell is reported beside a missing word, each in its own cause group, by its reference", () => {
    const groups = missingWordGroups(
      source(["chat", "dort"]),
      page("chat"),
      [],
      [],
      [{ ref: "Budget!B1", cause: "number" }],
    );
    expect(groups).toEqual([
      { cause: "text", words: ["dort"] },
      { cause: "number", words: ["Budget!B1"] },
    ]);
  });

  it("no cell given, none reported", () => {
    expect(missingWordGroups(source(["chat"]), page("chat"))).toEqual([]);
  });
});

describe("conservationFromMissingWordGroups — the record holds no word, whatever its length", () => {
  it("a 50 000-character « word » and a short one: counts only, nothing stored", () => {
    const blob = "A".repeat(50_000);
    const record = conservationFromMissingWordGroups([{ cause: "text", words: [blob, "Montant"] }]);
    expect(record.causeCounts).toEqual([{ cause: "text", count: 2 }]);
    expect(record.missingWords).toBe(2);
    expect(JSON.stringify(record)).not.toMatch(/AAAA|Montant/);
  });

  it("no group: ok, and no cause count at all", () => {
    expect(conservationFromMissingWordGroups([])).toEqual(OK);
  });
});

// The third argument of `missingWordGroups` is the words of the document's diagram blocks (`diagramTextWords`, from
// the reading door), the fourth the SmartArt words its picture does not show (`undrawnDiagramWords`).
describe("missingWordGroups — the words of a diagram are not lost", () => {
  it("a source word that lives only in the diagram text is not missing: no group, decision ok", () => {
    const groups = missingWordGroups(source(["chat", "zebrafish"]), page("Le chat."), ["zebrafish"]);
    expect(groups).toEqual([]);
    expect(conservationFromMissingWordGroups(groups)).toMatchObject({ decision: "ok", missingWords: 0 });
  });

  it("a word absent from the page AND from the diagram text is still lost, cause « text »", () => {
    const groups = missingWordGroups(source(["chat", "zebrafish", "dort"]), page("Le chat."), ["zebrafish"]);
    expect(groups).toEqual([{ cause: "text", words: ["dort"] }]);
  });

  it("one occurrence of the diagram text excuses one occurrence of the word, not all of them", () => {
    const groups = missingWordGroups(source(["zebrafish", "zebrafish"]), page("Rien."), ["zebrafish"]);
    expect(groups).toEqual([{ cause: "text", words: ["zebrafish"] }]);
  });

  it("a SmartArt word its picture does not show is still lost, « smartart-not-drawn », though the diagram text holds it", () => {
    const groups = missingWordGroups(source(["alpha", "beta"]), page("Rien."), ["alpha", "beta"], ["alpha"]);
    expect(groups).toEqual([{ cause: "smartart-not-drawn", words: ["alpha"] }]);
    expect(conservationFromMissingWordGroups(groups)).toMatchObject({
      decision: "missing-words",
      missingWords: 1,
      causeCounts: [{ cause: "smartart-not-drawn", count: 1 }],
    });
  });

  it("only a plain « text » word is excused: a word already given another cause keeps it", () => {
    const withCause: DeclaredFormatConservationSource = {
      units: [{ label: "", tokens: [{ text: "zebrafish", cause: "footnote" }] }],
      deliberateRemovals: [],
      excludedUnits: new Set(),
    };
    expect(missingWordGroups(withCause, page("Rien."), ["zebrafish"])).toEqual([
      { cause: "footnote", words: ["zebrafish"] },
    ]);
  });
});
