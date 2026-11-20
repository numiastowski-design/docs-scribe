/** Holds the one table of what we knowingly do not render: every line that promises a word boundary must really leave one,
 *  and must not decorate where there is nothing to separate.
 *  The test WALKS the table rather than naming its lines: a line added without the behaviour fails here, which is the whole
 *  point of having a table instead of a rule scattered over the readers.
 *  Called by: nothing (a test). */

import { describe, expect, it } from "vitest";

import { framedSkippedText, SKIPPED_ELEMENTS, type SkippedElement } from "./skipped-elements";

const KINDS = Object.keys(SKIPPED_ELEMENTS) as SkippedElement[];
const BOTH = { left: true, right: true };

describe("la table de ce qu'on sait ne pas savoir rendre (29/10/2026)", () => {
  it("n'est pas vide, et chaque ligne dit ce que la page reçoit", () => {
    expect(KINDS.length).toBeGreaterThan(0);
    for (const kind of KINDS) {
      const rule = SKIPPED_ELEMENTS[kind];
      expect(["its text", "nothing"], kind).toContain(rule.publishes);
      expect(rule.why.length, `${kind} : la ligne doit dire POURQUOI`).toBeGreaterThan(40);
    }
  });

  for (const kind of KINDS.filter((k) => SKIPPED_ELEMENTS[k].boundary === true)) {
    describe(kind, () => {
      it("sépare de ses deux voisins", () => {
        expect(framedSkippedText(kind, "u", BOTH)).toBe(" u ");
      });

      it("ne décore pas : rien à séparer, rien d'ajouté", () => {
        expect(framedSkippedText(kind, "u", { left: false, right: false })).toBe("u");
      });

      it("ne double pas une espace que la source donne déjà", () => {
        expect(framedSkippedText(kind, " u ", BOTH)).toBe(" u ");
        expect(framedSkippedText(kind, "u ", BOTH)).toBe(" u ");
      });

      it("laisse UNE frontière même quand on n'a rien pu lire", () => {
        // An equation holding only a space is dropped whole: skipping a thing is not the same as its never being there.
        expect(framedSkippedText(kind, "", BOTH)).toBe(" ");
        expect(framedSkippedText(kind, "", { left: false, right: false })).toBe("");
      });
    });
  }
});
