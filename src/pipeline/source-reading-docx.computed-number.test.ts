import { describe, expect, it } from "vitest";

import { computedNumberOf } from "./source-reading-docx";

describe("computedNumberOf: the label Word draws before a numbered paragraph, and nothing more (26/10/2026)", () => {
  it("a number or a one-word label keeps its whole prefix", () => {
    expect(computedNumberOf("4.7 Scope of the work", ["Scope of the work"])?.trim()).toBe("4.7");
    expect(computedNumberOf("Partie 1 - Le contexte", ["Le contexte"])?.trim()).toBe("Partie 1 -");
  });

  it("limit, written: a label word that carries punctuation (« Art. 1. ») loses its word — the shape cannot tell « Art. » from « page. »", () => {
    expect(computedNumberOf("Art. 1. Le texte", ["Le texte"])?.trim()).toBe("1.");
    expect(computedNumberOf("page. 3. Le texte", ["Le texte"])?.trim()).toBe("3.");
  });

  it("🔴 guard: three short words before the label are a sentence, not a label (the two-token limit)", () => {
    // FAILS WITH: « voir le point 1. » — a head of three short words taken whole.
    expect(computedNumberOf("voir le point 1. Le texte", ["Le texte"])?.trim()).toBe("1.");
  });

  it("🔴 a sentence glued before the label is not part of it (bench, docx-reel 08 p5)", () => {
    // FAILS WITH: « • individuals​1 in: o » — the bench then took « individuals » and « in » for numbers and counted their Word copies lost.
    expect(computedNumberOf("• individuals​1 in: o private schools", ["private schools"])?.trim()).toBe("o");
  });
  it("🔴 the first numbered paragraph that fits does not decide when a LONGER one also fits: « 7. » is the label of « 7. Annex B – Sample Policy » (shape of corpus 202 p1, +1 invented word, a number, since a51102b)", () => {
    // Two numbered paragraphs end the block: « Sample Policy » (listed first) and « Annex B – Sample Policy » (the heading itself). The first one cuts the
    // block at « 7. Annex B – », whose head « 7. Annex » is not two plain words, so since 26/10/2026 only « B – » is kept and the heading's number « 7 » is no
    // longer in the bench's number bag. Before a51102b the whole prefix was kept and hid it by accident.
    // FAILS WITH: expected 'B –' to be '7.' // Object.is equality
    expect(computedNumberOf("7. Annex B – Sample Policy", ["Sample Policy", "Annex B – Sample Policy"])?.trim()).toBe(
      "7.",
    );
  });
});
