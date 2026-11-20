/**
 * The one costly case of `namespaces.test.ts`, on its own (10/11/2026). That file's support uses vi.doMock, which the
 * `isolate: false` lots do not honour, so it stays in « main »; this case imports only the tree, so it can run in
 * « pdf-serial », away from the contention that once took it to 21.8 s.
 * A scope copied per element (as `ooxml/xml.ts#canonicalize` does) would cost levels × bindings here: 500 × 1 000
 * declarations below 500 levels, then 20 000 elements each using one of them. A scope copied at each declaring element
 * costs about 1.25 × 10^8 entries (≈ 50 s), at every element about 10^10; ours ≈ 6 s under vitest, mostly the tokenizer.
 * Cap 20 s (counter-review of 25/10/2026 of the namespace scopes: 10 s flaked under a loaded full check; still 2.5× under the cheaper failing
 * mutant).
 * Called by: nothing (a test).
 */
import { describe, expect, it } from "vitest";

import { buildTree } from "./tree";

describe("xml namespaces — the work is bounded by what is read, never by the scopes in force", () => {
  it("500 levels declaring 1 000 prefixes each, then 20 000 uses: linear", () => {
    const levels = 500;
    const perLevel = 1_000;
    const declarations = (level: number): string =>
      Array.from({ length: perLevel }, (_, k) => ` xmlns:p${level}_${k}="u${k}"`).join("");
    const open = Array.from({ length: levels }, (_, level) => `<e${declarations(level)}>`).join("");
    const uses = Array.from(
      { length: 20_000 },
      (_, k) => `<p${k % levels}_${k % perLevel}:x p0_${k % perLevel}:y="1"/>`,
    ).join("");
    const source = open + uses + "</e>".repeat(levels);
    const started = performance.now();
    expect(buildTree(source, "document")).toHaveLength(1);
    expect(performance.now() - started).toBeLessThan(20_000);
  }, 60_000);
});
