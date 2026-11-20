import { describe, expect, it } from "vitest";

import { isHiddenFlag, isXsdTrue } from "./ooxml-bool";

// Decision n° 7 (12/10/2026): only the « hidden » flag is tolerant; the other booleans stay at the norm.
// Run: npx vitest run --project=main src/lib/ooxml-bool.test.ts

describe("isHiddenFlag", () => {
  it.each(["1", "true", "TRUE", " 1 ", " True\n"])('"%s" is hidden', (v) => expect(isHiddenFlag(v)).toBe(true));
  it.each(["0", "false", "", " ", "yes", "on", undefined])("%j is not hidden", (v) =>
    expect(isHiddenFlag(v)).toBe(false),
  );
});

describe("isXsdTrue stays strict", () => {
  it.each(["TRUE", " 1 "])('"%s" is not xsd true', (v) => expect(isXsdTrue(v)).toBe(false));
});
