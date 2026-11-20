import { describe, expect, it } from "vitest";

import { finiteNumber, intAttr } from "./ooxml-num";

// Decision n° 12 (12/10/2026): a blank numeric attribute is absent, as in the PowerPoint readers (never 0).
// Run: npx vitest run --project=main src/lib/ooxml-num.test.ts

describe("finiteNumber", () => {
  it("an empty or blank string is undefined, not 0", () => {
    expect(finiteNumber("")).toBeUndefined();
    expect(finiteNumber("  ")).toBeUndefined();
    expect(finiteNumber("\t\n")).toBeUndefined();
  });
  it("written numbers still read, zero included", () => {
    expect(finiteNumber("0")).toBe(0);
    expect(finiteNumber(" 12 ")).toBe(12);
    expect(finiteNumber("-3.5")).toBe(-3.5);
    expect(finiteNumber(undefined)).toBeUndefined();
    expect(finiteNumber("abc")).toBeUndefined();
    expect(finiteNumber("Infinity")).toBeUndefined();
  });
});

describe("intAttr", () => {
  it("unchanged: leading digits only", () => {
    expect(intAttr("12px")).toBe(12);
    expect(intAttr("x")).toBeUndefined();
  });
});
