import { describe, expect, it } from "vitest";

import { externalLinkTarget } from "./external-link-target";

describe("externalLinkTarget — only an External relationship's own target may become an href", () => {
  it("keeps the target of an external relationship", () => {
    expect(externalLinkTarget({ external: true, target: "https://example.org/a" })).toBe("https://example.org/a");
  });

  it("drops the target of an internal relationship, and a link that is not there", () => {
    expect(externalLinkTarget({ external: false, target: "https://example.org/a" })).toBeUndefined();
    expect(externalLinkTarget(undefined)).toBeUndefined();
  });

  it("never takes one of our internal forms from a file, whatever its case", () => {
    expect(externalLinkTarget({ external: true, target: "drive:abc" })).toBeUndefined();
    expect(externalLinkTarget({ external: true, target: "DRIVE:abc" })).toBeUndefined();
    expect(externalLinkTarget({ external: true, target: "#anchor" })).toBeUndefined();
  });

  it("leaves the scheme to the renderer: a target of another scheme passes here", () => {
    expect(externalLinkTarget({ external: true, target: "javascript:alert(1)" })).toBe("javascript:alert(1)");
  });
});
