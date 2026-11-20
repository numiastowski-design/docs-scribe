import { describe, expect, it } from "vitest";
import { parseXml, qualifiedNameOf } from "./xml";

// A namespace URI equal to an Object.prototype key (« toString », « constructor ») is no namespace the reader knows:
// the producer's own prefix stays (review of 07/10/2026, cleanup).
describe("namespace URIs that are Object.prototype keys", () => {
  it("leaves the prefix of a namespace « toString » or « constructor » as written", () => {
    for (const uri of ["toString", "constructor"]) {
      const root = parseXml(`<x:doc xmlns:x="${uri}"><x:p/></x:doc>`)!;
      expect(qualifiedNameOf(root)).toBe("x:doc");
    }
  });
});
