import { describe, expect, it, vi } from "vitest";

import type { OpcPackage, Relationship } from "./ooxml/package";
import { parseXml } from "./ooxml/xml";
import { readCachesWithoutModel } from "./smartart";

// The drawing caches no data model claims are given to the frames in the order of their part numbers; two parts with the
// SAME number are told apart by their names. That order must be the same on every machine: UTF-16 code units, never the
// order of the runtime's locale.
// Run: npx vitest run --project=main packages/docs-scribe/src/extractor/smartart.cache-order.test.ts

const NS =
  'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const cache = (word: string): string =>
  `<dsp:drawing ${NS}><dsp:spTree><dsp:sp><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>${word}</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`;

/** Four caches whose part names all carry the number 1: an upper-case initial, two lower-case ones, an accented one. */
const PARTS: Readonly<Record<string, string>> = {
  "d/zone1.xml": "zone",
  "d/été1.xml": "ete",
  "d/alpha1.xml": "alpha",
  "d/Zeta1.xml": "zeta",
};

/** Only `readXml` is called by `readCachesWithoutModel`. */
const fakePackage = { readXml: (path: string) => Promise.resolve(parseXml(cache(PARTS[path]!))) } as OpcPackage;

const relationships: ReadonlyMap<string, Relationship> = new Map(
  Object.keys(PARTS).map((target, rank) => [`rId${rank}`, { kind: "diagramDrawing", target, external: false }]),
);

describe("readCachesWithoutModel: caches with the same part number", () => {
  it("are ordered by UTF-16 code units, whatever the locale", async () => {
    const localeCompare = vi.spyOn(String.prototype, "localeCompare");
    try {
      const words = await readCachesWithoutModel(fakePackage, relationships, new Set(), 4);
      // Code units: "Z" (0x5A) < "a" (0x61) < "z" (0x7A) < "é" (0xE9).
      // Failure value under an English or French locale: alpha, ete, zeta, zone.
      expect(words).toEqual([["zeta"], ["alpha"], ["zone"], ["ete"]]);
      expect(localeCompare).not.toHaveBeenCalled();
    } finally {
      localeCompare.mockRestore();
    }
  });
});
