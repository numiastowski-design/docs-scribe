/**
 * THE SECOND READINGS ARE INDEPENDENT OF THE EXTRACTORS, held by a test. A conservation source's second reading
 * (`pipeline/source-reading-*.ts`) never reads a format through its OWN extractor (`extractor/docx-reader`, `pptx`, `xlsx`):
 * it reads straight off the zip, so a bug shared by both readings could not pass its own check. Allowed: the zip package
 * reader (the byte ceilings a real extractor already enforces) and the number-format ENGINE, so that a chart value reads
 * « 81% » on both sides (written limit in `source-reading-chart.ts`). The SmartArt control reading shares no module with
 * the extractor either.
 */
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join, posix } from "node:path";

import { describe, expect, it } from "vitest";

const PIPELINE = import.meta.dirname;
const SOURCE_ROOT = dirname(PIPELINE);

const ALLOWED_EXTRACTOR_IMPORTS = new Set(["extractor/ooxml/package", "extractor/xlsx-number-format"]);

/** The specifiers a file imports or re-exports, type-only ones included: a type names the extractor too. */
const specifiersOf = (source: string): string[] =>
  source
    .split("\n")
    .map((line) => /^\s*(?:import|export)\s[^"']*from\s+["']([^"']+)["']/u.exec(line)?.[1])
    .filter((specifier): specifier is string => specifier !== undefined);

describe("the second reading of conservation never reads a format by its own extractor", () => {
  const secondReadings = readdirSync(PIPELINE).filter(
    (name) => /^source-reading-[^/]+\.ts$/u.test(name) && !name.includes(".test."),
  );

  it("finds the second readings (a sweep that reads nothing guards nothing)", () => {
    expect(secondReadings.length).toBeGreaterThan(5);
  });

  it("🔴 none of them imports the docx, pptx or xlsx extractor beyond the two allowed modules", () => {
    const offences: string[] = [];
    for (const name of secondReadings)
      for (const specifier of specifiersOf(readFileSync(join(PIPELINE, name), "utf8"))) {
        if (!specifier.startsWith(".")) continue;
        const target = posix.normalize(posix.join("pipeline", specifier));
        const touchesExtractor = ["docx", "pptx", "xlsx"].some((format) => target.startsWith(`extractor/${format}`));
        if (touchesExtractor && !ALLOWED_EXTRACTOR_IMPORTS.has(target)) offences.push(`${name} → ${target}`);
      }
    expect(offences).toEqual([]);
  });

  it("🔴 the control reading of SmartArts shares no module with the extractor", () => {
    const imports = specifiersOf(readFileSync(join(PIPELINE, "source-reading-smartart.ts"), "utf8"));
    expect(imports.length).toBeGreaterThan(0);
    expect(imports.filter((specifier) => /extractor\//u.test(specifier) || /\/smartart$/u.test(specifier))).toEqual([]);
  });

  it("the source root it reads is the library's own `src/`", () => {
    expect(readdirSync(SOURCE_ROOT)).toContain("extractor");
  });
});
