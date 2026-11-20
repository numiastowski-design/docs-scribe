/**
 * Fuzz testing — extractors, priority 1 (docs/FUZZ-TESTING.md [purged 27/09/2026, git history] §2-4).
 * Same "do both" chantier as docx.property.test.ts — see that file's header
 * for the full method/property rationale, not repeated here.
 *
 * Written against the former extractor, whose try/catch wrapped only
 * `JSZip.loadAsync`. Since the own pptx reader (25/09/2026) the file is
 * read by `extractor/pptx-reader/`, which names a damaged package or slide
 * part `EXTRACT_FAILED` itself; these properties still hold the whole path
 * (reader, layout, drawing) to that single named failure.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import fc from "fast-check";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { ReaderError } from "doc-ir";
import { extractPptx } from "./pptx";
import { fuzzRuns, fuzzTestTimeout } from "../lib/docs-fuzz-runs.test.support";

// Same reasoning as xlsx.property.test.ts's own file-wide raise.
vi.setConfig({ testTimeout: fuzzTestTimeout(20000, process.env) });

const here = dirname(fileURLToPath(import.meta.url));
const REAL_FIXTURES = ["03-Presentation-Resultats-T1.pptx"] as const;
const fixtureBytes = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../fixtures/pptx", name)));
const KNOWN_FIXTURE_TEXT_FRAGMENT = "Resultats";

async function assertNeverCrashesOnlyNamedFailure(bytes: Uint8Array): Promise<void> {
  try {
    await extractPptx(bytes);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError); // P1
    const appError = error as ReaderError;
    expect(appError.code).toBe("EXTRACT_FAILED"); // P2
    expect(appError.message).not.toContain(KNOWN_FIXTURE_TEXT_FRAGMENT); // P6
  }
}

describe("extractPptx - corruption fuzzing, truncation (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("truncating a real pptx at any byte percentage never crashes the process, only ever fails named", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.integer({ min: 1, max: 99 }),
        async (fixtureName, percent) => {
          const original = fixtureBytes(fixtureName);
          const truncated = original.slice(0, Math.floor((original.length * percent) / 100));
          await assertNeverCrashesOnlyNamedFailure(truncated);
        },
      ),
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 },
    );
  });
});

const PPTX_REQUIRED_PARTS = [
  "[Content_Types].xml",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
  "ppt/slides/slide1.xml",
  "ppt/slides/_rels/slide1.xml.rels",
] as const;

describe("extractPptx - corruption fuzzing, missing required ZIP part (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a well-formed ZIP missing a required pptx part never crashes the process, only ever fails named", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.constantFrom(...PPTX_REQUIRED_PARTS),
        async (fixtureName, partToRemove) => {
          const zip = await JSZip.loadAsync(fixtureBytes(fixtureName));
          zip.remove(partToRemove);
          const corrupted = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(20, process.env), timeout: 5000 },
    );
  });
});

async function corruptSlideXmlAt(
  bytes: Uint8Array,
  slidePath: string,
  offsetFraction: number,
  kind: "unclosed-tag" | "control-char" | "undefined-entity",
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const slideFile = zip.file(slidePath);
  if (!slideFile) throw new Error(`fixture missing ${slidePath} - not a valid test input`);
  const xml = await slideFile.async("string");
  const offset = Math.max(0, Math.min(xml.length - 1, Math.floor(xml.length * offsetFraction)));
  let corrupted: string;
  switch (kind) {
    case "unclosed-tag":
      corrupted = xml.slice(0, offset) + "<p:corruptedUnclosed" + xml.slice(offset);
      break;
    case "control-char":
      corrupted = xml.slice(0, offset) + String.fromCharCode(0x01) + xml.slice(offset);
      break;
    case "undefined-entity":
      corrupted = xml.slice(0, offset) + "&undefinedEntity;" + xml.slice(offset);
      break;
  }
  zip.file(slidePath, corrupted);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
}

describe("extractPptx - corruption fuzzing, malformed slide XML (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("an unclosed tag, a raw control character, or an undefined entity injected anywhere in a slide's XML never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.integer({ min: 1, max: 6 }),
        fc.double({ min: 0, max: 1, noNaN: true }),
        fc.constantFrom("unclosed-tag", "control-char", "undefined-entity"),
        async (fixtureName, slideNumber, offsetFraction, kind) => {
          const corrupted = await corruptSlideXmlAt(
            fixtureBytes(fixtureName),
            `ppt/slides/slide${slideNumber}.xml`,
            offsetFraction,
            kind,
          );
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(40, process.env), timeout: 5000 },
    );
  });
});

describe("extractPptx - corruption fuzzing, broken slide .rels reference (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a slide relationship (image, chart, notes) pointing at a part that doesn't exist never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.integer({ min: 1, max: 6 }),
        async (fixtureName, slideNumber) => {
          const zip = await JSZip.loadAsync(fixtureBytes(fixtureName));
          const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
          const relsFile = zip.file(relsPath);
          if (!relsFile) return;
          const relsXml = await relsFile.async("string");
          const corruptedRels = relsXml.replace(/Target="[^"]*"/g, 'Target="../media/does-not-exist.png"');
          zip.file(relsPath, corruptedRels);
          const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
          await assertNeverCrashesOnlyNamedFailure(bytes);
        },
      ),
      { numRuns: fuzzRuns(15, process.env), timeout: 5000 },
    );
  });
});

describe("extractPptx - corruption fuzzing, malformed ppt/presentation.xml (fuzz: parseSlideOrder's own slide-list parse)", () => {
  it("an unclosed tag, a raw control character, or an undefined entity injected anywhere in presentation.xml never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.double({ min: 0, max: 1, noNaN: true }),
        fc.constantFrom("unclosed-tag", "control-char", "undefined-entity"),
        async (fixtureName, offsetFraction, kind) => {
          const corrupted = await corruptSlideXmlAt(
            fixtureBytes(fixtureName),
            "ppt/presentation.xml",
            offsetFraction,
            kind,
          );
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 },
    );
  });
});
