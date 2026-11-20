/**
 * Fuzz testing of the extractors, by corruption.
 *
 * `fast-check` core package only (no `@fast-check/vitest` wrapper).
 *
 * Method: a corrupted docx is not made of random bytes. Every corrupted docx here is
 * the REAL golden-corpus fixture, mutated via `jszip` — fast-check only chooses
 * WHERE to truncate, WHICH zip part to drop, WHERE in the XML to corrupt.
 *
 * THE PROPERTIES (MIKSUDOC « Spécification technique » §13, written before any test):
 *  - P1: never crashes the process — always resolves or throws ReaderError.
 *  - P2: a real failure is a NAMED failure (EXTRACT_FAILED), never UNKNOWN.
 *  - P5: bounded termination — no corrupted input hangs (per-case timeout).
 *  - P6: no document content leaks into the error message.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import fc from "fast-check";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { ReaderError } from "doc-ir";
import { extractDocx } from "./docx";
import { fuzzRuns, fuzzTestTimeout } from "../lib/docs-fuzz-runs.test.support";

// Same reasoning as xlsx.property.test.ts's own file-wide raise: repeated
// JSZip load/regenerate cycles across many fast-check cases can clear
// vitest's default 5s per-test timeout under `npm run check`'s full
// parallel load even when solidly under it in isolation.
vi.setConfig({ testTimeout: fuzzTestTimeout(20000, process.env) });

const here = dirname(fileURLToPath(import.meta.url));
const REAL_FIXTURES = ["01-Guide-Onboarding-Atria.docx", "02-Procedure-Support-IT.docx"] as const;
const fixtureBytes = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../fixtures/docx", name)));

// A real fragment of THIS repo's own fixture content (found by reading the structural assertions
// of the app's `src/renderer/storage.docx.test.ts`, not by re-reading the document
// itself) - if this string ever showed up in an error message, that would
// be a P6 violation (content leak). Deliberately distinctive, not a common
// English word that could false-positive.
const KNOWN_FIXTURE_TEXT_FRAGMENT = "Atria";

async function assertNeverCrashesOnlyNamedFailure(bytes: Uint8Array): Promise<void> {
  try {
    await extractDocx(bytes);
    // Resolving is fine - not every corruption breaks the reader's parser.
  } catch (error) {
    // P1: whatever was thrown, it must be OUR closed catalog, never a raw
    // unexpected exception type leaking through.
    expect(error).toBeInstanceOf(ReaderError);
    const appError = error as ReaderError;
    // P2: a real corruption must be classified, never silently UNKNOWN -
    // extractDocx's own catch-all always uses EXTRACT_FAILED today, so this
    // also confirms that wrapping never regresses to something looser.
    expect(appError.code).toBe("EXTRACT_FAILED");
    // P6: the message is ours (never a library's words), and must never
    // contain real document content.
    expect(appError.message).not.toContain(KNOWN_FIXTURE_TEXT_FRAGMENT);
  }
}

describe("extractDocx - corruption fuzzing, truncation (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("truncating a real docx at any byte percentage never crashes the process, only ever fails named", async () => {
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
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 }, // P5: each case bounded to 5s - a hang IS a finding, not a slow pass
    );
  });
});

const DOCX_REQUIRED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
  "word/_rels/document.xml.rels",
] as const;

describe("extractDocx - corruption fuzzing, missing required ZIP part (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a well-formed ZIP missing a required docx part never crashes the process, only ever fails named", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.constantFrom(...DOCX_REQUIRED_PARTS),
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

/** Corrupts `word/document.xml`'s TEXT at a fast-check-chosen offset -
 *  the recipe (offset, corruption kind), not raw random bytes, per
 *  FUZZ-TESTING.md [purged 27/09/2026, git history] §7's own methodology. */
async function corruptDocumentXmlAt(
  bytes: Uint8Array,
  offsetFraction: number,
  kind: "unclosed-tag" | "control-char" | "undefined-entity",
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) throw new Error("fixture missing word/document.xml - not a valid test input");
  const xml = await documentXmlFile.async("string");
  const offset = Math.max(0, Math.min(xml.length - 1, Math.floor(xml.length * offsetFraction)));
  let corrupted: string;
  switch (kind) {
    case "unclosed-tag":
      corrupted = xml.slice(0, offset) + "<w:corruptedUnclosed" + xml.slice(offset);
      break;
    case "control-char":
      corrupted = xml.slice(0, offset) + String.fromCharCode(0x01) + xml.slice(offset);
      break;
    case "undefined-entity":
      corrupted = xml.slice(0, offset) + "&undefinedEntity;" + xml.slice(offset);
      break;
  }
  zip.file("word/document.xml", corrupted);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
}

describe("extractDocx - corruption fuzzing, malformed document.xml (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("an unclosed tag, a raw control character, or an undefined entity injected anywhere in document.xml never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.double({ min: 0, max: 1, noNaN: true }),
        fc.constantFrom("unclosed-tag", "control-char", "undefined-entity"),
        async (fixtureName, offsetFraction, kind) => {
          const corrupted = await corruptDocumentXmlAt(fixtureBytes(fixtureName), offsetFraction, kind);
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 },
    );
  });
});

describe("extractDocx - corruption fuzzing, broken .rels reference (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a relationship pointing at a part that doesn't exist never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...REAL_FIXTURES), async (fixtureName) => {
        const zip = await JSZip.loadAsync(fixtureBytes(fixtureName));
        const relsFile = zip.file("word/_rels/document.xml.rels");
        if (!relsFile) return; // this fixture has no rels part to corrupt - not a failure of the property
        const relsXml = await relsFile.async("string");
        // Redirect every existing Target="..." to a part that certainly
        // doesn't exist in this archive - a plausible real-world corruption
        // (a partial sync, a manual zip edit) rather than random noise.
        const corruptedRels = relsXml.replace(/Target="[^"]*"/g, 'Target="media/does-not-exist.png"');
        zip.file("word/_rels/document.xml.rels", corruptedRels);
        const corrupted = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
        await assertNeverCrashesOnlyNamedFailure(corrupted);
      }),
      { numRuns: fuzzRuns(10, process.env), timeout: 5000 },
    );
  });
});
