/**
 * Fuzz testing — extractors, priority 1 (docs/FUZZ-TESTING.md [purged 27/09/2026, git history] §2-4).
 * Same "do both" chantier as docx.property.test.ts — see that file's header
 * for the full method/property rationale, not repeated here.
 *
 * Real finding worth flagging before any test: `extractXlsx`'s own
 * try/catch (xlsx.ts) wraps ONLY the `read()` call — everything after
 * (`sheetToTable`, the per-sheet loop) runs UNPROTECTED. `sheetToTable`
 * calls `utils.decode_range(ref)` directly on a sheet's own `!ref` (derived
 * from the real `<dimension ref="...">` XML element) with no guard of its
 * own — a prime candidate for a P1/P2 violation (an unwrapped exception
 * escaping as something other than ReaderError) if SheetJS's `read()` accepts
 * a workbook whose `!ref` string doesn't parse cleanly. Fuzzed explicitly
 * below, not left to chance.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import fc from "fast-check";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { ReaderError } from "doc-ir";
import { extractXlsx } from "./xlsx";
import { fuzzRuns, fuzzTestTimeout } from "../lib/docs-fuzz-runs.test.support";

// Real, repeated JSZip load/regenerate cycles across ~20-30 fast-check cases
// per test comfortably clear vitest's default 5s per-test timeout when run
// ALONE, but not always under `npm run check`'s full parallel load (found
// live, 04/09/2026 — 3 tests here timed out under load despite passing
// reliably in isolation just before). Raised file-wide rather than per-call.
vi.setConfig({ testTimeout: fuzzTestTimeout(20000, process.env) });

const here = dirname(fileURLToPath(import.meta.url));
const REAL_FIXTURES = ["04-Effectifs-Atria.xlsx"] as const;
const fixtureBytes = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx", name)));
const KNOWN_FIXTURE_TEXT_FRAGMENT = "Atria";

async function assertNeverCrashesOnlyNamedFailure(bytes: Uint8Array): Promise<void> {
  try {
    await extractXlsx(bytes); // async since the lot tableurs (step 0, 28/09/2026)
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError); // P1
    const appError = error as ReaderError;
    expect(appError.code).toBe("EXTRACT_FAILED"); // P2
    expect(appError.message).not.toContain(KNOWN_FIXTURE_TEXT_FRAGMENT); // P6
  }
}

describe("extractXlsx - corruption fuzzing, truncation (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  // The property is AWAITED (17/11/2026). Before, the test returned while it was still running: a failure came out as an
  // unhandled rejection of the run, attached to no test, and this one was reported passed.
  it("truncating a real xlsx at any byte percentage never crashes the process, only ever fails named", async () => {
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
    );
  });
});

const XLSX_REQUIRED_PARTS = [
  "[Content_Types].xml",
  "xl/workbook.xml",
  "xl/worksheets/sheet1.xml",
  "xl/sharedStrings.xml",
] as const;

describe("extractXlsx - corruption fuzzing, missing required ZIP part (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a well-formed ZIP missing a required xlsx part never crashes the process, only ever fails named", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.constantFrom(...XLSX_REQUIRED_PARTS),
        async (fixtureName, partToRemove) => {
          const zip = await JSZip.loadAsync(fixtureBytes(fixtureName));
          zip.remove(partToRemove);
          const corrupted = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(20, process.env) },
    );
  });
});

async function corruptSheetXmlAt(
  bytes: Uint8Array,
  offsetFraction: number,
  kind: "unclosed-tag" | "control-char" | "undefined-entity",
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const sheetFile = zip.file("xl/worksheets/sheet1.xml");
  if (!sheetFile) throw new Error("fixture missing xl/worksheets/sheet1.xml - not a valid test input");
  const xml = await sheetFile.async("string");
  const offset = Math.max(0, Math.min(xml.length - 1, Math.floor(xml.length * offsetFraction)));
  let corrupted: string;
  switch (kind) {
    case "unclosed-tag":
      corrupted = xml.slice(0, offset) + "<corruptedUnclosed" + xml.slice(offset);
      break;
    case "control-char":
      corrupted = xml.slice(0, offset) + String.fromCharCode(0x01) + xml.slice(offset);
      break;
    case "undefined-entity":
      corrupted = xml.slice(0, offset) + "&undefinedEntity;" + xml.slice(offset);
      break;
  }
  zip.file("xl/worksheets/sheet1.xml", corrupted);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
}

describe("extractXlsx - corruption fuzzing, malformed sheet XML (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("an unclosed tag, a raw control character, or an undefined entity injected anywhere in sheet1.xml never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.double({ min: 0, max: 1, noNaN: true }),
        fc.constantFrom("unclosed-tag", "control-char", "undefined-entity"),
        async (fixtureName, offsetFraction, kind) => {
          const corrupted = await corruptSheetXmlAt(fixtureBytes(fixtureName), offsetFraction, kind);
          await assertNeverCrashesOnlyNamedFailure(corrupted);
        },
      ),
      { numRuns: fuzzRuns(30, process.env) },
    );
  });
});

// Verified before trusting a green result (not assumed): does corrupting
// the XML attribute actually reach `sheetToTable`'s unguarded
// `utils.decode_range(ref)` call at all? Checked directly (a throwaway
// script, not committed) - SheetJS's own `read()` silently RECOMPUTES
// `!ref` from the real populated cells, ignoring whatever the declared
// `<dimension ref="...">` attribute says, for every value tried here
// (empty, single-cell, absurd range, negative). So this specific
// corruption vector never reaches the unguarded call - a real, verified
// finding (SheetJS defends against this upstream), not proof the unguarded
// call itself is safe against every possible `!ref` shape. Kept as
// defense-in-depth, described honestly.
describe("extractXlsx - corruption fuzzing, malformed dimension ref (verified: SheetJS recomputes !ref from real cells, never reaches sheetToTable's unguarded utils.decode_range with the corrupted value)", () => {
  it('a nonsensical <dimension ref="..."> value never crashes the process, only ever fails named', async () => {
    const PATHOLOGICAL_REFS = ["", "Z", "A1:", ":A1", "A0:B0", "999999999999:A1", "A1:A1:A1", "-1:-1"] as const;
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...REAL_FIXTURES),
        fc.constantFrom(...PATHOLOGICAL_REFS),
        async (fixtureName, badRef) => {
          const zip = await JSZip.loadAsync(fixtureBytes(fixtureName));
          const sheetFile = zip.file("xl/worksheets/sheet1.xml");
          if (!sheetFile) return;
          const xml = await sheetFile.async("string");
          const corrupted = xml.replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="${badRef}"/>`);
          zip.file("xl/worksheets/sheet1.xml", corrupted);
          const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
          await assertNeverCrashesOnlyNamedFailure(bytes);
        },
      ),
      { numRuns: fuzzRuns(20, process.env) },
    );
  });
});
