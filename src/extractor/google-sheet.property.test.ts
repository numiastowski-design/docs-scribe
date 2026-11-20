/**
 * Fuzz testing — extractors, priority 1, Google-native path (docs/
 * FUZZ-TESTING.md [purged 27/09/2026, git history] §2-4). Same "do both" chantier as docx.property.test.ts.
 * `extractGoogleSheet` is pure, synchronous object-shape folding (no ZIP,
 * no byte parsing) — the doc's own framing applies directly: "JSON valide
 * mais champs obligatoires absents ou de type inattendu."
 */

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { ReaderError } from "doc-ir";
import { extractGoogleSheet, type GoogleSheetMeta } from "./google-sheet";

async function assertNeverCrashesOnlyNamedFailure(meta: unknown, valuesByTitle: unknown): Promise<void> {
  try {
    await extractGoogleSheet(meta as GoogleSheetMeta, valuesByTitle as Map<string, string[][]>);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError); // P1
    expect((error as ReaderError).code).toBe("EXTRACT_FAILED"); // P2
  }
}

describe("extractGoogleSheet - regression: fixed unclassified crash on missing meta.sheets (STATUS.md 04/09/2026, found by fuzzing, fixed same day)", () => {
  it("now fails with a named EXTRACT_FAILED ReaderError, never a raw TypeError, when meta.sheets is missing", async () => {
    // Root cause (was): `extractGoogleSheet` did `[...meta.sheets]` with
    // zero validation - `meta.sheets` missing/undefined threw "TypeError:
    // meta.sheets is not iterable" directly, never wrapped. The outer sync
    // loop (run-sync.ts:209, `error instanceof ReaderError ? error.code :
    // "UNKNOWN"`) already caught it before it could crash the whole batch
    // (P1/P3 were safe one layer up), but reclassified it as bare
    // "UNKNOWN" - a real P2 violation ("un échec est toujours un échec
    // nommé"). Fix: an explicit `Array.isArray` guard raises the real,
    // named EXTRACT_FAILED directly.
    let thrown: unknown;
    try {
      await extractGoogleSheet({} as GoogleSheetMeta, new Map());
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(ReaderError);
    expect((thrown as ReaderError).code).toBe("EXTRACT_FAILED");
  });
});

describe("extractGoogleSheet - corruption fuzzing, malformed meta/values shapes (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("missing/malformed sheet metadata fields never crash the process, only ever fail named", async () => {
    const malformedSheetArbitrary = fc.record(
      {
        properties: fc.record({
          title: fc.oneof(fc.string(), fc.constant(undefined)),
          index: fc.oneof(fc.integer(), fc.constant(Number.NaN), fc.constant(undefined)),
          hidden: fc.option(fc.boolean(), { nil: undefined }),
        }),
        merges: fc.option(
          fc.array(
            fc.record({
              startRowIndex: fc.oneof(fc.integer({ min: -10, max: 10 }), fc.constant(Number.NaN)),
              endRowIndex: fc.oneof(fc.integer({ min: -10, max: 10 }), fc.constant(Number.NaN)),
              startColumnIndex: fc.oneof(fc.integer({ min: -10, max: 10 }), fc.constant(Number.NaN)),
              endColumnIndex: fc.oneof(fc.integer({ min: -10, max: 10 }), fc.constant(Number.NaN)),
            }),
          ),
          { nil: undefined },
        ),
        rowMetadata: fc.option(fc.array(fc.record({ hiddenByUser: fc.option(fc.boolean(), { nil: undefined }) })), {
          nil: undefined,
        }),
        columnMetadata: fc.option(fc.array(fc.record({ hiddenByUser: fc.option(fc.boolean(), { nil: undefined }) })), {
          nil: undefined,
        }),
      },
      { requiredKeys: ["properties"] },
    );
    const metaArbitrary = fc.record({ sheets: fc.array(malformedSheetArbitrary, { maxLength: 4 }) });
    const valuesArbitrary = fc
      .array(fc.tuple(fc.string(), fc.array(fc.array(fc.string(), { maxLength: 4 }), { maxLength: 4 })), {
        maxLength: 4,
      })
      .map((entries) => new Map(entries));

    await fc.assert(
      fc.asyncProperty(metaArbitrary, valuesArbitrary, async (meta, valuesByTitle) => {
        await assertNeverCrashesOnlyNamedFailure(meta, valuesByTitle);
      }),
    );
  });
});
