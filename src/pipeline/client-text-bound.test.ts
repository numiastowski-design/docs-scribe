/**
 * Two bounds this package holds equal to the base's: the length client text is cut at, and the image byte budget.
 * WHERE THEY COME FROM. Until 19/11/2026 the publisher's app held both equalities in its own tests; each compares a
 * constant of this package with one of the base, so it is held here, where both can be read.
 * Run: npx vitest run --project=main packages/docs-scribe/src/pipeline/client-text-bound.test.ts
 */
import { describe, expect, it } from "vitest";

import { MAX_PAYLOAD_BYTES } from "doc-ir";
import { MAX_NOTICE_SHEET_NAME_CHARS } from "doc-ir/internal";

import { IMAGE_MAX_TOTAL_BYTES } from "../extractor/ooxml/package";
import { MISSING_WORD_SHOWN_MAX_CHARS } from "./declared-format-conservation";

describe("cutNoticeName — a sheet name shown in a dashboard notice is bounded (item E, 30/09/2026)", () => {
  it("uses the same bound as the missing-words preview (one 64-character rule for client text on the dashboard)", () => {
    expect(MAX_NOTICE_SHEET_NAME_CHARS).toBe(MISSING_WORD_SHOWN_MAX_CHARS);
  });
});

describe("reader-limits — the byte ceiling has one source", () => {
  it("IMAGE_MAX_TOTAL_BYTES equals MAX_PAYLOAD_BYTES", () => {
    // Failure value: a budget of its own, so a picture set the payload ceiling accepts is cut, or the reverse.
    expect(IMAGE_MAX_TOTAL_BYTES).toBe(MAX_PAYLOAD_BYTES);
  });
});
