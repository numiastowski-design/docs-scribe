/**
 * The small names of the published light door (`docs-scribe/light`) that no other test of the package takes: the MIME
 * types of the formats read, the two predicates on a classified type, the two word keys and the two names of the
 * reader's limits. One short describe per module that defines them; every expected value is written here by hand, never read from the
 * module under test.
 */
import { describe, expect, it } from "vitest";

import {
  classifyMimeType,
  conservationKey,
  cutNoticeName,
  DOCX_MIME,
  GOOGLE_DOC,
  GOOGLE_SHEET,
  GOOGLE_SLIDES,
  isGoogleNative,
  isSupportedKind,
  isWorkbookKind,
  MAX_PAYLOAD_BYTES,
  PPTX_MIME,
  wordKeysOf,
  XLSX_MIME,
} from "./docs-light";

describe("the MIME types of the formats read (lib/mime-types.ts)", () => {
  it("the three Office types are the ones IANA registered for OOXML", () => {
    expect(DOCX_MIME).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(XLSX_MIME).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(PPTX_MIME).toBe("application/vnd.openxmlformats-officedocument.presentationml.presentation");
  });

  it("the three Google types are the ones Drive gives its native files", () => {
    expect(GOOGLE_DOC).toBe("application/vnd.google-apps.document");
    expect(GOOGLE_SHEET).toBe("application/vnd.google-apps.spreadsheet");
    expect(GOOGLE_SLIDES).toBe("application/vnd.google-apps.presentation");
  });

  it("isGoogleNative: a Google type is native, an Office file stored in Drive is not", () => {
    for (const type of [GOOGLE_DOC, GOOGLE_SHEET, GOOGLE_SLIDES, "application/vnd.google-apps.form"])
      expect(isGoogleNative(type), type).toBe(true);
    for (const type of [DOCX_MIME, XLSX_MIME, PPTX_MIME, "application/pdf", "", "text/vnd.google-apps.document"])
      expect(isGoogleNative(type), type).toBe(false);
  });
});

describe("the predicates on a classified type (extractor/mime.ts)", () => {
  it("each of the six constants classifies as its own kind", () => {
    expect(
      [DOCX_MIME, XLSX_MIME, PPTX_MIME, GOOGLE_DOC, GOOGLE_SHEET, GOOGLE_SLIDES].map((type) => classifyMimeType(type)),
    ).toEqual(["docx", "xlsx", "pptx", "google-doc", "google-sheet", "google-slides"]);
  });

  it("isSupportedKind: every kind a reader exists for, and neither « unsupported » nor « shortcut »", () => {
    for (const kind of ["google-doc", "google-sheet", "google-slides", "docx", "xlsx", "pptx", "pdf", "image"] as const)
      expect(isSupportedKind(kind), kind).toBe(true);
    expect(isSupportedKind("unsupported")).toBe(false);
    expect(isSupportedKind("shortcut")).toBe(false);
  });

  it("isWorkbookKind: an xlsx and a Google Sheet, nothing else", () => {
    expect(isWorkbookKind("xlsx")).toBe(true);
    expect(isWorkbookKind("google-sheet")).toBe(true);
    for (const kind of [
      "google-doc",
      "google-slides",
      "docx",
      "pptx",
      "pdf",
      "image",
      "unsupported",
      "shortcut",
    ] as const)
      expect(isWorkbookKind(kind), kind).toBe(false);
  });
});

describe("the form two words are compared in (doc-ir, pipeline/text-conservation.ts)", () => {
  it("conservationKey keeps letters and digits only, in lower case, a ligature unfolded", () => {
    expect(conservationKey("L'École,")).toBe("lécole");
    expect(conservationKey("2026-11")).toBe("202611");
    expect(conservationKey("ﬁn")).toBe("fin");
    expect(conservationKey("« … »")).toBe("");
  });

  it("wordKeysOf splits on white space and drops a word that holds no letter or digit", () => {
    expect(wordKeysOf("Le chat — dort,\n\tdéjà !")).toEqual(["le", "chat", "dort", "déjà"]);
    expect(wordKeysOf("aujourd'hui porte-monnaie")).toEqual(["aujourdhui", "portemonnaie"]);
    expect(wordKeysOf("  …  ")).toEqual([]);
  });
});

describe("the names of the reader's limits (doc-ir, lib/reader-limits.ts)", () => {
  it("MAX_PAYLOAD_BYTES is fifty million bytes", () => {
    expect(MAX_PAYLOAD_BYTES).toBe(50_000_000);
  });

  it("cutNoticeName keeps a name of 64 characters whole", () => {
    const name = "n".repeat(64);
    expect(cutNoticeName(name)).toBe(name);
    expect(cutNoticeName("Budget 2026")).toBe("Budget 2026");
    expect(cutNoticeName("")).toBe("");
  });

  it("cutNoticeName cuts a longer name to its first 64 characters and ends it with « … »", () => {
    const cut = cutNoticeName(`${"a".repeat(64)}bcd`);
    expect(cut).toBe(`${"a".repeat(64)}…`);
    expect(cut).toHaveLength(65);
  });
});
