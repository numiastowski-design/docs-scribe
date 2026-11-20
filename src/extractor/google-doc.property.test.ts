/**
 * Fuzz testing of the extractors, Google-native path: this path receives API responses, not files, so the
 * corruption is applied to the response. The export is a .docx, read by the docx reader. The corruption vectors: a
 * truncated response, an empty ZIP, a ZIP without the expected part (the main document part, or the package
 * relationships that name it), plus malformed XML inside the main part.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import fc from "fast-check";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import { ReaderError } from "doc-ir";
import { extractGoogleDoc } from "./google-doc";
import { fuzzRuns, fuzzTestTimeout } from "../lib/docs-fuzz-runs.test.support";

// Same reasoning as xlsx.property.test.ts's own file-wide raise.
vi.setConfig({ testTimeout: fuzzTestTimeout(20000, process.env) });

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_BYTES = new Uint8Array(
  readFileSync(join(here, "../../fixtures/gdoc-docx/01-Guide-Onboarding-Atria.docx")),
);
const KNOWN_FIXTURE_TEXT_FRAGMENT = "Atria";
const MAIN_PART = "word/document.xml";

async function assertNeverCrashesOnlyNamedFailure(bytes: Uint8Array): Promise<void> {
  try {
    await extractGoogleDoc(bytes);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError); // P1
    const appError = error as ReaderError;
    expect(appError.code).toBe("EXTRACT_FAILED"); // P2
    expect(appError.message).not.toContain(KNOWN_FIXTURE_TEXT_FRAGMENT); // P6
  }
}

describe("extractGoogleDoc - corruption fuzzing, truncated/empty response (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("truncating the export at any byte percentage, or an entirely empty zip, never crashes the process", async () => {
    await fc.assert(
      fc.asyncProperty(fc.integer({ min: 0, max: 99 }), async (percent) => {
        const truncated = FIXTURE_BYTES.slice(0, Math.floor((FIXTURE_BYTES.length * percent) / 100));
        await assertNeverCrashesOnlyNamedFailure(truncated);
      }),
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 },
    );

    const emptyZip = new Uint8Array(await new JSZip().generateAsync({ type: "uint8array" }));
    await assertNeverCrashesOnlyNamedFailure(emptyZip);
  });
});

describe("extractGoogleDoc - corruption fuzzing, ZIP without the expected part (fuzz: extractors priority 1, MIKSUDOC « Spécification technique » §13)", () => {
  it("a well-formed ZIP without its main document part, or without the package relationships, only ever fails named", async () => {
    for (const missing of [MAIN_PART, "_rels/.rels"]) {
      const zip = await JSZip.loadAsync(FIXTURE_BYTES);
      zip.remove(missing);
      const corrupted = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
      await expect(extractGoogleDoc(corrupted)).rejects.toSatisfy(
        (error) => error instanceof ReaderError && error.code === "EXTRACT_FAILED",
      );
    }
  });
});

describe("extractGoogleDoc - corruption fuzzing, malformed XML in the main part", () => {
  it("control characters or broken tags injected anywhere in document.xml never crash the process", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.double({ min: 0, max: 1, noNaN: true }),
        fc.constantFrom("unclosed-tag", "control-char", "broken-attribute"),
        async (offsetFraction, kind) => {
          const zip = await JSZip.loadAsync(FIXTURE_BYTES);
          const entry = zip.file(MAIN_PART);
          if (!entry) return;
          const xml = await entry.async("string");
          const offset = Math.max(0, Math.min(xml.length - 1, Math.floor(xml.length * offsetFraction)));
          let corrupted: string;
          switch (kind) {
            case "unclosed-tag":
              corrupted = xml.slice(0, offset) + "<w:p unclosed" + xml.slice(offset);
              break;
            case "control-char":
              corrupted = xml.slice(0, offset) + String.fromCharCode(0x01) + xml.slice(offset);
              break;
            case "broken-attribute":
              corrupted = xml.slice(0, offset) + '<w:t w:val="' + xml.slice(offset); // unterminated attribute
              break;
          }
          zip.file(MAIN_PART, corrupted);
          const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
          await assertNeverCrashesOnlyNamedFailure(bytes);
        },
      ),
      { numRuns: fuzzRuns(30, process.env), timeout: 5000 },
    );
  });
});
