import { describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";

/** Review of 27/09/2026 night, finding 6: the blanket that turned ANY failure
 *  of the reader into « the Google Docs export could not be opened » survived
 *  here after it left docx and pptx (27/09 evening, finding 4). A bug of OUR
 *  reader must stay ours — UNKNOWN, retried, visibly our defect — never blame
 *  the user's document. The damaged-file message itself travels through the
 *  reader's own `onDamaged` hook (the app's `google-app/src/pipeline/normalize-links.google-doc.test.ts`). */
vi.mock("./docx-reader", () => ({
  readDocx: vi.fn(async () => {
    throw new TypeError("Cannot read properties of undefined (reading 'kids')");
  }),
}));

const { extractGoogleDoc } = await import("./google-doc");

describe("extractGoogleDoc — a defect of our reader is not a damaged export", () => {
  it("🔴 a TypeError from the reader comes out as that TypeError, never as EXTRACT_FAILED", async () => {
    const failure = await extractGoogleDoc(new Uint8Array([1])).then(
      () => "read",
      (e: unknown) => e,
    );
    expect(failure).toBeInstanceOf(TypeError);
    expect(failure).not.toBeInstanceOf(ReaderError);
  });
});
