import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractPptx } from "../pptx";

// Decision of 08/10/2026: an unreadable theme keeps the deck published with default colours (existing behaviour)
// and is NOW said in one { kind: "styles-defaulted" } notice. An unreadable layout or master keeps refusing the deck.
// Witness: fixtures/pptx/03-Presentation-Resultats-T1.pptx (same technique as pptx-reader.test.ts, « arbitrage 28/09 »).
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/pptx-reader/pptx-reader.damaged-theme-notice.test.ts
const here = dirname(fileURLToPath(import.meta.url));
const deck = (): Uint8Array =>
  new Uint8Array(readFileSync(join(here, "../../../fixtures/pptx/03-Presentation-Resultats-T1.pptx")));

const corruptXml = async (path: string): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(deck());
  const xml = await zip.file(path)!.async("string");
  const offset = xml.indexOf("?>") + 1;
  zip.file(path, xml.slice(0, offset) + "<p:corruptedUnclosed" + xml.slice(offset));
  return zip.generateAsync({ type: "uint8array" });
};

const defaulted = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "styles-defaulted");

describe("S3 pptx — theme unreadable", () => {
  it("non-regression: the untouched deck says nothing about styles", async () => {
    const { doc, notices } = await extractPptx(deck());
    expect(doc.blocks.length).toBeGreaterThan(0);
    expect(defaulted(notices)).toEqual([]);
  });

  it("non-regression: a damaged theme still publishes the deck, its text intact", async () => {
    const { doc } = await extractPptx(await corruptXml("ppt/theme/theme1.xml"));
    expect(JSON.stringify(doc)).toContain("Résultats T1 2026");
  });

  it("🔴 a theme whose XML does not parse: one styles-defaulted notice with no other field", async () => {
    // Failure value (before): no notice at all.
    const { notices } = await extractPptx(await corruptXml("ppt/theme/theme1.xml"));
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });
});

describe("S3 pptx — layout and master unreadable keep refusing", () => {
  it.each(["ppt/slideLayouts/slideLayout1.xml", "ppt/slideMasters/slideMaster1.xml"])(
    "non-regression: %s whose XML does not parse still rejects with EXTRACT_FAILED",
    async (path) => {
      await expect(extractPptx(await corruptXml(path))).rejects.toMatchObject({ code: "EXTRACT_FAILED" });
    },
  );
});
