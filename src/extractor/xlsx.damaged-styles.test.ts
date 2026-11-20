import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { extractXlsx } from "./xlsx";

// Decision of 08/10/2026: an unreadable xl/styles.xml no longer refuses the workbook. It is published, numbers
// and dates shown WITHOUT format (raw value), and the loss is said in one { kind: "styles-defaulted" } notice.
// An unreadable sharedStrings.xml keeps failing (the cell texts would be lost).
// Witness: fixtures/xlsx/09-Effectifs-Dates.xlsx, B2 = 45362 with style xf 1 (numFmtId 14, a date), B3 = 45963.
// Runs in the "main" vitest project:
//   npx vitest run --project=main src/extractor/xlsx.damaged-styles.test.ts
const here = dirname(fileURLToPath(import.meta.url));
const dates = (): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures/xlsx/09-Effectifs-Dates.xlsx")));

const edited = async (edit: (zip: JSZip) => void): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(dates());
  edit(zip);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

/** Breaks the DEFLATE stream of one entry behind its intact local header. */
const corruptStream = (bytes: Uint8Array, entry: string): Uint8Array => {
  const out = new Uint8Array(bytes);
  const name = new TextEncoder().encode(entry);
  const view = new DataView(out.buffer, out.byteOffset);
  let nameAt = -1;
  for (let i = 30; i < out.length - name.length; i++) {
    if (!name.every((b, j) => out[i + j] === b)) continue;
    if (view.getUint32(i - 30, true) === 0x04034b50 && view.getUint16(i - 30 + 26, true) === name.length) {
      nameAt = i;
      break;
    }
  }
  expect(nameAt).toBeGreaterThan(0);
  const dataStart = nameAt + name.length + view.getUint16(nameAt - 30 + 28, true);
  for (let i = 0; i < 8; i++) out[dataStart + i] = (out[dataStart + i] ?? 0) ^ 0xff;
  return out;
};

const defaulted = (notices: readonly { kind: string }[] | undefined): unknown[] =>
  (notices ?? []).filter((n) => n.kind === "styles-defaulted");
const dump = (value: unknown): string => JSON.stringify(value);
const WORDS = ["Ana Ferreira", "Luc Moreau", "Comptabilité", "Support"];

describe("S2 xlsx — xl/styles.xml unreadable", () => {
  it("non-regression: the untouched workbook shows its dates formatted (not as 45362) and says nothing", async () => {
    const { doc, notices } = await extractXlsx(dates());
    const json = dump(doc);
    // Proves the fixture exercises the date format: the raw serial must NOT appear when the styles are readable.
    expect(json).not.toContain("45362");
    expect(json).not.toContain("45963");
    for (const w of WORDS) expect(json).toContain(w);
    expect(defaulted(notices)).toEqual([]);
  });

  it("🔴 XML that does not parse: published, dates shown raw, texts intact, one styles-defaulted notice with no other field", async () => {
    // Failure value (before): extractXlsx rejects with EXTRACT_FAILED.
    const bytes = await edited((zip) => zip.file("xl/styles.xml", '<?xml version="1.0"?><styleSheet <<< not xml'));
    const { doc, notices } = await extractXlsx(bytes);
    const json = dump(doc);
    expect(json).toContain("45362");
    expect(json).toContain("45963");
    for (const w of WORDS) expect(json).toContain(w);
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });

  it("🔴 corrupt DEFLATE stream: published, dates shown raw, one styles-defaulted notice", async () => {
    const bytes = corruptStream(await edited(() => undefined), "xl/styles.xml");
    const { doc, notices } = await extractXlsx(bytes);
    const json = dump(doc);
    expect(json).toContain("45362");
    for (const w of WORDS) expect(json).toContain(w);
    expect(defaulted(notices)).toEqual([{ kind: "styles-defaulted" }]);
  });
});

describe("S2 xlsx — sharedStrings.xml unreadable keeps failing", () => {
  it("non-regression: XML that does not parse still rejects with EXTRACT_FAILED", async () => {
    const bytes = await edited((zip) => zip.file("xl/sharedStrings.xml", '<?xml version="1.0"?><sst <<< not xml'));
    await expect(extractXlsx(bytes)).rejects.toMatchObject({ code: "EXTRACT_FAILED" });
  });

  it("non-regression: a corrupt DEFLATE stream still rejects with EXTRACT_FAILED", async () => {
    const bytes = corruptStream(await edited(() => undefined), "xl/sharedStrings.xml");
    await expect(extractXlsx(bytes)).rejects.toMatchObject({ code: "EXTRACT_FAILED" });
  });
});
