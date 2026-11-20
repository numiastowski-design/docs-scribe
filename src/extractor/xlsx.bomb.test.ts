/**
 * Decompression bomb. The reader is lazy: it follows only the relationships it publishes, so a hidden sheet or an
 * entry nothing references is never opened at all, and the protection is the per-part ceiling applied to every part
 * really read. The package is spied on (a file apart: `vi.mock` holds for the whole file) to prove that nothing was
 * read, not to infer it from the absence of an error. The hostile packages are built in memory.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { crc32, deflateRawSync } from "node:zlib";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { PART_MAX_BYTES } from "./ooxml/package";
import { extractXlsx } from "./xlsx";
import { PAGE_SIZED_LIMITS } from "./page-caps.test.support";

const readPaths = vi.hoisted(() => ({ seen: new Set<string>() }));

vi.mock("./ooxml/package", async (importOriginal) => {
  const real = await importOriginal<typeof import("./ooxml/package")>();
  return {
    ...real,
    openPackage: async (...args: Parameters<typeof real.openPackage>) => {
      const pkg = await real.openPackage(...args);
      return {
        ...pkg,
        readXml: (path: string) => (readPaths.seen.add(path), pkg.readXml(path)),
        readBytes: (path: string, maxBytes?: number) => (readPaths.seen.add(path), pkg.readBytes(path, maxBytes)),
      };
    },
  };
});

/** A zip written entry by entry through node's own zlib (the docx reader's
 *  tests build theirs the same way): tens of MB of one byte deflate to a
 *  few KB in milliseconds. */
const handZip = (entries: readonly { name: string; bytes: Uint8Array }[]): Uint8Array => {
  const enc = new TextEncoder();
  const pieces: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = enc.encode(entry.name);
    const data = new Uint8Array(deflateRawSync(entry.bytes));
    const crc = crc32(entry.bytes);
    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(8, 8, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, entry.bytes.length, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    const dir = new Uint8Array(46 + name.length);
    const dv = new DataView(dir.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    dv.setUint16(10, 8, true);
    dv.setUint32(16, crc, true);
    dv.setUint32(20, data.length, true);
    dv.setUint32(24, entry.bytes.length, true);
    dv.setUint16(28, name.length, true);
    dv.setUint32(42, offset, true);
    dir.set(name, 46);
    pieces.push(local, data);
    central.push(dir);
    offset += local.length + data.length;
  }
  const dirSize = central.reduce((n, d) => n + d.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, dirSize, true);
  ev.setUint32(16, offset, true);
  const all = [...pieces, ...central, end];
  const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of all) {
    out.set(p, at);
    at += p.length;
  }
  return out;
};

const here = dirname(fileURLToPath(import.meta.url));
const fixtureEntries = async (): Promise<{ name: string; bytes: Uint8Array }[]> => {
  const zip = await JSZip.loadAsync(readFileSync(join(here, "../../fixtures/xlsx/04-Effectifs-Atria.xlsx")));
  const out: { name: string; bytes: Uint8Array }[] = [];
  for (const file of Object.values(zip.files))
    if (!file.dir) out.push({ name: file.name, bytes: await file.async("uint8array") });
  return out;
};

/** A sheet part whose cells are real, padded past the per-part ceiling. */
const paddedSheet = (): Uint8Array =>
  new TextEncoder().encode(
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>kept</t></is></c></row></sheetData>' +
      `${" ".repeat(PART_MAX_BYTES + 1_000)}</worksheet>`,
  );

// A part inflating past its ceiling is a hostile-input cap (lot « fichier abîmé », 25/10/2026): refused for
// safety, no attachment offered — the same refusal as a docx or pptx bomb (`ooxml/package.ts`).
const UNSAFE = { code: "FILE_UNSAFE", message: "the file inflates past its size ceiling" };
const tabTitles = (doc: { blocks: IRBlock[] }): string[] =>
  (doc.blocks[0] as Extract<IRBlock, { kind: "tabs" }>).tabs.map((t) => t.title);

describe("extractXlsx — a decompression bomb only reaches parts our reader actually follows", () => {
  it("🔴 the fixture itself reads through unharmed: three tabs, the hidden one dropped", async () => {
    const { doc } = await extractXlsx(handZip(await fixtureEntries()), undefined, PAGE_SIZED_LIMITS);
    expect(tabTitles(doc)).toEqual(["2025", "2026"]); // A7 (29/09/2026): Grand-Onglet is withheld, no tab at all
  });

  it("🔴 a PUBLISHED sheet (sheet1.xml, tab '2025') inflating past PART_MAX_BYTES: refused for safety", async () => {
    const entries = await fixtureEntries();
    const bomb = handZip([
      ...entries.filter((e) => e.name !== "xl/worksheets/sheet1.xml"),
      { name: "xl/worksheets/sheet1.xml", bytes: paddedSheet() },
    ]);
    expect(bomb.length).toBeLessThan(1_000_000);
    expect(
      await extractXlsx(bomb, undefined, PAGE_SIZED_LIMITS).then(
        () => "read",
        (e: unknown) => e,
      ),
    ).toMatchObject(UNSAFE);
    // Own deadline (08/10/2026): this file stays in « main » (vi.mock) and inflates megabytes; it went past 5 s beside other workers.
  }, 20_000);

  // The visible tab '2025' is read and enforced (case above); the HIDDEN
  // tab (sheet4.xml, Archive-Interne) and an entry nothing references are
  // never opened at all — same padding, no refusal because there is
  // nothing to refuse: our reader never asks for their bytes.
  for (const target of ["xl/worksheets/sheet4.xml", "xl/media/unreferenced.bin"]) {
    it(`🔴 ${target} (hidden tab / unreferenced entry) padded past PART_MAX_BYTES: never read, extraction unharmed`, async () => {
      const entries = await fixtureEntries();
      const bomb = handZip([...entries.filter((e) => e.name !== target), { name: target, bytes: paddedSheet() }]);
      expect(bomb.length).toBeLessThan(1_000_000);
      readPaths.seen.clear();
      const { doc } = await extractXlsx(bomb, undefined, PAGE_SIZED_LIMITS);
      expect(tabTitles(doc)).toEqual(["2025", "2026"]); // A7 (29/09/2026): Grand-Onglet withheld, no tab; unharmed: the bomb sat where nothing reads
      expect(readPaths.seen.has(target)).toBe(false); // the proof: never even opened
    });
  }
});
