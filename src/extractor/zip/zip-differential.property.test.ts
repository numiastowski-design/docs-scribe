/**
 * Slice 4 of the jszip detachment: differential fuzz of our ZIP reader
 * (`openZip`, the product's reader since 05/10/2026, on every file) against jszip.
 *
 * Damaged archives (truncated, one bit flipped, a declared size or length
 * rewritten — PKWARE APPNOTE 6.3.10 §4.3.7 local header, §4.3.12 central
 * header, §4.3.16 end record) go through both readers. The rule held: our
 * reader NEVER accepts what jszip refuses, on the structure (central
 * directory, local headers) nor on an entry's data; when both accept, they
 * agree on the names and on every entry's bytes. Our reader may refuse more
 * (it is stricter by design: names, overlaps, local/central agreement).
 * Any other exception than our two refusal classes is a crash, and fails.
 *
 * Seeds: two real Office documents of the repository (Word, and a Google
 * Sheets export) and one hand-built archive with a data descriptor
 * (§4.3.9, bit 3 of §4.4.4), a stored entry, an empty entry and a directory.
 * fast-check only chooses WHERE and HOW to damage them.
 */
import { readFileSync } from "node:fs";
import { crc32, deflateRawSync } from "node:zlib";

import fc from "fast-check";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { ReaderError } from "doc-ir";
import { readCentralDirectory } from "./central-directory";
import { readLocalHeaders } from "./local-headers";
import { openZip, ZipEntryRefusal } from "./zip-package";

// Each case loads an archive twice and inflates every part: same budget as the other extractor property suites.
vi.setConfig({ testTimeout: fuzzTestTimeout(20_000, process.env) });

/** Large enough that no seed reaches it: a ceiling refusal is still a refusal, in the allowed direction. */
const CEILING = 64_000_000;

// ---------------------------------------------------------------------------
// The two readers, reduced to one verdict each.
// ---------------------------------------------------------------------------

type Verdict =
  | { readonly ok: false; readonly stage: "structure" }
  | {
      readonly ok: true;
      /** Every entry name, directories included, sorted. */
      readonly names: readonly string[];
      /** Each file's bytes, or undefined when that entry's data was refused. */
      readonly data: ReadonlyMap<string, Uint8Array | undefined>;
    };

type Reader = (bytes: Uint8Array) => Promise<Verdict>;

/** Ours: the product's reader, `openZip` (what `ooxml/package.ts` calls since jszip left, 05/10/2026), each
 *  file read under the per-part ceiling; the names, directories included, come from the same headers it reads.
 *  Refusals must be our classes; anything else propagates and fails the test. */
const ours: Reader = async (bytes) => {
  let entries;
  let zip;
  try {
    entries = readLocalHeaders(bytes).entries;
    zip = openZip(bytes, { partMaxBytes: CEILING, packageMaxBytes: Number.MAX_SAFE_INTEGER });
  } catch (e) {
    if (e instanceof ReaderError) return { ok: false, stage: "structure" };
    throw e;
  }
  const data = new Map<string, Uint8Array | undefined>();
  for (const name of zip.names) {
    try {
      data.set(name, await zip.readBytes(name));
    } catch (e) {
      if (!(e instanceof ZipEntryRefusal)) throw e;
      data.set(name, undefined);
    }
  }
  return { ok: true, names: entries.map((e) => e.name).sort(), data };
};

const theirs: Reader = async (bytes) => {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { createFolders: false });
  } catch {
    return { ok: false, stage: "structure" };
  }
  const data = new Map<string, Uint8Array | undefined>();
  for (const file of Object.values(zip.files)) {
    if (file.dir) continue;
    data.set(
      file.name,
      await file.async("uint8array").then(
        (b) => b,
        () => undefined,
      ),
    );
  }
  return { ok: true, names: Object.keys(zip.files).sort(), data };
};

const refusesEntry = (v: Verdict, name: string): boolean => !v.ok || v.data.get(name) === undefined;

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((x, i) => x === b[i]);

/**
 * Every way `candidate` departs from the rule against jszip, as our own
 * literals plus an entry index (never a name: a name is archive content).
 * Empty when the rule holds.
 */
async function departures(bytes: Uint8Array, candidate: Reader): Promise<string[]> {
  const a = await candidate(bytes);
  const b = await theirs(bytes);
  if (!a.ok) return [];
  if (!b.ok) return ["accepts a structure jszip refuses"];
  if (a.names.length !== b.names.length || a.names.some((n, i) => n !== b.names[i])) return ["entry names differ"];
  const found: string[] = [];
  [...a.data].forEach(([name, mine], i) => {
    if (mine === undefined) return;
    const other = b.data.get(name);
    if (other === undefined) found.push(`entry ${i}: accepts data jszip refuses`);
    else if (!sameBytes(mine, other)) found.push(`entry ${i}: bytes differ`);
  });
  return found;
}

// ---------------------------------------------------------------------------
// Seeds.
// ---------------------------------------------------------------------------

const u16 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff];
const u32 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];

interface Part {
  name: string;
  data: Uint8Array;
  deflate: boolean;
  /** Bit 3: zero CRC and sizes locally, then a signed data descriptor (§4.3.9.3). */
  descriptor?: boolean;
  /** Overrides the compressed bytes (deflated parts only). */
  compressed?: Uint8Array;
}

/** A complete archive, written the way §4.3.6 lays it out. No zip64, no comment. */
function build(parts: Part[]): Uint8Array {
  const body: number[] = [];
  const cd: number[] = [];
  for (const p of parts) {
    const name = [...new TextEncoder().encode(p.name)];
    const packed = p.compressed ?? (p.deflate ? new Uint8Array(deflateRawSync(p.data)) : p.data);
    const crc = crc32(p.data);
    const flags = p.descriptor ? 0x08 : 0;
    const method = p.deflate ? 8 : 0;
    const offset = body.length;
    const local = p.descriptor ? [0, 0, 0] : [crc, packed.length, p.data.length];
    body.push(...u32(0x04034b50), ...u16(20), ...u16(flags), ...u16(method), ...u16(0), ...u16(0x21));
    body.push(
      ...u32(local[0]!),
      ...u32(local[1]!),
      ...u32(local[2]!),
      ...u16(name.length),
      ...u16(0),
      ...name,
      ...packed,
    );
    if (p.descriptor) body.push(...u32(0x08074b50), ...u32(crc), ...u32(packed.length), ...u32(p.data.length));
    cd.push(...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(flags), ...u16(method), ...u16(0), ...u16(0x21));
    cd.push(...u32(crc), ...u32(packed.length), ...u32(p.data.length), ...u16(name.length), ...u16(0), ...u16(0));
    cd.push(...u16(0), ...u16(0), ...u32(p.name.endsWith("/") ? 0x10 : 0), ...u32(offset), ...name);
  }
  const eocd = [...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(parts.length), ...u16(parts.length)];
  eocd.push(...u32(cd.length), ...u32(body.length), ...u16(0));
  return new Uint8Array([...body, ...cd, ...eocd]);
}

const text = (s: string): Uint8Array => new TextEncoder().encode(s);

const HAND_BUILT = build([
  {
    name: "[Content_Types].xml",
    data: text(`<Types>${"<Default/>".repeat(40)}</Types>`),
    deflate: true,
    descriptor: true,
  },
  { name: "word/", data: new Uint8Array(), deflate: false },
  { name: "word/document.xml", data: text("<w:document>plain stored part</w:document>"), deflate: false },
  { name: "word/empty.xml", data: new Uint8Array(), deflate: false },
  { name: "word/styles.xml", data: text(`<w:styles>${"<w:style/>".repeat(60)}</w:styles>`), deflate: true },
]);

const SEEDS: readonly Uint8Array[] = [
  ...["fixtures/docx/07-Fiche-Contact.docx", "fixtures/xlsx-reel/21-GSheets-Planning-Formations.xlsx"].map(
    (path) => new Uint8Array(readFileSync(path)),
  ),
  HAND_BUILT,
];

// ---------------------------------------------------------------------------
// Damage.
// ---------------------------------------------------------------------------

/** Where each declared size, length and offset field sits in a VALID seed, found by walking its records. */
function sizeFields(seed: Uint8Array): { at: number; width: 2 | 4 }[] {
  const view = new DataView(seed.buffer, seed.byteOffset, seed.byteLength);
  const cd = readCentralDirectory(seed);
  const fields: { at: number; width: 2 | 4 }[] = [];
  let at = cd.offset;
  for (let i = 0; i < cd.count; i++) {
    // Central header: compressed, uncompressed, name length, extra length, comment length (our walk from one record to
    // the next reads it), local header offset.
    fields.push(
      { at: at + 20, width: 4 },
      { at: at + 24, width: 4 },
      { at: at + 28, width: 2 },
      { at: at + 30, width: 2 },
    );
    fields.push({ at: at + 32, width: 2 });
    fields.push({ at: at + 42, width: 4 });
    const local = view.getUint32(at + 42, true);
    // Local header: compressed, uncompressed, name length, extra length.
    fields.push(
      { at: local + 18, width: 4 },
      { at: local + 22, width: 4 },
      { at: local + 26, width: 2 },
      { at: local + 28, width: 2 },
    );
    at += 46 + view.getUint16(at + 28, true) + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
  }
  // End record: entries on this disk, total entries, directory size, directory offset, comment length.
  fields.push(
    { at: at + 8, width: 2 },
    { at: at + 10, width: 2 },
    { at: at + 12, width: 4 },
    { at: at + 16, width: 4 },
  );
  fields.push({ at: at + 20, width: 2 });
  return fields;
}

type Damage =
  | { kind: "truncate"; keep: number }
  | { kind: "flip"; position: number; bit: number }
  | { kind: "size"; field: number; value: number; relative: boolean }
  /** An entry's central record pointed at ANOTHER entry's real local header (shared or swapped headers). A random
   *  offset almost never lands on one: 0 or 1 time in 3 000 draws (counter-review of 05/10/2026). */
  | { kind: "aim"; entry: number; header: number };

/** One damaged copy of the seed; fast-check's numbers are reduced into the seed's own ranges. */
function damage(seed: Uint8Array, d: Damage): Uint8Array {
  if (d.kind === "truncate") return seed.slice(0, d.keep % seed.length);
  if (d.kind === "size") return damageAt(seed, seed, d);
  if (d.kind === "aim") return aimAt(seed, seed, d);
  const out = seed.slice();
  out[d.position % out.length]! ^= 1 << d.bit;
  return out;
}

const damageArb: fc.Arbitrary<Damage> = fc.oneof(
  fc.record({ kind: fc.constant("truncate" as const), keep: fc.nat() }),
  fc.record({ kind: fc.constant("flip" as const), position: fc.nat(), bit: fc.integer({ min: 0, max: 7 }) }),
  fc.record({ kind: fc.constant("aim" as const), entry: fc.nat(), header: fc.nat() }),
  fc.record({
    kind: fc.constant("size" as const),
    field: fc.nat(),
    // A value inside the archive too (the seeds are 2.5 to 7.4 KB): over [0, 2^32) alone an offset almost never lands
    // on another header, nor a size on the next entry (review of 05/10/2026).
    value: fc.oneof(fc.constantFrom(0, 0xffff, 0xffffffff), fc.nat({ max: 16_384 }), fc.nat({ max: 0xffffffff })),
    relative: fc.boolean(),
  }),
);

// ---------------------------------------------------------------------------
// The tests.
// ---------------------------------------------------------------------------

describe("zip reader — differential against jszip: the seeds", () => {
  it("the « aim » damage really points an entry at another entry's local header, and ours refuses it", async () => {
    let aimed = 0;
    for (const seed of SEEDS) {
      const fields = headerOffsetFields(seed);
      if (fields.length < 2) continue;
      const bytes = aimAt(seed, seed, { kind: "aim", entry: 0, header: 1 });
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      expect(view.getUint32(fields[0]!.at, true)).toBe(fields[1]!.offset);
      expect(fields[1]!.offset).not.toBe(fields[0]!.offset);
      // Refused by `readLocalHeaders`: the aimed header carries the other entry's name (the local/central name check),
      // before the overlap check is reached. So this does NOT guard the overlap refusal (checked 05/10/2026: with it
      // disabled this stays green); `local-headers.test.ts` « overlapping-entries bomb » does.
      expect((await ours(bytes)).ok).toBe(false);
      expect(await departures(bytes, ours)).toEqual([]);
      aimed++;
    }
    expect(aimed).toBeGreaterThan(0);
  });

  it("every seed is read the same way by both readers, and the hand-built one really carries a data descriptor", async () => {
    for (const seed of SEEDS) {
      expect(await departures(seed, ours)).toEqual([]);
      const verdict = await ours(seed);
      expect(verdict.ok).toBe(true);
    }
    const handBuilt = readLocalHeaders(HAND_BUILT).entries;
    expect(handBuilt.some((e) => e.hasDataDescriptor)).toBe(true);
    expect(handBuilt.some((e) => e.method === 0 && !e.dir)).toBe(true);
  });

  it("the comparison bites: a reader that accepts everything is caught on a truncated archive", async () => {
    // Without this, a `departures` that always answered [] would make every property below pass.
    const acceptsAll: Reader = () => Promise.resolve({ ok: true, names: [], data: new Map() });
    const truncated = SEEDS[0]!.slice(0, 100);
    expect(await departures(truncated, acceptsAll)).toEqual(["accepts a structure jszip refuses"]);
  });

  it("the comparison bites on data too: bytes jszip refuses, or other bytes, are both caught", async () => {
    const seed = HAND_BUILT;
    const real = await theirs(seed);
    if (!real.ok) throw new Error("seed refused by jszip");
    const altered = new Map(real.data);
    altered.set("word/document.xml", text("other bytes"));
    const otherBytes: Reader = () => Promise.resolve({ ...real, data: altered });
    expect(await departures(seed, otherBytes)).toEqual(["entry 1: bytes differ"]);

    // 0xff 0xff: block type 3, reserved (§5.5.2): jszip refuses the data, a reader that returns bytes is caught.
    const garbage = build([
      { name: "a.xml", data: text("abc"), deflate: true, compressed: new Uint8Array([0xff, 0xff]) },
    ]);
    expect(refusesEntry(await theirs(garbage), "a.xml")).toBe(true);
    const claimsBytes: Reader = () =>
      Promise.resolve({ ok: true, names: ["a.xml"], data: new Map([["a.xml", text("abc")]]) });
    expect(await departures(garbage, claimsBytes)).toEqual(["entry 0: accepts data jszip refuses"]);
  });
});

describe("zip reader — differential against jszip: damaged archives", () => {
  it("the fuzz bites: a reader that skips the local headers (slice 1 alone) is caught", async () => {
    // Central names and no data: jszip extracts with the LOCAL name, so a flipped local byte must show up.
    const centralOnly: Reader = (bytes) => {
      try {
        const names = readCentralDirectory(bytes)
          .entries.map((e) => e.name)
          .sort();
        return Promise.resolve({ ok: true, names, data: new Map() });
      } catch {
        return Promise.resolve({ ok: false, stage: "structure" });
      }
    };
    const result = await fc.check(
      fc.asyncProperty(fc.nat({ max: SEEDS.length - 1 }), damageArb, async (seed, d) => {
        expect(await departures(damage(SEEDS[seed]!, d), centralOnly)).toEqual([]);
      }),
      { numRuns: 300, seed: 20261005 },
    );
    expect(result.failed).toBe(true);
  });

  it("one damage: never accepted by ours when jszip refuses it, never read differently", async () => {
    let runs = 0;
    let accepted = 0;
    await fc.assert(
      fc.asyncProperty(fc.nat({ max: SEEDS.length - 1 }), damageArb, async (seed, d) => {
        const bytes = damage(SEEDS[seed]!, d);
        runs++;
        // Counted when bytes are really compared: at least one entry whose data ours read.
        const mine = await ours(bytes);
        if (mine.ok && [...mine.data.values()].some((b) => b !== undefined)) accepted++;
        expect(await departures(bytes, ours)).toEqual([]);
      }),
      { numRuns: fuzzRuns(300, process.env) },
    );
    // Not vacuous: a property over archives whose data ours always refuses would hold trivially. 10 % is the floor
    // (measured 05/10/2026 by the counter-review: 32 to 41 % over 8 × 300 runs, before the « aim » damage was added).
    expect(accepted).toBeGreaterThan(runs / 10);
  });

  it("two or three damages in a row: same rule", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.nat({ max: SEEDS.length - 1 }),
        fc.array(damageArb, { minLength: 2, maxLength: 3 }),
        async (seed, ds) => {
          // Each later damage is aimed with the ORIGINAL seed's field map, on purpose: offsets may then miss, which is damage too.
          let bytes = SEEDS[seed]!;
          for (const d of ds) {
            if (d.kind === "size") bytes = damageAt(bytes, SEEDS[seed]!, d);
            else if (d.kind === "aim") bytes = aimAt(bytes, SEEDS[seed]!, d);
            else bytes = damage(bytes, d);
          }
          expect(await departures(bytes, ours)).toEqual([]);
        },
      ),
      { numRuns: fuzzRuns(150, process.env) },
    );
  });
});

/** Where each central record's local-header offset field sits in a VALID seed, and the offset it holds. */
function headerOffsetFields(seed: Uint8Array): { at: number; offset: number }[] {
  const view = new DataView(seed.buffer, seed.byteOffset, seed.byteLength);
  const cd = readCentralDirectory(seed);
  const out: { at: number; offset: number }[] = [];
  let at = cd.offset;
  for (let i = 0; i < cd.count; i++) {
    out.push({ at: at + 42, offset: view.getUint32(at + 42, true) });
    at += 46 + view.getUint16(at + 28, true) + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
  }
  return out;
}

function aimAt(bytes: Uint8Array, seed: Uint8Array, d: Damage & { kind: "aim" }): Uint8Array {
  const fields = headerOffsetFields(seed);
  if (fields.length === 0) return bytes;
  const { at } = fields[d.entry % fields.length]!;
  if (at + 4 > bytes.length) return bytes;
  const out = bytes.slice();
  new DataView(out.buffer).setUint32(at, fields[d.header % fields.length]!.offset, true);
  return out;
}

/** A size damage aimed with `seed`'s field map but written into `bytes` (same length unless truncated earlier). */
function damageAt(bytes: Uint8Array, seed: Uint8Array, d: Damage & { kind: "size" }): Uint8Array {
  const fields = sizeFields(seed);
  const { at, width } = fields[d.field % fields.length]!;
  if (at + width > bytes.length) return bytes;
  const out = bytes.slice();
  const view = new DataView(out.buffer);
  const max = width === 2 ? 0xffff : 0xffffffff;
  const old = width === 2 ? view.getUint16(at, true) : view.getUint32(at, true);
  const value = d.relative ? Math.min(max, Math.max(0, old + (d.value % 33) - 16)) : d.value % (max + 1);
  if (width === 2) view.setUint16(at, value, true);
  else view.setUint32(at, value, true);
  return out;
}

describe("zip reader — the empty deflated entry (plan, slice 4: « jszip non vérifié »)", () => {
  // APPNOTE 6.3.10 §4.3.8: « Zero-byte files, directories, and other file types that contain no content MUST NOT
  // include file data. » A zero-byte entry marked deflated (method 8) with 0 compressed bytes follows that sentence.
  const EMPTY_DEFLATED = build([
    { name: "word/empty.xml", data: new Uint8Array(), deflate: true, compressed: new Uint8Array() },
  ]);

  // Measured 05/10/2026 (the first version of this test expected both to refuse, and failed on jszip): jszip reads it
  // as 0 bytes. Decided at integration the same day: ours reads it as 0 bytes too (`zip-package.ts`), since it follows
  // the specification and a real document carrying one must not be refused after the switch away from jszip.
  it("both read it as 0 bytes", async () => {
    const mine = await ours(EMPTY_DEFLATED);
    const other = await theirs(EMPTY_DEFLATED);
    expect(mine.ok && mine.data.get("word/empty.xml")).toEqual(new Uint8Array());
    expect(other.ok && other.data.get("word/empty.xml")).toEqual(new Uint8Array());
    expect(await departures(EMPTY_DEFLATED, ours)).toEqual([]);
  });
});
