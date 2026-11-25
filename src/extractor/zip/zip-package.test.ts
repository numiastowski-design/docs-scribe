/**
 * Slice 5 of the jszip detachment: `openZip`, everything `ooxml/package.ts`
 * asks of jszip — open (or refuse) an archive, list its files, know a file
 * without inflating it, read one under a per-read and a package-wide ceiling.
 *
 * The package-layer tests of `docx-reader.test.ts` (« decompression is
 * bounded », « arbitrage 28/09/2026 ») and of `cfb.test.ts` (« a non-zip byte
 * soup ») are REPLAYED here against the new module, with their own numbers.
 * One rule is new (review of 05/10/2026): bytes produced by an entry that is
 * then refused as damaged still count against the package ceiling.
 * jszip stays the independent reading on every real Office fixture.
 */
import { readFileSync } from "node:fs";
import { crc32, deflateRawSync } from "node:zlib";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { publicFixtures } from "../public-fixtures.test.support";
import { PACKAGE_MAX_BYTES, PART_MAX_BYTES } from "../ooxml/package";
import { openZip, ZipEntryRefusal } from "./zip-package";

// ---------------------------------------------------------------------------
// A hand-built archive: stored or deflated entries, directories, lying sizes.
// ---------------------------------------------------------------------------

interface Entry {
  name: string;
  /** The uncompressed content (absent for a directory). */
  plain?: Uint8Array;
  method?: 0 | 8;
  /** Replaces the compressed bytes written (a broken or truncated stream). */
  raw?: Uint8Array;
  /** Replaces the declared uncompressed size. */
  declaredSize?: number;
}

const u16 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff];
const u32 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function zip(entries: readonly Entry[]): Uint8Array {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const e of entries) {
    const name = utf8(e.name);
    const plain = e.plain ?? new Uint8Array();
    const method = e.method ?? 8;
    const data = e.raw ?? (method === 8 ? new Uint8Array(deflateRawSync(plain)) : plain);
    const common = [
      ...u16(20),
      ...u16(0),
      ...u16(method),
      ...u16(0),
      ...u16(0x21),
      ...u32(crc32(plain)),
      ...u32(data.length),
      ...u32(e.declaredSize ?? plain.length),
      ...u16(name.length),
      ...u16(0),
    ];
    const local = concat([new Uint8Array([...u32(0x04034b50), ...common]), name, data]);
    centrals.push(
      concat([
        new Uint8Array([
          ...u32(0x02014b50),
          ...u16(20),
          ...common,
          ...u16(0),
          ...u16(0),
          ...u16(0),
          ...u32(0),
          ...u32(offset),
        ]),
        name,
      ]),
    );
    locals.push(local);
    offset += local.length;
  }
  const cd = concat(centrals);
  const end = new Uint8Array([
    ...u32(0x06054b50),
    ...u16(0),
    ...u16(0),
    ...u16(entries.length),
    ...u16(entries.length),
    ...u32(cd.length),
    ...u32(offset),
    ...u16(0),
  ]);
  return concat([...locals, cd, end]);
}

/** Deterministic bytes that deflate hardly compresses (a 32-bit LCG). */
function noise(length: number, seed: number): Uint8Array {
  const out = new Uint8Array(length);
  let s = seed >>> 0;
  for (let i = 0; i < length; i++) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
}

async function refusalOf(read: Promise<unknown>): Promise<ZipEntryRefusal> {
  const outcome: unknown = await read.then(
    () => "read",
    (e: unknown) => e,
  );
  if (outcome instanceof ZipEntryRefusal) return outcome;
  throw new Error(`expected a ZipEntryRefusal, got ${outcome instanceof Error ? outcome.name : String(outcome)}`);
}

const LIMITS = { partMaxBytes: 1000, packageMaxBytes: 1_000_000 };

// ---------------------------------------------------------------------------

describe("openZip — what package.ts asks of jszip", () => {
  const bytes = zip([
    { name: "[Content_Types].xml", plain: utf8("<Types/>") },
    { name: "word/", method: 0 },
    { name: "word/document.xml", plain: utf8("<w:document/>") },
    { name: "word/media/a.png", plain: new Uint8Array(10).fill(3), method: 0 },
  ]);

  it("the file names in central directory order, directories left out (openArchive's `names`)", () => {
    expect(openZip(bytes, LIMITS).names).toEqual(["[Content_Types].xml", "word/document.xml", "word/media/a.png"]);
  });

  it("knows a file without inflating it; a directory or an absent name is no file (jszip's `zip.file`)", async () => {
    const pkg = openZip(bytes, { partMaxBytes: 0, packageMaxBytes: 0 });
    expect(pkg.has("word/document.xml")).toBe(true);
    expect(pkg.has("word/")).toBe(false);
    expect(pkg.has("word/missing.xml")).toBe(false);
    expect(await pkg.readBytes("word/missing.xml")).toBeUndefined();
    expect(await pkg.readBytes("word/")).toBeUndefined();
  });

  it("reads stored and deflated files to their exact bytes", async () => {
    const pkg = openZip(bytes, LIMITS);
    expect(new TextDecoder().decode(await pkg.readBytes("word/document.xml"))).toBe("<w:document/>");
    expect(await pkg.readBytes("word/media/a.png")).toEqual(new Uint8Array(10).fill(3));
  });

  it("a stored read is a fresh buffer, never a view on the archive", async () => {
    const copy = bytes.slice();
    const out = (await openZip(copy, LIMITS).readBytes("word/media/a.png"))!;
    expect(out.buffer).not.toBe(copy.buffer);
    out.fill(9);
    expect(copy).toEqual(bytes);
  });

  it("an empty deflated entry with no stream reads as 0 bytes (APPNOTE 6.3.10 §4.3.8, as jszip)", async () => {
    const pkg = openZip(zip([{ name: "empty.xml", raw: new Uint8Array() }]), LIMITS);
    expect(await pkg.readBytes("empty.xml")).toEqual(new Uint8Array());
  });

  it("a ceiling that is not a safe integer >= 0 is a caller's bug, thrown at once", () => {
    expect(() => openZip(bytes, { partMaxBytes: Number.NaN, packageMaxBytes: 1 })).toThrow(/limits/);
    expect(() => openZip(bytes, { partMaxBytes: 1, packageMaxBytes: -1 })).toThrow(/limits/);
  });
});

describe("openZip — replayed from cfb.test.ts: what is not a zip", () => {
  it("a non-zip byte soup, an empty input and a truncated archive are refused at open, as ReaderError, never as an entry refusal", () => {
    const good = zip([{ name: "a.xml", plain: utf8("<a/>") }]);
    for (const input of [
      utf8("this is not a zip at all, just words"),
      new Uint8Array(),
      good.subarray(0, good.length - 5),
    ]) {
      let thrown: unknown;
      try {
        openZip(input, LIMITS);
      } catch (e) {
        thrown = e;
      }
      expect(thrown).toBeInstanceOf(ReaderError);
      expect(thrown).not.toBeInstanceOf(ZipEntryRefusal);
      expect((thrown as ReaderError).code).toBe("EXTRACT_FAILED");
    }
  });

  it("the refusal echoes no byte of the archive", () => {
    const secret = "Confidential-Name-Of-A-Client";
    const bytes = zip([{ name: secret, plain: utf8("x") }]);
    // Break the end record's entry count: refused by the central directory reader.
    bytes[bytes.length - 12] = 9;
    expect(() => openZip(bytes, LIMITS)).toThrow(ReaderError);
    try {
      openZip(bytes, LIMITS);
    } catch (e) {
      expect((e as Error).message).not.toContain(secret);
    }
  });
});

describe("openZip — replayed from docx-reader.test.ts « decompression is bounded »", () => {
  const media = (n: number): Entry => ({ name: `word/media/m${n}.png`, plain: new Uint8Array(800).fill(n) });
  const bytes = zip([media(1), media(2), media(3), { name: "word/media/big.png", plain: new Uint8Array(1001) }]);

  it("each read counts against the part ceiling; exactly the ceiling is accepted", async () => {
    const pkg = openZip(bytes, LIMITS);
    expect((await refusalOf(pkg.readBytes("word/media/big.png"))).kind).toBe("ceiling");
    expect((await pkg.readBytes("word/media/m1.png"))?.length).toBe(800);
    expect((await pkg.readBytes("word/media/m1.png", 800))?.length).toBe(800);
  });

  it("a caller's tighter budget for one read (the pictures' ceiling) is a ceiling too", async () => {
    const pkg = openZip(bytes, LIMITS);
    expect((await refusalOf(pkg.readBytes("word/media/m1.png", 799))).kind).toBe("ceiling");
  });

  it("sizeOf gives the declared size without inflating, and nothing for an absent file", async () => {
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2000 });
    expect(pkg.sizeOf("word/media/m1.png")).toBe(800);
    expect(pkg.sizeOf("word/media/big.png")).toBe(1001);
    expect(pkg.sizeOf("word/media/absent.png")).toBeUndefined();
  });

  it("a ceiling refusal says WHICH ceiling: the read's own allowance, or the package's", async () => {
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2000 });
    expect((await refusalOf(pkg.readBytes("word/media/m1.png", 799))).reason).toBe("read");
    expect((await refusalOf(pkg.readBytes("word/media/big.png"))).reason).toBe("read");
    // 800 + 1001 spent, then 800 more: the package ceiling is passed whatever the read's own allowance says.
    expect((await refusalOf(pkg.readBytes("word/media/m1.png", 799))).reason).toBe("package");
  });

  it("every read counts against the package ceiling", async () => {
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2 * 800 + 400 });
    expect((await pkg.readBytes("word/media/m1.png"))?.length).toBe(800);
    expect((await pkg.readBytes("word/media/m2.png"))?.length).toBe(800);
    expect((await refusalOf(pkg.readBytes("word/media/m3.png"))).kind).toBe("ceiling");
  });

  it("a part read twice counts twice", async () => {
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2 * 800 + 400 });
    await pkg.readBytes("word/media/m1.png");
    await pkg.readBytes("word/media/m1.png");
    expect((await refusalOf(pkg.readBytes("word/media/m1.png"))).kind).toBe("ceiling");
  });

  it("a bomb (70 MB of zeros, a few KB of zip) is a ceiling under the real per-part ceiling", async () => {
    const bomb = zip([{ name: "word/document.xml", plain: new Uint8Array(70_000_000) }]);
    expect(bomb.length).toBeLessThan(200_000);
    const pkg = openZip(bomb, { partMaxBytes: PART_MAX_BYTES, packageMaxBytes: PACKAGE_MAX_BYTES });
    expect((await refusalOf(pkg.readBytes("word/document.xml"))).kind).toBe("ceiling");
  });
});

describe("openZip — moved from inflate-entry.test.ts (05/10/2026, when inflateEntry left): the declared size is never trusted", () => {
  const LIMITS = { partMaxBytes: 1_000_000, packageMaxBytes: 10_000_000 };

  it("a bomb that declares 10 bytes is refused as a ceiling", async () => {
    const bytes = zip([{ name: "word/secret-bomb.xml", plain: new Uint8Array(2_000_000), declaredSize: 10 }]);
    const refusal = await refusalOf(openZip(bytes, LIMITS).readBytes("word/secret-bomb.xml"));
    expect(refusal.kind).toBe("ceiling");
    expect(refusal.message).not.toContain("secret-bomb");
  });

  it("a small entry that declares a size past the ceiling is damaged, not a bomb: the declaration alone refuses nothing", async () => {
    const bytes = zip([{ name: "word/secret-liar.xml", plain: new Uint8Array(100).fill(7), declaredSize: 50_000_000 }]);
    const refusal = await refusalOf(openZip(bytes, LIMITS).readBytes("word/secret-liar.xml"));
    expect(refusal.kind).toBe("damaged");
    expect(refusal.message).not.toContain("secret-liar");
  });

  it("a stored entry whose declared size differs from its data is damaged", async () => {
    const bytes = zip([{ name: "a.xml", plain: new Uint8Array(10).fill(1), method: 0, declaredSize: 12 }]);
    expect((await refusalOf(openZip(bytes, LIMITS).readBytes("a.xml"))).kind).toBe("damaged");
  });

  it("bytes that are no deflate stream are damaged", async () => {
    // Block type 3 is reserved (RFC 1951 §3.2.3): no inflater accepts it.
    const bytes = zip([{ name: "a.xml", plain: new Uint8Array(10), raw: Uint8Array.of(0x07, 0x00, 0x00, 0x00) }]);
    expect((await refusalOf(openZip(bytes, LIMITS).readBytes("a.xml"))).kind).toBe("damaged");
  });
});

describe("openZip — replayed from docx-reader.test.ts « arbitrage 28/09/2026 »: which refusal", () => {
  it("a broken stream is `damaged` (degradable); a part over its ceiling is `ceiling` (never)", async () => {
    const bytes = zip([
      { name: "word/media/broken.png", raw: new Uint8Array([0xff, 0xff, 0xff, 0xff]), declaredSize: 1000 },
      { name: "word/media/big.png", plain: new Uint8Array(2000).fill(7) },
    ]);
    const pkg = openZip(bytes, LIMITS);
    const broken = await refusalOf(pkg.readBytes("word/media/broken.png"));
    expect(broken.kind).toBe("damaged");
    expect(broken.message).toBe("zip: entry data is damaged");
    expect((await refusalOf(pkg.readBytes("word/media/big.png"))).kind).toBe("ceiling");
  });
});

describe("openZip — new rule (review of 05/10/2026): a refused entry's bytes are not free", () => {
  it("the bytes inflated by an entry whose declared size lies still count against the package ceiling", async () => {
    // 600 bytes really inflate, then the size check refuses them as damaged.
    const bytes = zip([
      { name: "liar.xml", plain: new Uint8Array(600).fill(1), declaredSize: 5 },
      { name: "honest.xml", plain: new Uint8Array(500).fill(2) },
    ]);
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 1000 });
    expect((await refusalOf(pkg.readBytes("liar.xml"))).kind).toBe("damaged");
    // 600 + 500 > 1000: refused; free damage would have let it through.
    expect((await refusalOf(pkg.readBytes("honest.xml"))).kind).toBe("ceiling");
  });

  it("the bytes inflated before a truncated stream breaks count too", async () => {
    const plain = noise(300_000, 7);
    const full = new Uint8Array(deflateRawSync(plain));
    const bytes = zip([
      { name: "cut.bin", plain, raw: full.subarray(0, 150_000) },
      { name: "next.bin", plain: new Uint8Array(100_000).fill(4), method: 0 },
    ]);
    const pkg = openZip(bytes, { partMaxBytes: 1_000_000, packageMaxBytes: 200_000 });
    expect((await refusalOf(pkg.readBytes("cut.bin"))).kind).toBe("damaged");
    // ~150 KB were produced before the break: 100 KB more no longer fit.
    expect((await refusalOf(pkg.readBytes("next.bin"))).kind).toBe("ceiling");
    // Witness: the same read with the damaged one never attempted fits.
    expect(
      (await openZip(bytes, { partMaxBytes: 1_000_000, packageMaxBytes: 200_000 }).readBytes("next.bin"))?.length,
    ).toBe(100_000);
  });

  it("a stored entry refused past its read ceiling is charged up to that ceiling, as jszip's stream was, not its whole length", async () => {
    // google-drawings.ts swallows a refused part and reads on: a 2 000-byte stored part refused under a 1 000-byte
    // ceiling must not spend 2 000 of the package budget (review of 05/10/2026).
    const bytes = zip([
      { name: "big.png", plain: new Uint8Array(2000).fill(5), method: 0 },
      { name: "next.png", plain: new Uint8Array(1000).fill(6), method: 0 },
    ]);
    const pkg = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2500 });
    expect((await refusalOf(pkg.readBytes("big.png"))).kind).toBe("ceiling");
    expect((await pkg.readBytes("next.png"))?.length).toBe(1000);
    // And never free: charged at least just past the ceiling (1 001), so 1 000 more no longer fit under 2 000.
    const tight = openZip(bytes, { partMaxBytes: 1000, packageMaxBytes: 2000 });
    expect((await refusalOf(tight.readBytes("big.png"))).kind).toBe("ceiling");
    expect((await refusalOf(tight.readBytes("next.png"))).kind).toBe("ceiling");
  });
});

// ---------------------------------------------------------------------------
// The independent reading: jszip on real documents nobody here wrote.
// ---------------------------------------------------------------------------

/** The public Office documents (`publicFixtures`), as paths from the root of the repository. The measurement corpora
 *  versed on 14/10/2026 (25 Word, 36 workbooks) measure headings, not the zip reader; with them this walk passed its 20 s
 *  deadline. */
function officeFixtures(): string[] {
  return publicFixtures(".docx", ".pptx", ".xlsx").filter(
    (path) => !path.startsWith("fixtures/docx-heading-reel/") && !path.startsWith("fixtures/xlsx-heading-reel/"),
  );
}

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((x, i) => x === b[i]);

describe("openZip — agrees with jszip", () => {
  const fixtures = officeFixtures();

  it("the fixture corpus is not empty (a silent zero would prove nothing)", () => {
    expect(fixtures.length).toBeGreaterThan(20);
  });

  it(
    "on every real Office fixture: the same file names, the same `has`, the same bytes for every part, or both refuse",
    async () => {
      const disagreements: string[] = [];
      let parts = 0;
      for (const path of fixtures) {
        const bytes = new Uint8Array(readFileSync(path));
        let theirs: JSZip | undefined;
        try {
          theirs = await JSZip.loadAsync(bytes, { createFolders: false });
        } catch {
          theirs = undefined;
        }
        let ours: ReturnType<typeof openZip> | undefined;
        try {
          ours = openZip(bytes, { partMaxBytes: PART_MAX_BYTES, packageMaxBytes: PACKAGE_MAX_BYTES });
        } catch (e) {
          if (!(e instanceof ReaderError)) throw e;
        }
        if (!theirs || !ours) {
          if (theirs || ours) disagreements.push(`${path}: one reader refused the archive`);
          continue;
        }
        const theirNames = Object.keys(theirs.files).filter((n) => !theirs.files[n]!.dir);
        if (JSON.stringify(theirNames) !== JSON.stringify(ours.names)) disagreements.push(`${path}: names`);
        for (const name of Object.keys(theirs.files)) {
          if ((theirs.file(name) !== null) !== ours.has(name)) disagreements.push(`${path}: has ${name}`);
        }
        for (const name of theirNames) {
          const a = await theirs
            .file(name)!
            .async("uint8array")
            .then(
              (b) => b,
              () => undefined,
            );
          const b = await ours.readBytes(name).then(
            (x) => x,
            (e: unknown) => {
              if (e instanceof ZipEntryRefusal) return undefined;
              throw e;
            },
          );
          parts++;
          if (a === undefined || b === undefined ? a !== b : !sameBytes(a, b)) disagreements.push(`${path}: ${name}`);
        }
      }
      expect(disagreements).toEqual([]);
      expect(parts).toBeGreaterThan(200);
    },
    // A local derogation, like `pdf-bundle.test.ts` (deleted, see git history): it inflates every part of every real document, twice.
    20_000,
  );
});

describe("ZipEntryRefusal — the code each class carries (lot « fichier abîmé », 25/10/2026)", () => {
  it("🔴 a ceiling is FILE_UNSAFE (a bomb: nothing is proposed), damage stays EXTRACT_FAILED", () => {
    expect(new ZipEntryRefusal("ceiling").code).toBe("FILE_UNSAFE");
    expect(new ZipEntryRefusal("damaged").code).toBe("EXTRACT_FAILED");
  });
});
