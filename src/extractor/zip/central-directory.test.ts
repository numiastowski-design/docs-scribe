/**
 * Slice 1 of the jszip detachment: the central directory of a ZIP archive
 * (names, directory flags, entry count), read without jszip.
 *
 * Two kinds of evidence. Hand-built archives pin each rule of the format
 * (PKWARE APPNOTE 6.3.9 §4.3.12, §4.3.14-§4.3.16, §4.4.4, §4.4.15, §4.4.17)
 * and each hostile shape. Then jszip itself is the second, independent
 * reading: on every real Office fixture of the repository, both readers must
 * agree on names and directory flags, or both refuse.
 */
import { readFileSync } from "node:fs";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { publicFixtures } from "../public-fixtures.test.support";
import { readCentralDirectory } from "./central-directory";

// ---------------------------------------------------------------------------
// A minimal archive builder: central directory records and end records only.
// The reader under test never looks at local headers or entry data, so none
// are written; the central directory starts at `prefix` bytes of filler.
// ---------------------------------------------------------------------------

interface CdRecord {
  name: string | Uint8Array;
  flags?: number;
  madeBy?: number; // high byte of "version made by" (§4.4.2): 0 = MS-DOS, 3 = UNIX
  externalAttributes?: number;
  extra?: Uint8Array;
  comment?: Uint8Array;
}

const u16 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff];
const u32 = (v: number): number[] => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
const u64 = (v: bigint): number[] => {
  const out: number[] = [];
  for (let i = 0n; i < 8n; i++) out.push(Number((v >> (8n * i)) & 0xffn));
  return out;
};
const ascii = (s: string): Uint8Array => new TextEncoder().encode(s);

function cdRecord(r: CdRecord): number[] {
  const name = typeof r.name === "string" ? ascii(r.name) : r.name;
  const extra = r.extra ?? new Uint8Array();
  const comment = r.comment ?? new Uint8Array();
  return [
    ...u32(0x02014b50),
    ...u16(((r.madeBy ?? 0) << 8) | 20), // version made by
    ...u16(20), // version needed
    ...u16(r.flags ?? 0),
    ...u16(0), // method
    ...u16(0), // time
    ...u16(0), // date
    ...u32(0), // crc
    ...u32(0), // compressed size
    ...u32(0), // uncompressed size
    ...u16(name.length),
    ...u16(extra.length),
    ...u16(comment.length),
    ...u16(0), // disk number start
    ...u16(0), // internal attributes
    ...u32(r.externalAttributes ?? 0),
    ...u32(0), // local header offset
    ...name,
    ...extra,
    ...comment,
  ];
}

interface EocdFields {
  disk?: number;
  cdDisk?: number;
  entriesOnDisk?: number;
  entries?: number;
  cdSize?: number;
  cdOffset?: number;
  comment?: Uint8Array;
  commentLength?: number;
}

function eocd(f: Required<Pick<EocdFields, "entries" | "cdSize" | "cdOffset">> & EocdFields): number[] {
  const comment = f.comment ?? new Uint8Array();
  return [
    ...u32(0x06054b50),
    ...u16(f.disk ?? 0),
    ...u16(f.cdDisk ?? 0),
    ...u16(f.entriesOnDisk ?? f.entries),
    ...u16(f.entries),
    ...u32(f.cdSize),
    ...u32(f.cdOffset),
    ...u16(f.commentLength ?? comment.length),
    ...comment,
  ];
}

/** An archive whose central directory holds `records`, with `prefix` filler
 *  bytes before it (standing in for local headers and data). */
function archive(records: CdRecord[], opts: { prefix?: number; eocd?: EocdFields } = {}): Uint8Array {
  const prefix = new Array<number>(opts.prefix ?? 0).fill(0x41);
  const cd = records.flatMap(cdRecord);
  return Uint8Array.from([
    ...prefix,
    ...cd,
    ...eocd({ entries: records.length, cdSize: cd.length, cdOffset: prefix.length, ...opts.eocd }),
  ]);
}

/** The same archive in ZIP64 form (§4.3.14, §4.3.15): every EOCD field that
 *  can be saturated is, and the real values live in the zip64 records. */
function archive64(
  records: CdRecord[],
  opts: { entries64?: bigint; offset64?: bigint; withLocator?: boolean } = {},
): Uint8Array {
  const prefix = new Array<number>(8).fill(0x41);
  const cd = records.flatMap(cdRecord);
  const eocd64At = prefix.length + cd.length;
  const eocd64 = [
    ...u32(0x06064b50),
    ...u64(44n), // size of the remaining record
    ...u16(45), // version made by
    ...u16(45), // version needed
    ...u32(0), // this disk
    ...u32(0), // disk with the central directory
    ...u64(opts.entries64 ?? BigInt(records.length)),
    ...u64(opts.entries64 ?? BigInt(records.length)),
    ...u64(BigInt(cd.length)),
    ...u64(opts.offset64 ?? BigInt(prefix.length)),
  ];
  const locator = [...u32(0x07064b50), ...u32(0), ...u64(BigInt(eocd64At)), ...u32(1)];
  return Uint8Array.from([
    ...prefix,
    ...cd,
    ...eocd64,
    ...(opts.withLocator === false ? [] : locator),
    ...eocd({ entries: 0xffff, cdSize: 0xffffffff, cdOffset: 0xffffffff, disk: 0xffff, cdDisk: 0xffff }),
  ]);
}

function refusal(bytes: Uint8Array): ReaderError {
  try {
    readCentralDirectory(bytes);
  } catch (e) {
    if (e instanceof ReaderError) return e;
    throw e;
  }
  throw new Error("expected a refusal, got a reading");
}

// ---------------------------------------------------------------------------

describe("readCentralDirectory — what a well-formed archive declares", () => {
  it("names, directory flags and count, in central directory order", () => {
    const dir = readCentralDirectory(
      archive([{ name: "word/" }, { name: "word/document.xml" }, { name: "[Content_Types].xml" }], { prefix: 30 }),
    );
    expect(dir.count).toBe(3);
    expect(dir.entries).toEqual([
      { name: "word/", dir: true },
      { name: "word/document.xml", dir: false },
      { name: "[Content_Types].xml", dir: false },
    ]);
  });

  it("an empty archive (EOCD alone) has no entries", () => {
    expect(readCentralDirectory(archive([]))).toEqual({ entries: [], count: 0, offset: 0 });
  });

  it("a directory flagged by the MS-DOS attribute 0x10 only is a directory, named with a trailing slash", () => {
    const dir = readCentralDirectory(
      archive([
        { name: "media", externalAttributes: 0x10 },
        { name: "a.xml", externalAttributes: 0x20 },
      ]),
    );
    expect(dir.entries).toEqual([
      { name: "media/", dir: true },
      { name: "a.xml", dir: false },
    ]);
  });

  it("a directory flagged by the UNIX mode S_IFDIR (made by UNIX) is a directory; the same bits from MS-DOS are not", () => {
    const sIfDir = (0o040755 << 16) >>> 0;
    const dir = readCentralDirectory(
      archive([
        { name: "unixdir", madeBy: 3, externalAttributes: sIfDir },
        { name: "dosfile", madeBy: 0, externalAttributes: sIfDir & 0xffff0000 },
      ]),
    );
    expect(dir.entries).toEqual([
      { name: "unixdir/", dir: true },
      { name: "dosfile", dir: false },
    ]);
  });

  it("a UTF-8 name (general purpose bit 11) is decoded as UTF-8", () => {
    const dir = readCentralDirectory(archive([{ name: ascii("média/été.png"), flags: 1 << 11 }]));
    expect(dir.entries).toEqual([{ name: "média/été.png", dir: false }]);
  });

  it("extra fields and entry comments are skipped by their declared lengths", () => {
    const dir = readCentralDirectory(
      archive([{ name: "a", extra: Uint8Array.from([0x99, 0x99, 2, 0, 7, 7]), comment: ascii("note") }, { name: "b" }]),
    );
    expect(dir.entries.map((e) => e.name)).toEqual(["a", "b"]);
  });

  it("the end record is found behind an archive comment", () => {
    const dir = readCentralDirectory(archive([{ name: "x.xml" }], { eocd: { comment: ascii("archive comment") } }));
    expect(dir.entries).toEqual([{ name: "x.xml", dir: false }]);
  });

  it("reads a view into a larger buffer (non-zero byteOffset) without looking outside it", () => {
    const inner = archive([{ name: "x.xml" }]);
    const outer = new Uint8Array(inner.length + 64).fill(0x50);
    outer.set(inner, 32);
    expect(readCentralDirectory(outer.subarray(32, 32 + inner.length)).entries).toEqual([
      { name: "x.xml", dir: false },
    ]);
  });

  it("ZIP64: count and offset come from the zip64 end record when the EOCD is saturated", () => {
    const dir = readCentralDirectory(archive64([{ name: "a.xml" }, { name: "b/" }]));
    expect(dir.count).toBe(2);
    expect(dir.entries).toEqual([
      { name: "a.xml", dir: false },
      { name: "b/", dir: true },
    ]);
  });
});

describe("readCentralDirectory — hostile archives are refused, by our own words", () => {
  const SECRET = "Confidential salary review.xml";

  it("bytes with no end record", () => {
    const e = refusal(ascii("this is not a zip archive at all, just text long enough"));
    expect(e.code).toBe("EXTRACT_FAILED");
    expect(e.message).toBe("zip: no end of central directory record");
  });

  it("an empty input", () => {
    expect(refusal(new Uint8Array()).message).toBe("zip: no end of central directory record");
  });

  it("a declared count far beyond what the central directory bytes can hold, before any allocation", () => {
    const e = refusal(archive([{ name: "a" }], { eocd: { entries: 0xfffe, entriesOnDisk: 0xfffe } }));
    expect(e.message).toBe("zip: declared entry count exceeds the central directory");
  });

  it("a ZIP64 count beyond what the bytes can hold, and beyond 2^53", () => {
    expect(refusal(archive64([{ name: "a" }], { entries64: 1n << 40n })).message).toBe(
      "zip: declared entry count exceeds the central directory",
    );
    expect(refusal(archive64([{ name: "a" }], { entries64: 1n << 60n })).message).toBe(
      "zip: declared entry count exceeds the central directory",
    );
  });

  it("a central directory declared past the end of the bytes", () => {
    const e = refusal(archive([{ name: "a" }], { eocd: { cdOffset: 0x7fffffff } }));
    expect(e.message).toBe("zip: central directory out of bounds");
  });

  it("a central directory that does not end where the end records begin", () => {
    const e = refusal(archive([{ name: "a" }], { prefix: 10, eocd: { cdOffset: 4 } }));
    expect(e.message).toBe("zip: central directory out of bounds");
  });

  it("a name that runs past the central directory", () => {
    const bytes = archive([{ name: SECRET }]);
    // Declare a name length longer than the record holds: offset 28 of the record.
    bytes[28] = 0xff;
    bytes[29] = 0x00;
    const e = refusal(bytes);
    expect(e.message).toBe("zip: central directory record out of bounds");
    expect(e.message).not.toContain("salary");
  });

  it("fewer records than declared", () => {
    const one = cdRecord({ name: "a" });
    const bytes = Uint8Array.from([...one, ...eocd({ entries: 2, cdSize: one.length, cdOffset: 0 })]);
    expect(refusal(bytes).message).toBe("zip: declared entry count exceeds the central directory");
  });

  it("more records than declared", () => {
    expect(refusal(archive([{ name: "a" }, { name: "b" }], { eocd: { entries: 1 } })).message).toBe(
      "zip: entry count does not match the central directory",
    );
  });

  it("a record without its signature", () => {
    const bytes = archive([{ name: "a" }, { name: "b" }]);
    bytes[0] = 0x00;
    expect(refusal(bytes).message).toBe("zip: bad central directory record signature");
  });

  it("an archive spread over several disks", () => {
    expect(refusal(archive([{ name: "a" }], { eocd: { disk: 1 } })).message).toBe(
      "zip: multi-disk archives are not supported",
    );
    expect(refusal(archive([{ name: "a" }], { eocd: { entriesOnDisk: 0 } })).message).toBe(
      "zip: multi-disk archives are not supported",
    );
  });

  it("an encrypted entry (general purpose bit 0), as jszip does", () => {
    const e = refusal(archive([{ name: SECRET, flags: 1 }]));
    expect(e.message).toBe("zip: encrypted entries are not supported");
    expect(e.message).not.toContain("salary");
  });

  it("two records with the same name: the readers of the archive could disagree on which one counts", () => {
    const e = refusal(archive([{ name: SECRET }, { name: SECRET }]));
    expect(e.message).toBe("zip: duplicate entry name");
    expect(e.message).not.toContain("salary");
  });

  it("a file entry and a directory entry that would take the same name once the slash is added", () => {
    expect(refusal(archive([{ name: "media/" }, { name: "media", externalAttributes: 0x10 }])).message).toBe(
      "zip: duplicate entry name",
    );
  });

  it("two end records that both close the bytes exactly: which one is real cannot be decided", () => {
    const inner = eocd({ entries: 0, cdSize: 0, cdOffset: 0 });
    // The outer comment holds a fake end record whose own comment length
    // reaches exactly the end of the bytes too.
    const fake = [...inner.slice(0, 20), ...u16(0)];
    const real = archive([], { eocd: { comment: Uint8Array.from(fake) } });
    expect(refusal(real).message).toBe("zip: ambiguous end of central directory record");
  });

  it("an end record whose comment length runs past the bytes is not an end record", () => {
    const bytes = archive([{ name: "a" }], { eocd: { commentLength: 5 } });
    expect(refusal(bytes).message).toBe("zip: no end of central directory record");
  });

  it("ZIP64 saturated fields without the zip64 locator", () => {
    expect(refusal(archive64([{ name: "a" }], { withLocator: false })).message).toBe("zip: zip64 end record not found");
  });

  it("ZIP64 offset beyond 2^53 is refused rather than rounded", () => {
    expect(refusal(archive64([{ name: "a" }], { offset64: (1n << 60n) + 8n })).message).toBe(
      "zip: central directory out of bounds",
    );
  });

  it("a name that climbs out of the archive with a `..` segment, wherever it sits", () => {
    for (const name of ["../evil.xml", "word/../../evil.xml", "word/..", "..", "word\\..\\evil.xml"]) {
      const e = refusal(archive([{ name }]));
      expect(e.message, name).toBe("zip: unsafe entry name");
      expect(e.message).not.toContain("evil");
    }
  });

  it("a name jszip would silently rewrite: `.` segment, empty segment, absolute path, backslash, drive letter, NUL", () => {
    for (const name of [
      "./a.xml",
      "word/./a.xml",
      "word//a.xml",
      "/etc/passwd",
      "word\\document.xml",
      "C:/a.xml",
      "a\u0000.xml",
    ]) {
      expect(refusal(archive([{ name }])).message, JSON.stringify(name)).toBe("zip: unsafe entry name");
    }
  });

  it("a `..` entry that is a directory by its MS-DOS attribute only is refused too", () => {
    expect(refusal(archive([{ name: "word/..", externalAttributes: 0x10 }])).message).toBe("zip: unsafe entry name");
  });

  it("names that merely contain dots are not paths that climb", () => {
    const dir = readCentralDirectory(
      archive([{ name: "..a/b..xml" }, { name: "a.../" }, { name: "docProps/core.xml" }]),
    );
    expect(dir.entries.map((e) => e.name)).toEqual(["..a/b..xml", "a.../", "docProps/core.xml"]);
  });

  it("never echoes a name in any refusal of this suite", () => {
    const shapes = [
      archive([{ name: SECRET, flags: 1 }]),
      archive([{ name: SECRET }, { name: SECRET }]),
      archive([{ name: SECRET }], { eocd: { entries: 9 } }),
    ];
    for (const bytes of shapes) expect(refusal(bytes).message).not.toMatch(/salary|Confidential/);
  });
});

// ---------------------------------------------------------------------------
// The independent reading: jszip on real documents nobody here wrote.
// ---------------------------------------------------------------------------

/** The public Office documents (`publicFixtures`), as paths from the root of the repository. */
function officeFixtures(): string[] {
  return publicFixtures(".docx", ".pptx", ".xlsx");
}

type Reading = { kind: "read"; entries: string[] } | { kind: "refused" };

async function jszipReading(bytes: Uint8Array): Promise<Reading> {
  try {
    const zip = await JSZip.loadAsync(bytes, { createFolders: false });
    return {
      kind: "read",
      entries: Object.values(zip.files)
        .map((f) => `${f.dir ? "D" : "F"} ${f.name}`)
        .sort(),
    };
  } catch {
    return { kind: "refused" };
  }
}

function ourReading(bytes: Uint8Array): Reading {
  try {
    return {
      kind: "read",
      entries: readCentralDirectory(bytes)
        .entries.map((e) => `${e.dir ? "D" : "F"} ${e.name}`)
        .sort(),
    };
  } catch (e) {
    if (e instanceof ReaderError) return { kind: "refused" };
    throw e;
  }
}

describe("readCentralDirectory — agrees with jszip", () => {
  const fixtures = officeFixtures();

  it("the fixture corpus is not empty (a silent zero would prove nothing)", () => {
    expect(fixtures.length).toBeGreaterThan(20);
  });

  it("on every real Office fixture: same names, same directory flags, or both refuse", async () => {
    const disagreements: string[] = [];
    let read = 0;
    for (const path of fixtures) {
      const bytes = new Uint8Array(readFileSync(path));
      const theirs = await jszipReading(bytes);
      const ours = ourReading(bytes);
      if (ours.kind === "read") read++;
      if (JSON.stringify(theirs) !== JSON.stringify(ours)) disagreements.push(path);
    }
    expect(disagreements).toEqual([]);
    expect(read).toBeGreaterThan(20);
  });

  const written = (comment: string): Promise<Uint8Array> => {
    const zip = new JSZip();
    zip.folder("word")!.file("document.xml", "<w/>");
    zip.file("média/été.png", new Uint8Array([1, 2, 3]));
    zip.file("[Content_Types].xml", "<Types/>");
    return zip.generateAsync({ type: "uint8array", comment });
  };

  it("on archives jszip writes, with folders, UTF-8 names and a comment", async () => {
    const bytes = await written("an archive comment");
    const theirs = await jszipReading(bytes);
    expect(theirs.kind).toBe("read");
    expect(ourReading(bytes)).toEqual(theirs);
    expect(readCentralDirectory(bytes).count).toBe(Object.keys((await JSZip.loadAsync(bytes)).files).length);
  });

  it("KNOWN DIVERGENCE: a comment holding the EOCD signature — jszip refuses its own archive, this reader does not", async () => {
    // jszip takes the LAST `PK\5\6` in the bytes as the end record, even
    // inside the comment; the end record that closes the bytes exactly is
    // the real one. Observed with jszip 3.10.1 (« End of data reached »).
    const bytes = await written("PK\u0005\u0006 looks like a signature");
    expect(await jszipReading(bytes)).toEqual({ kind: "refused" });
    expect(ourReading(bytes)).toEqual(await jszipReading(await written("an archive comment")));
  });
});
