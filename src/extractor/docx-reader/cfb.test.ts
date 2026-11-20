/**
 * A password-protected Office file is NOT a zip: it is a CFB compound file
 * ([MS-CFB]) whose directory carries the `EncryptionInfo` and
 * `EncryptedPackage` streams ([MS-OFFCRYPTO] §2.3.4.4 — the same discriminant
 * Microsoft's own `EncryptedPackageEnvelope.IsEncryptedPackageEnvelope` uses).
 * Until 28/09/2026 the reader refused it as « damaged », a lie to the admin:
 * the file is intact, it is protected (28/09: name it, and tell it
 * apart from the OTHER CFB case — a legacy binary `.doc`/`.xls`/`.ppt`, whose
 * directory instead carries `WordDocument`/`Workbook`/`PowerPoint Document`).
 *
 * The corpus holds no protected specimen, and fabricating one would mean
 * writing an encryptor — so, exactly like `pdf-decrypt.test.ts` falsifies
 * `/U`, these tests FABRICATE minimal CFB containers: header, one FAT
 * sector, one directory sector. What the classifier sees on a real file is
 * exactly these structures.
 */
import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { ReaderError } from "doc-ir";
import { openPackage } from "../ooxml/package";
import { classifyCfb } from "./cfb";

const SECTOR = 512;
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;

/** A minimal CFB v3 container: root entry plus the given streams, all empty.
 *  One FAT sector (sector 0), one directory sector (sector 1). */
export function cfbWith(streamNames: string[]): Uint8Array {
  if (streamNames.length > 3) throw new Error("one directory sector holds root + 3 entries");
  const bytes = new Uint8Array(SECTOR * 3);
  const view = new DataView(bytes.buffer);

  // Header ([MS-CFB] §2.2).
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version 3 → 512-byte sectors
  view.setUint16(28, 0xfffe, true); // byte order
  view.setUint16(30, 0x0009, true); // sector shift
  view.setUint16(32, 0x0006, true); // mini sector shift
  view.setUint32(44, 1, true); // one FAT sector
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, 0x1000, true); // mini stream cutoff
  view.setUint32(60, ENDOFCHAIN, true); // no mini FAT
  view.setUint32(68, ENDOFCHAIN, true); // no DIFAT chain
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, i === 0 ? 0 : FREESECT, true);

  // FAT (sector 0): itself, then the single-sector directory chain.
  const fat = SECTOR;
  view.setUint32(fat, FATSECT, true);
  view.setUint32(fat + 4, ENDOFCHAIN, true);
  for (let i = 2; i < SECTOR / 4; i++) view.setUint32(fat + i * 4, FREESECT, true);

  // Directory (sector 1): 128-byte entries ([MS-CFB] §2.6.1).
  const writeEntry = (index: number, name: string, objectType: number) => {
    const at = SECTOR * 2 + index * 128;
    for (let i = 0; i < name.length; i++) view.setUint16(at + i * 2, name.charCodeAt(i), true);
    view.setUint16(at + 64, (name.length + 1) * 2, true); // length counts the nul
    view.setUint8(at + 66, objectType); // 5 root, 2 stream, 0 unused
    view.setUint8(at + 67, 1); // black
    view.setUint32(at + 68, FREESECT, true); // no left sibling
    view.setUint32(at + 72, FREESECT, true); // no right sibling
    view.setUint32(at + 76, objectType === 5 && streamNames.length > 0 ? 1 : FREESECT, true);
    view.setUint32(at + 116, ENDOFCHAIN, true); // no stream data
  };
  writeEntry(0, "Root Entry", 5);
  streamNames.forEach((name, i) => writeEntry(i + 1, name, 2));
  return bytes;
}

const refusalOf = (bytes: Uint8Array): Promise<ReaderError | "opened"> =>
  openPackage(bytes).then(
    () => "opened" as const,
    (e) => e as ReaderError,
  );

describe("openPackage — a CFB container is named for what it is, never « damaged » (28/09/2026)", () => {
  it("🔴 the encrypted-package streams refuse as PASSWORD-PROTECTED, definitively", async () => {
    const error = await refusalOf(cfbWith(["EncryptionInfo", "EncryptedPackage"]));
    expect(error).toBeInstanceOf(ReaderError);
    expect((error as ReaderError).code).toBe("DOCUMENT_PASSWORD_PROTECTED");
  });

  it("🔴 a legacy binary Word file (WordDocument stream) refuses as UNSUPPORTED, not damaged", async () => {
    const error = await refusalOf(cfbWith(["WordDocument", "1Table"]));
    expect((error as ReaderError).code).toBe("UNSUPPORTED_TYPE");
  });

  it("🔴 legacy Excel and PowerPoint streams classify the same way", async () => {
    expect(((await refusalOf(cfbWith(["Workbook"]))) as ReaderError).code).toBe("UNSUPPORTED_TYPE");
    expect(((await refusalOf(cfbWith(["PowerPoint Document"]))) as ReaderError).code).toBe("UNSUPPORTED_TYPE");
  });

  it("a CFB with none of the known streams stays « damaged » — an unknown container decides nothing", async () => {
    const error = await refusalOf(cfbWith(["SomethingElse"]));
    expect((error as ReaderError).code).toBe("EXTRACT_FAILED");
  });

  it("a non-zip, non-CFB byte soup stays « damaged », exactly as before", async () => {
    const error = await refusalOf(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]));
    expect((error as ReaderError).code).toBe("EXTRACT_FAILED");
  });

  it("a CFB truncated right after its signature stays « damaged », never crashes", async () => {
    const error = await refusalOf(cfbWith(["EncryptedPackage"]).slice(0, 100));
    expect((error as ReaderError).code).toBe("EXTRACT_FAILED");
  });

  it(
    "🔴 a CYCLIC DIFAT chain answers « other » fast — never accumulates until the hop ceiling (revue HIGH 28/09 soir, constat 1)",
    { timeout: 5_000 },
    async () => {
      // v4 (4096-byte sectors): a DIFAT sector whose next-pointer loops onto
      // itself. Before the fix each of the 102 400 allowed hops pushed ~1023
      // entries (~105M array cells) before answering — a hostile-input cost on
      // bytes that already failed the zip open. A visited set answers at the
      // second visit. ⚠️ The OUTCOME is « other » before and after — only the
      // work changes, and V8 absorbed the old walk under this timeout — so
      // this test pins the classification, and the cost guard is verified by
      // construction (`cfb.ts`, visited sets + maxSectors), not by outcome.
      const sectorSize = 4096;
      const bytes = new Uint8Array(512 + sectorSize * 2);
      const view = new DataView(bytes.buffer);
      bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
      view.setUint16(26, 0x0004, true);
      view.setUint16(28, 0xfffe, true);
      view.setUint16(30, 0x000c, true); // 4096-byte sectors
      view.setUint32(48, 1, true); // directory at sector 1 (never reached: the DIFAT loops first)
      view.setUint32(68, 0, true); // first DIFAT sector: 0
      for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, FREESECT, true);
      // DIFAT sector 0: every entry a plausible FAT sector number (1), the
      // next-pointer (last field) → itself. Each hop of the old walk pushed
      // all 1023 entries.
      const difatAt = sectorSize; // sector 0 starts after one sector-size header slot
      for (let i = 0; i < sectorSize / 4 - 1; i++) view.setUint32(difatAt + i * 4, 1, true);
      view.setUint32(difatAt + sectorSize - 4, 0, true); // the cycle

      const error = await refusalOf(bytes);
      expect((error as ReaderError).code).toBe("EXTRACT_FAILED"); // « other »: the refusal stays « damaged »
    },
  );

  // 17/11/2026, decided: a value that is no sector number, written where the DIFAT names a FAT sector, reads as an EMPTY entry,
  // wherever it sits. Before, the answer depended on its place: kept as a sector in the first entry, it sent the directory walk
  // out of the file (« other »), while after the real FAT sector it changed nothing (« encrypted »). One file, one answer: the
  // protected document is said protected, the only thing its owner can act on.
  it.each([
    ["DIFSECT", 0xfffffffc],
    ["FATSECT", FATSECT],
    ["the reserved value above the last sector number", 0xfffffffb],
  ])(
    "%s in the header DIFAT reads as an empty entry, before or after the real FAT sector: « encrypted » both ways",
    (_name, marker) => {
      const before = cfbWith(["EncryptionInfo", "EncryptedPackage"]);
      const view = new DataView(before.buffer);
      view.setUint32(76, marker, true);
      view.setUint32(80, 0, true);
      expect(classifyCfb(before)).toBe("encrypted");

      const after = cfbWith(["EncryptionInfo", "EncryptedPackage"]);
      new DataView(after.buffer).setUint32(80, marker, true);
      expect(classifyCfb(after)).toBe("encrypted");
    },
  );

  it("the last regular sector number is still a sector: out of this small file, « other »", () => {
    // The witness of the bound: 0xFFFFFFFA is a sector number, not a marker. It is kept, looked up, and found out of the file.
    const bytes = cfbWith(["EncryptionInfo", "EncryptedPackage"]);
    const view = new DataView(bytes.buffer);
    view.setUint32(76, 0xfffffffa, true);
    view.setUint32(80, 0, true);
    expect(classifyCfb(bytes)).toBe("other");
  });

  it("🔴 a REAL protected document classifies the same as the fabricated one", async () => {
    // `07-Fiche-Contact.docx` encrypted by msoffcrypto-tool 28/09/2026 with
    // the deliberately public password « miksumia-test » — agile encryption,
    // the shape Word itself writes. Before this lot the refusal said
    // « damaged »; the two expected values differ, so the fixture exercises
    // the classification, not just the refusal. It lives under `refuses/`:
    // a document whose EXPECTED outcome is a refusal has no published words,
    // so the conservation sweeps of `fixtures/docx/` must not see it.
    const bytes = await readFile("fixtures/docx/refuses/09-Fiche-Contact-Protegee.docx");
    const error = await refusalOf(new Uint8Array(bytes));
    expect((error as ReaderError).code).toBe("DOCUMENT_PASSWORD_PROTECTED");
  });
});
