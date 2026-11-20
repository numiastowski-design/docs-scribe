import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { isDeclaredFormatSource } from "../pipeline/conservation-source";
import { readGoogleSheetConservationSource } from "../pipeline/source-reading-gsheet";
import { readXlsxConservationSourceUnder } from "../pipeline/source-reading-xlsx";
import { extractGoogleSheet, type GoogleSheetMeta } from "./google-sheet";
import { extractXlsx } from "./xlsx";
import {
  anchoredPictureZip,
  DRAWING_PART,
  DRAWING_RELS_PART,
  IMAGE_RELATIONSHIP,
  PICTURE_PART,
  RELATIONSHIPS_HEAD,
} from "./xlsx-anchored-picture.test.support";

// Plan step 5: a workbook is read under the limits it is given (`ReaderLimits`), the product's when it is given none, and the extraction and its
// independent second reading take the SAME limits. Run: npx vitest run --project=main src/extractor/reader-limits.sheets.test.ts

type Limits = NonNullable<Parameters<typeof extractXlsx>[2]>;

const workbookOf = async (sheets: readonly { name: string; cellsXml: string }[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const RELS = (inner: string): string =>
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(
      sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  sheets.forEach((s, i) =>
    zip.file(
      `xl/worksheets/sheet${i + 1}.xml`,
      `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${s.cellsXml}</sheetData></worksheet>`,
    ),
  );
  return zip.generateAsync({ type: "uint8array" });
};
const letter = (n: number): string => String.fromCharCode(64 + n);
/** `rows` rows of `cols` inline-string cells, the word of each `${tag}${row}${letter}`. */
const gridXml = (tag: string, rows: number, cols: number): string =>
  Array.from(
    { length: rows },
    (_, r) =>
      `<row r="${r + 1}">${Array.from({ length: cols }, (_, c) => `<c r="${letter(c + 1)}${r + 1}" t="inlineStr"><is><t>${tag}${r + 1}${letter(c + 1)}x</t></is></c>`).join("")}</row>`,
  ).join("");

const titlesOf = (doc: { blocks: readonly { kind: string }[] }): string[] => {
  const tabs = doc.blocks[0] as unknown as { kind: string; tabs?: { title: string }[] } | undefined;
  return tabs?.kind === "tabs" ? (tabs.tabs ?? []).map((t) => t.title) : [];
};
const xlsxSource = async (bytes: Uint8Array, limits?: Limits): Promise<{ words: string[]; excluded: number[] }> => {
  const source = await readXlsxConservationSourceUnder(bytes, limits);
  if (!isDeclaredFormatSource(source)) throw new Error("the second reading failed");
  return {
    words: source.units.flatMap((u) => u.tokens.map((t) => t.text)),
    excluded: [...source.excludedUnits].sort((a, b) => a - b),
  };
};

describe("xlsx — the tab, row and column caps are the limits it is given", () => {
  const three = workbookOf([
    { name: "A", cellsXml: gridXml("a", 2, 2) },
    { name: "B", cellsXml: gridXml("b", 2, 2) },
    { name: "C", cellsXml: gridXml("c", 2, 2) },
  ]);
  const tall = workbookOf([
    { name: "Tall", cellsXml: gridXml("t", 4, 2) },
    { name: "Small", cellsXml: gridXml("s", 1, 1) },
  ]);
  const wide = workbookOf([
    { name: "Wide", cellsXml: gridXml("w", 2, 3) },
    { name: "Small", cellsXml: gridXml("s", 1, 1) },
  ]);

  it("workbookTabs: 2 keeps two tabs and names the third in a workbook-cap notice; absent, all three", async () => {
    expect(titlesOf((await extractXlsx(await three)).doc)).toEqual(["A", "B", "C"]);
    // FAILS WITH (limits ignored): expected [ 'A', 'B', 'C' ] to deeply equal [ 'A', 'B' ].
    const limited = await extractXlsx(await three, undefined, { workbookTabs: 2 });
    expect(titlesOf(limited.doc)).toEqual(["A", "B"]);
    expect(limited.notices).toContainEqual({ kind: "workbook-cap", sheets: ["C"], more: 0 });
  });

  it("sheetRows: 3 withholds a 4-row sheet and says so; the next sheet stays", async () => {
    expect(titlesOf((await extractXlsx(await tall)).doc)).toEqual(["Tall", "Small"]);
    // FAILS WITH (limits ignored): expected [ 'Tall', 'Small' ] to deeply equal [ 'Small' ].
    const limited = await extractXlsx(await tall, undefined, { sheetRows: 3 });
    expect(titlesOf(limited.doc)).toEqual(["Small"]);
    expect(limited.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Tall", rows: 4, cols: 2 });
  });

  it("sheetCols: 2 withholds a 3-column sheet and says so", async () => {
    const limited = await extractXlsx(await wide, undefined, { sheetCols: 2 });
    expect(titlesOf(limited.doc)).toEqual(["Small"]);
    expect(limited.notices).toContainEqual({ kind: "sheet-withheld", sheet: "Wide", rows: 2, cols: 3 });
  });

  it("the second reading excludes the same sheets under the SAME limits, and none without", async () => {
    expect((await xlsxSource(await three)).excluded).toEqual([]);
    // FAILS WITH (limits ignored): expected [] to deeply equal [ 2 ].
    expect((await xlsxSource(await three, { workbookTabs: 2 })).excluded).toEqual([2]);
    expect((await xlsxSource(await tall, { sheetRows: 3 })).excluded).toEqual([0]);
    expect((await xlsxSource(await wide, { sheetCols: 2 })).excluded).toEqual([0]);
    expect((await xlsxSource(await tall, { sheetRows: 3 })).words).not.toContain("t4ax");
  });
});

describe("xlsx — the byte allowance of the pictures is the limit it is given", () => {
  // The made-up workbook with one anchored picture (`xlsx-anchored-picture.test.support.ts`), its drawing rewritten with two.
  const twoImages = async (size: number): Promise<Uint8Array> => {
    const zip = anchoredPictureZip();
    const drawing = await zip.file(DRAWING_PART)!.async("string");
    const anchor = /<xdr:twoCellAnchor[\s\S]*<\/xdr:twoCellAnchor>/.exec(drawing)![0];
    expect(anchor).toContain('r:embed="rId1"'); // the rewrite below really changes the workbook
    zip.file(
      DRAWING_PART,
      drawing.replace(
        anchor,
        ["rId1", "rId2"].map((id) => anchor.replace('r:embed="rId1"', `r:embed="${id}"`)).join(""),
      ),
    );
    zip.file(
      DRAWING_RELS_PART,
      `${RELATIONSHIPS_HEAD}${["a", "b"].map((n, i) => `<Relationship Id="rId${i + 1}" Type="${IMAGE_RELATIONSHIP}" Target="../media/${n}.png"/>`).join("")}</Relationships>`,
    );
    zip.remove(PICTURE_PART);
    zip.file("xl/media/a.png", new Uint8Array(size).fill(0));
    zip.file("xl/media/b.png", new Uint8Array(size).fill(1));
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 1 } });
  };

  it("two 30 000-byte pictures are read by default and refused as damaged past payloadBytes: 50 000", async () => {
    const bytes = await twoImages(30_000);
    expect((await extractXlsx(bytes)).images).toHaveLength(2);
    // FAILS WITH (limits ignored): expected 'read' to match object { code: 'EXTRACT_FAILED', … }.
    expect(
      await extractXlsx(bytes, undefined, { payloadBytes: 50_000 }).then(
        () => "read",
        (e: unknown) => e,
      ),
    ).toMatchObject({ code: "EXTRACT_FAILED" });
    expect((await extractXlsx(bytes, undefined, { payloadBytes: 60_000 })).images).toHaveLength(2);
  });
});

describe("Google Sheet — the tab, row and column caps are the limits it is given, in both readings", () => {
  const sheet = (title: string, index: number): GoogleSheetMeta["sheets"][number] => ({
    properties: { title, index, sheetType: "GRID" },
  });
  const rows = (tag: string, n: number, cols: number): string[][] =>
    Array.from({ length: n }, (_, r) => Array.from({ length: cols }, (_, c) => `${tag}${r + 1}${letter(c + 1)}x`));
  const meta: GoogleSheetMeta = { sheets: [sheet("A", 0), sheet("B", 1), sheet("C", 2)] };
  const values = new Map([
    ["A", rows("a", 2, 2)],
    ["B", rows("b", 4, 3)],
    ["C", rows("c", 1, 1)],
  ]);
  const source = (limits?: Limits): { excluded: number[]; words: string[] } => {
    const read = readGoogleSheetConservationSource(meta, values, limits);
    if (!isDeclaredFormatSource(read)) throw new Error("the second reading failed");
    return {
      excluded: [...read.excludedUnits].sort((a, b) => a - b),
      words: read.units.flatMap((u) => u.tokens.map((t) => t.text)),
    };
  };

  it("by default the three tabs are published", async () => {
    expect(titlesOf((await extractGoogleSheet(meta, values)).doc)).toEqual(["A", "B", "C"]);
    expect(source().excluded).toEqual([]);
  });

  it("workbookTabs: 2, sheetRows: 3 and sheetCols: 2 each leave their sheet out, and the source expects none of its words", async () => {
    // FAILS WITH (limits ignored): expected [ 'A', 'B', 'C' ] to deeply equal [ 'A', 'B' ].
    expect(titlesOf((await extractGoogleSheet(meta, values, { workbookTabs: 2 })).doc)).toEqual(["A", "B"]);
    expect(source({ workbookTabs: 2 }).excluded).toEqual([2]);
    const tall = await extractGoogleSheet(meta, values, { sheetRows: 3 });
    expect(titlesOf(tall.doc)).toEqual(["A", "C"]);
    expect(tall.notices).toContainEqual({ kind: "sheet-withheld", sheet: "B", rows: 4, cols: 3 });
    expect(source({ sheetRows: 3 }).excluded).toEqual([1]);
    expect(titlesOf((await extractGoogleSheet(meta, values, { sheetCols: 2 })).doc)).toEqual(["A", "C"]);
    expect(source({ sheetCols: 2 }).excluded).toEqual([1]);
    expect(source({ sheetRows: 3 }).words).not.toContain("b1ax");
  });
});
