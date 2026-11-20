/**
 * The OPC inventory of the workbook. The kind of a tab (sheet or chart sheet) and the `veryHidden` state live in
 * `xl/workbook.xml` and its relationships; without them the reader can neither publish a chart sheet nor hold the
 * exposure guard (a `veryHidden` sheet is never published).
 * Witness: fixture 11, a made-up workbook with two visible sheets, a chart sheet and a `veryHidden` sheet.
 */
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { ReaderError } from "doc-ir";
import { findChildren, parseXml, qualifiedNameOf } from "./ooxml/xml";
import { readWorkbookSheets } from "./xlsx-parts";

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(dirname(fileURLToPath(import.meta.url)), "../../fixtures/xlsx", name)));

describe("readWorkbookSheets — order, name, state, NATURE, part path", () => {
  it("🔴 fixture 11: two visible worksheets, a chartsheet, a veryHidden worksheet, in workbook order", async () => {
    const { sheets } = await readWorkbookSheets(fixture("11-FABRIQUE-Ventes-Graphique-Liens.xlsx"));
    expect(sheets.map((s) => ({ name: s.name, state: s.state, kind: s.kind }))).toEqual([
      { name: "Ventes 2026", state: "visible", kind: "worksheet" },
      { name: "Synthèse", state: "visible", kind: "worksheet" },
      { name: "Graphique", state: "visible", kind: "chartsheet" },
      { name: "Paramètres", state: "veryHidden", kind: "worksheet" },
    ]);
    // The part paths are what later steps read (tables, drawings, charts):
    // worksheets and chartsheets live in different folders.
    expect(sheets[0]!.partPath).toMatch(/^xl\/worksheets\//);
    expect(sheets[2]!.partPath).toMatch(/^xl\/chartsheets\//);
  });

  it("byte soup refuses as EXTRACT_FAILED with the Excel wording, same as the grid layer", async () => {
    const error = await readWorkbookSheets(new Uint8Array([1, 2, 3])).then(
      () => undefined,
      (e) => e as ReaderError,
    );
    expect(error?.code).toBe("EXTRACT_FAILED");
    expect(error?.message).toContain("Excel");
  });
});

/** A valid producer may bind a prefix (`<x:worksheet xmlns:x="…spreadsheetml/2006/main">`), and an ECMA-376 Strict file
 *  spells the namespace differently: read by its spelling, the whole inventory would be lost in silence. */
describe("spreadsheetml namespaces canonicalize whatever prefix the producer bound (constat 5)", () => {
  it("🔴 a prefixed workbook and a Strict-URI worksheet read like Excel's own spelling", () => {
    const root = parseXml(
      '<x:sheets xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheet name="A"/></x:sheets>',
    )!;
    expect(qualifiedNameOf(root)).toBe("sheets");
    expect(findChildren(root, "sheet")).toHaveLength(1);

    const strict = parseXml(
      '<w:worksheet xmlns:w="http://purl.oclc.org/ooxml/spreadsheetml/main"><w:sheetData/></w:worksheet>',
    );
    expect(qualifiedNameOf(strict!)).toBe("worksheet");
    expect(findChildren(strict!, "sheetData")).toHaveLength(1);

    const drawing = parseXml(
      '<d:wsDr xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><d:pic/></d:wsDr>',
    );
    expect(qualifiedNameOf(drawing!)).toBe("xdr:wsDr");
    expect(findChildren(drawing!, "xdr:pic")).toHaveLength(1);
  });
});
