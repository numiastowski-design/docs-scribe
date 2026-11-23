/**
 * Cost of the chart frames of a sheet in the structure reading (`source-structure-sheet.ts#chartShapesOf`): a drawing whose
 * frames all point at one heavy chart part reads that part once, and a sheet with more frames than the bound abstains
 * (`unreadable`) instead of counting wrongly. Against wall-clock bounds (starting values to calibrate).
 * The workbooks are made here, in memory: one sheet, one row, a drawing of N frames pointing at one shared chart part, or
 * at N distinct copies of it.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { isDeclaredFormatSource } from "./conservation-source";
import { MAX_SHEET_CHART_FRAMES } from "./source-structure-sheet";
import { readXlsxConservationSourceUnder } from "./source-reading-xlsx";

const BOUND_MS = 2_000;

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

/** A bar chart of `series` series over `points` categories (the page publishes a table of `points + 1` rows). */
const barChart = (series: number, points: number): string => {
  const cats = Array.from(
    { length: points },
    (_, p) => `<c:pt idx="${String(p)}"><c:v>Cat${String(p)}x</c:v></c:pt>`,
  ).join("");
  const values = Array.from(
    { length: points },
    (_, p) => `<c:pt idx="${String(p)}"><c:v>${String(p + 7)}</c:v></c:pt>`,
  ).join("");
  const one = (s: number): string =>
    `<c:ser><c:idx val="${String(s)}"/><c:order val="${String(s)}"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Serie${String(s)}x</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:cat><c:strRef><c:f>S!$A$2</c:f><c:strCache><c:ptCount val="${String(points)}"/>${cats}</c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>S!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${String(points)}"/>${values}</c:numCache></c:numRef></c:val></c:ser>`;
  return (
    `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>` +
    `<c:plotArea><c:barChart>${Array.from({ length: series }, (_, s) => one(s)).join("")}</c:barChart></c:plotArea></c:chart></c:chartSpace>`
  );
};

const chartAnchor = (relId: string): string =>
  `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;

/** One sheet holding one row of text, and a drawing of `frames` frames on `chart` (one shared part, or one copy per frame). */
const workbookWithChartFrames = async (chart: string, frames: number, distinct: boolean): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>`),
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData>` +
      `<row r="1"><c r="A1" t="inlineStr"><is><t>w1xgrid</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    RELS(`<Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing1.xml"/>`),
  );
  const ids = Array.from({ length: frames }, (_, k) => (distinct ? `rId${String(k + 1)}` : "rId1"));
  const targetOf = (k: number): string => (distinct ? `../charts/chart1_${String(k)}.xml` : "../charts/chart1.xml");
  zip.file(
    "xl/drawings/drawing1.xml",
    `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">` +
      ids.map(chartAnchor).join("") +
      "</xdr:wsDr>",
  );
  const used = distinct ? ids.map((_, k) => k) : [0];
  zip.file(
    "xl/drawings/_rels/drawing1.xml.rels",
    RELS(used.map((k) => `<Relationship Id="${ids[k] ?? ""}" Type="${R}/chart" Target="${targetOf(k)}"/>`).join("")),
  );
  for (const k of used) zip.file(`xl/${targetOf(k).slice(3)}`, chart);
  return zip.generateAsync({ type: "uint8array" });
};

const readTimed = async (frames: number, distinct = false) => {
  const bytes = await workbookWithChartFrames(barChart(distinct ? 4 : 8, distinct ? 100 : 400), frames, distinct);
  const start = performance.now();
  const source = await readXlsxConservationSourceUnder(bytes, undefined);
  return { source, ms: performance.now() - start };
};

describe("chart frames of a sheet", () => {
  it(
    `${String(MAX_SHEET_CHART_FRAMES)} frames on one 8 x 400 chart are read in under 2 s: the chart part is read once, each frame adds its table`,
    { timeout: 60_000 },
    async () => {
      const { source, ms } = await readTimed(MAX_SHEET_CHART_FRAMES);
      if (!isDeclaredFormatSource(source)) throw new Error("expected a declared-format source");
      // FAILS WITH: seconds (every frame inflates and parses the chart part again: 1 000 frames took 157 s).
      expect(ms).toBeLessThan(BOUND_MS);
      const added = source.structure?.units.reduce((sum, unit) => sum + unit.added.tables, 0);
      // One table per frame, plus the sheet's own grid (the same count the app pinned before this test moved here).
      expect(added).toBe(MAX_SHEET_CHART_FRAMES + 1);
    },
  );

  it(
    "1 000 frames are past the bound: the structure reading is refused in under 2 s, never a wrong count",
    { timeout: 60_000 },
    async () => {
      const { source, ms } = await readTimed(1_000);
      // FAILS WITH: a timeout (157 s), or a declared-format source whose count was computed from 1 000 frames.
      expect(ms).toBeLessThan(BOUND_MS);
      // The structure alone abstains: the words are read all the same (a read failure would refuse the whole document).
      expect(isDeclaredFormatSource(source)).toBe(true);
      expect(isDeclaredFormatSource(source) && source.structure?.unreadable).toBe(true);
    },
  );

  it(
    "1 000 frames on 1 000 DISTINCT chart parts are refused in under 2 s: the frames are counted before any chart part is read",
    { timeout: 120_000 },
    async () => {
      const { source, ms } = await readTimed(1_000, true);
      // FAILS WITH: seconds (every distinct part inflated and parsed before the count is tested).
      expect(ms).toBeLessThan(BOUND_MS);
      expect(isDeclaredFormatSource(source) && source.structure?.unreadable).toBe(true);
    },
  );
});
