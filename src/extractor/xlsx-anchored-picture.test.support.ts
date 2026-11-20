/**
 * Test support: a made-up Excel workbook with ONE picture anchored in its only sheet, written part by part with JSZip.
 * Decides: the parts of that workbook, laid out as Excel lays them — the sheet names its drawing by a relationship
 * (`xl/worksheets/_rels/sheet1.xml.rels`), the drawing anchors one `xdr:pic` whose `a:blip r:embed="rId1"` names the
 * image part (`xl/drawings/_rels/drawing1.xml.rels` -> `xl/media/image1.png`), and `[Content_Types].xml` declares the
 * PNG extension. The tests that damage or multiply the picture rewrite exactly these parts.
 * WHY MADE UP. The real workbook these tests were first written on (one drawing, one picture) stays in the publisher's
 * private repository: the application's tests hold it there.
 * NOT here: any assertion about what the reader does with the workbook (the tests), nor the workbooks without a
 * drawing, which each test file builds itself.
 * Called by: `xlsx.test.ts`, `xlsx.image-budget.test.ts`, `xlsx.missing-pieces.test.ts`, `xlsx.read-once.test.ts`,
 * `xlsx.cells-once.test.ts`, `reader-limits.sheets.test.ts`.
 */

import JSZip from "jszip";

import { encodePng } from "doc-ir/internal";

export const DRAWING_PART = "xl/drawings/drawing1.xml";
export const DRAWING_RELS_PART = "xl/drawings/_rels/drawing1.xml.rels";
export const PICTURE_PART = "xl/media/image1.png";
export const RELATIONSHIPS_HEAD =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/** The declaration of the PNG extension, as it stands in `[Content_Types].xml`: a test removes exactly this text. */
export const PNG_CONTENT_TYPE = '<Default Extension="png" ContentType="image/png"/>';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PICTURE_SIDE_PX = 8;

/** The picture: 8 x 8 pixels, a grey ramp, written by our own PNG encoder. */
const picturePng = (): Uint8Array => {
  const pixels: number[] = [];
  for (let index = 0; index < PICTURE_SIDE_PX * PICTURE_SIDE_PX; index += 1) {
    const grey = (index * 4) % 256;
    pixels.push(grey, grey, grey);
  }
  const png = encodePng(PICTURE_SIDE_PX, PICTURE_SIDE_PX, "rgb-24", pixels);
  if (png === undefined) throw new Error("the picture of the made-up workbook could not be encoded");
  return png;
};

const relationships = (inner: string): string => `${DECLARATION}${RELATIONSHIPS_HEAD}${inner}</Relationships>`;

const inlineCell = (reference: string, text: string): string =>
  `<c r="${reference}" t="inlineStr"><is><t>${text}</t></is></c>`;

/** A title alone on its row, a row of three column headings, two rows of values. */
const SHEET_ROWS =
  `<row r="1">${inlineCell("A1", "Mileage report")}</row>` +
  `<row r="2">${inlineCell("A2", "Date")}${inlineCell("B2", "Description")}${inlineCell("C2", "Distance")}</row>` +
  `<row r="3">${inlineCell("A3", "Monday")}${inlineCell("B3", "Visit to a customer")}${inlineCell("C3", "forty-two")}</row>` +
  `<row r="4">${inlineCell("A4", "Tuesday")}${inlineCell("B4", "Training session")}${inlineCell("C4", "seventeen")}</row>`;

const DRAWING_XML =
  `${DECLARATION}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ` +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<xdr:twoCellAnchor editAs="oneCell">' +
  "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
  "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
  '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
  `<xdr:blipFill><a:blip xmlns:r="${OFFICE_RELATIONSHIPS}" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="514350" cy="438150"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
  "</xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";

/** The workbook's archive, open: a test edits its parts before writing it out. */
export const anchoredPictureZip = (): JSZip => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      PNG_CONTENT_TYPE +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      "</Types>",
  );
  zip.file(
    "_rels/.rels",
    relationships(`<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>`),
  );
  zip.file(
    "xl/workbook.xml",
    `${DECLARATION}<workbook xmlns="${MAIN}" xmlns:r="${OFFICE_RELATIONSHIPS}"><sheets>` +
      '<sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>',
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    relationships(`<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/>`),
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `${DECLARATION}<worksheet xmlns="${MAIN}" xmlns:r="${OFFICE_RELATIONSHIPS}"><dimension ref="A1:C4"/>` +
      `<sheetData>${SHEET_ROWS}</sheetData><drawing r:id="rId1"/></worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    relationships(`<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/drawing" Target="../drawings/drawing1.xml"/>`),
  );
  zip.file(DRAWING_PART, DRAWING_XML);
  zip.file(
    DRAWING_RELS_PART,
    relationships(`<Relationship Id="rId1" Type="${IMAGE_RELATIONSHIP}" Target="../media/image1.png"/>`),
  );
  zip.file(PICTURE_PART, picturePng());
  return zip;
};

/** The workbook's bytes, untouched: one sheet, one drawing, one picture. */
export const anchoredPictureWorkbook = async (): Promise<Uint8Array> =>
  new Uint8Array(await anchoredPictureZip().generateAsync({ type: "uint8array" }));
