/**
 * The ceilings the tests of the declared formats (Office, Google) read a document with: one set of destination caps,
 * small enough for a made-up or a real fixture to pass each of them.
 * Decides: the values these tests use, and nothing about a reading: the library holds no destination cap of its own
 * (`doc-ir/lib/reader-limits.ts`: a field left out is `NO_CAP`), a caller sets them, and here the caller is a test.
 * The values are those of a page of the publisher's application on the day the fixtures were written; they are this
 * package's own copy, so that its tests depend on nothing outside it. A value changed here changes what the tests
 * build, never what a consumer gets.
 * NOT here: the reader's own bounds (`doc-ir/lib/reader-limits.ts`), nor the caps of the PDF reader's tests
 * (`pdf-page-caps.test.support.ts`).
 * Called by: the tests of the docx, pptx, xlsx and Google readers and of their second readings.
 */

import type { ReaderLimits } from "doc-ir";

/** Slide tabs one presentation page shows. */
export const MAX_PRESENTATION_TABS = 21;

/** Elements one slide, drawing or SmartArt may hold. */
export const MAX_SLIDE_ELEMENTS = 500;

/** Sheet tabs one workbook page shows. */
export const MAX_WORKBOOK_TABS = 20;

/** Rows of a worksheet tab. */
export const MAX_SHEET_ROWS = 55;

/** Rows of a table of any format. */
export const MAX_TABLE_ROWS = 200;

/** Columns of a table of any format, and of a worksheet tab: one number. */
export const MAX_TABLE_COLS = 21;
export const MAX_SHEET_COLS = MAX_TABLE_COLS;

/** The caps above as a reading takes them, in one object. Two readings of one document (the extraction and its
 *  second reading) are given the SAME object, or they withhold different tables, sheets and slides. */
export const PAGE_SIZED_LIMITS = {
  presentationTabs: MAX_PRESENTATION_TABS,
  slideElements: MAX_SLIDE_ELEMENTS,
  workbookTabs: MAX_WORKBOOK_TABS,
  sheetRows: MAX_SHEET_ROWS,
  sheetCols: MAX_SHEET_COLS,
  tableRows: MAX_TABLE_ROWS,
  tableCols: MAX_TABLE_COLS,
} as const satisfies ReaderLimits;
