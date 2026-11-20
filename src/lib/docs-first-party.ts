/**
 * THE FIRST-PARTY ENTRY OF THE DECLARED-FORMATS PACKAGE THAT LOADS READERS: what the publisher's own app takes from
 * it beyond the two published doors, when the name comes with a zip or an XML reader. NOT a promise to third
 * parties — a name here may change or leave without notice.
 *
 * WHY THESE NAMES ARE NOT IN THE APP. The four second readings read a format, so they live with the reader of that
 * format; the app only wires each one beside the first reading. The bounded XML parser and the three accessors of its
 * nodes are how the app re-reads the cells of a page it published: it must refuse what the readers refuse. The two
 * source types are what the readers accept for Google's own images.
 * WHAT THE PUBLISHER'S TOOLS AND TESTS TAKE, and production does not: the raw second readings (`readDocxSource`,
 * `readPptxSource`, `readXlsxSource`: the text a sweep compares with the page word for word, which no conservation
 * source carries), the numbers Word computes (`readNumberedParagraphs`, `computedNumberOf`), and the package opened
 * under the readers' own ceilings (`openPackage`).
 *
 * Decides: which reader-loading names of the package the app may take without their being public;
 * `src/app-library-surface.test.ts` pins the list and refuses a name nobody takes.
 * NOT here: the published doors (`./docs.ts`, `./docs-light.ts`), the load-free names
 * (`./docs-first-party-light.ts`), nor any logic — every name is re-exported from the module that owns it.
 * Called by: `src/readers/index.ts`, `src/readers/dispatch.ts`; the app's tests, `tools/` and `features/`.
 */

export type { GoogleDrawingSource } from "../extractor/docx-reader";
export { openPackage } from "../extractor/ooxml/package";
export type { GoogleSlidesSource } from "../extractor/pptx";
export { XmlFragmentParser } from "../extractor/xml/xml-parser";
export {
  computedNumberOf,
  readDocxConservationSource,
  readNumberedParagraphs,
  readSource as readDocxSource,
} from "../pipeline/source-reading-docx";
export { readGoogleSheetConservationSource } from "../pipeline/source-reading-gsheet";
export { attrOf, kidsOf, tagOf } from "../pipeline/source-reading-package";
export type { N } from "../pipeline/source-reading-package";
export { readPptxConservationSource, readPptxSource } from "../pipeline/source-reading-pptx";
export { readXlsxConservationSourceUnder, readXlsxSource } from "../pipeline/source-reading-xlsx";
