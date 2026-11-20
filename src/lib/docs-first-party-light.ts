/**
 * THE LOAD-FREE FIRST-PARTY ENTRY OF THE DECLARED-FORMATS PACKAGE: what the publisher's own app takes from it beyond
 * the two published doors, when the name loads no reader. NOT a promise to third parties — a name here may change or
 * leave without notice.
 *
 * WHY THESE NAMES ARE NOT IN THE APP. They are small helpers the readers own and the app borrows instead of keeping a
 * second copy: the MIME tables, the XML fragment helpers, the UTF-8 length, the word split, and the slide-image port
 * the PowerPoint reader declares. The app's light functions take this entry, so it stays free of every reader.
 *
 * Decides: which load-free names of the package the app may take without their being public;
 * `src/app-library-surface.test.ts` pins the list and refuses a name nobody takes.
 * NOT here: the published doors (`./docs.ts`, `./docs-light.ts`), what loads a reader (`./docs-first-party.ts`), what
 * a browser bundle embeds (`./docs-first-party-browser.ts`), nor any logic — every name is re-exported from the
 * module that owns it.
 * Called by: `src/readers/light.ts`, `src/readers/extractor.ts`; `ACCEPTED_FORMATS_TEXT` by the dashboard's test that
 * holds its own sentence equal to this one.
 */

export { ACCEPTED_FORMATS_TEXT } from "../extractor/mime";
export type { MimeClassification } from "../extractor/mime";
export { EXT_BY_MIME, GOOGLE_FOLDER, SOURCE_TYPE_BY_EXTENSION, listedMimeType } from "./mime-types";
export type { SlideImage, SlideImagePort, SlidePage } from "./slide-image-port";
export { utf8Length } from "./utf8";
export { wordsOf } from "./words";
export { decodeEntities, element, escapeAttr, escapeText, parseXmlFragment, serializeXmlFragment, text } from "./xml";
export type { XmlNode } from "./xml";
