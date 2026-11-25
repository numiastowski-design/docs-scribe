/** Test support: the calls to our reader that the `xml/*.test.ts` files share. `eventsOf` and `eventsOfBytes` give the
 *  tokenizer's own flat view of a part (for the tokenizer's tests); `readerAccepts` says whether the whole reader reads a
 *  part (`tree.ts` in document mode, after `decode.ts` when bytes are given; for the fuzzes that pick their seeds among
 *  the parts it reads). The file name keeps « shadow » from the time these calls ran beside another parser, and is not
 *  renamed, by project rule.
 *  NOT here: reading the parts of the fixtures (`xml-fixtures.test.support.ts`), the damage of the fuzzes
 *  (`xml-damage.test.support.ts`), the reading of the reader's sources through the TypeScript parser
 *  (`xml-refusals.test.support.ts`, the only one that loads the compiler), our reader itself (`tokenizer.ts`, `tree.ts`).
 *  Called by: the tests of `src/extractor/xml/`. */

import { ReaderError } from "doc-ir";
import { decodeXmlPart } from "./decode";
import { tokenize, type XmlSourceEncoding, type XmlTokenVisitor } from "./tokenizer";
import { buildTree } from "./tree";
import type { XmlEvent } from "./xml-fixtures.test.support";

/** Our tokenizer's events, a self-closing tag reported as open + close, root-level text dropped. A CDATA section is a
 *  text event of its own, even empty; a processing instruction reports nothing. */
export const eventsOf = (source: string, encoding: XmlSourceEncoding = "UTF-8"): XmlEvent[] => {
  const out: XmlEvent[] = [];
  let depth = 0;
  const visitor: XmlTokenVisitor = {
    open(name, attributes, selfClosing) {
      out.push(["open", name, attributes.map(([k, v]) => [k, v] as const)]);
      if (selfClosing) out.push(["close", name]);
      else depth++;
    },
    close(name) {
      out.push(["close", name]);
      depth--;
    },
    text(value) {
      if (depth > 0) out.push(["text", value]);
    },
    cdata(value) {
      if (depth > 0) out.push(["text", value]);
    },
  };
  tokenize(source, visitor, encoding);
  return out;
};

/** Our events for a part's BYTES: decoded by `decode.ts` (UTF-8 or UTF-16), then tokenized. */
export const eventsOfBytes = (bytes: Uint8Array): XmlEvent[] => {
  const { text, encoding } = decodeXmlPart(bytes);
  return eventsOf(text, encoding);
};

/** Whether our reader reads the part (the bytes through `decode.ts` when given, else the text); only our own refusal
 *  counts as « no », anything else thrown is a crash and propagates. */
export const readerAccepts = (text: string, bytes?: Uint8Array): boolean => {
  try {
    if (bytes === undefined) buildTree(text, "document");
    else {
      const decoded = decodeXmlPart(bytes);
      buildTree(decoded.text, "document", decoded.encoding);
    }
    return true;
  } catch (error) {
    if (error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: "))
      return false;
    throw error;
  }
};
