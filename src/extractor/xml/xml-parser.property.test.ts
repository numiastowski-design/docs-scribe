/**
 * A differential fuzz of the INTERFACE module (`xml-parser.ts`) against the reader under it, on mutated real parts: the
 * call shape adds nothing and loses nothing.
 *
 * - `XmlParser.parse(text)` gives exactly what `buildTree(text, "document")` gives (same tree, or the same refusal),
 *   except where its one rule of its own applies: text whose declaration names an ASCII-compatible encoding other than
 *   UTF-8, and holds only ASCII, which `buildTree` refuses (« declared encoding does not match the bytes ») and the
 *   module reads as the bytes it was (the declared encoding);
 *   and accented text under a Windows-1252 declaration without U+FFFD, which the module reads as it stands (« prefer UTF-8 »);
 * - `XmlParser.parse(bytes)` gives exactly what `decode.ts` then `buildTree` give;
 * - `XmlFragmentParser.parse(text)` gives exactly what `buildTree(text, "fragment", …, MAX_STORAGE_XHTML_DEPTH)` gives;
 * - any exception but our own refusal (`ReaderError`, `EXTRACT_FAILED`, `xml: …`) is a crash, and fails.
 *
 * Seeds: for each kind of part (its name with digits dropped), the smallest one under 8 KiB. fast-check chooses the
 * seed, the damage (truncation, an inserted fragment, a rewritten declaration) and where it lands (no bias).
 */
import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { MAX_XML_DEPTH } from "../ooxml/xml";
import { DECLARATION_MAX_BYTES, decodeXmlPart } from "./decode";
import { buildTree, MAX_XML_NODES } from "./tree";
import { MAX_STORAGE_XHTML_DEPTH, XmlFragmentParser, XmlParser } from "./xml-parser";
import { fixtureXmlParts, productParsers, type XNode } from "./xml-fixtures.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(60_000, process.env) });

const SEED_MAX_CHARS = 8 * 1024;
const MISMATCH = "xml: declared encoding does not match the bytes";

let seeds: string[] = [];
let part: XmlParser;
let fragment: XmlFragmentParser;

beforeAll(async () => {
  const { docxReaderOptions, fragmentCallerOptions } = await productParsers();
  part = new XmlParser(docxReaderOptions);
  fragment = new XmlFragmentParser(fragmentCallerOptions);
  const best = new Map<string, string>();
  await fixtureXmlParts((p) => {
    if (p.text.length > SEED_MAX_CHARS) return false;
    const kind = p.id.slice(p.id.indexOf("!") + 1).replace(/[0-9]+/g, "");
    const known = best.get(kind);
    if (known === undefined || p.text.length < known.length) best.set(kind, p.text);
    return false;
  });
  seeds = [...best.values()];
}, 300_000);

/** A tree as a string, or our refusal's literal; anything else thrown propagates (a crash fails the property). */
const run = (read: () => XNode[]): string => {
  try {
    return JSON.stringify(read());
  } catch (error) {
    if (error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: "))
      return `refused: ${error.message}`;
    throw error;
  }
};

const SNIPPETS = [
  "<",
  ">",
  "&",
  "&nbsp;",
  "&#0;",
  "<!--x-->",
  "<?pi x?>",
  "<![CDATA[x]]>",
  "<!DOCTYPE a>",
  "</w:p>",
  "<w:p>",
  ' a="1"',
  "\uFEFF",
  "\u00e9",
  "\uFFFD",
  "\r\n",
  "<zz:q/>",
];
const ENCODINGS = [
  "UTF-8",
  "utf-8",
  "Windows-1252",
  "ISO-8859-1",
  "us-ascii",
  "Shift_JIS",
  "UTF-16",
  "UTF-16LE",
  "x-unknown",
];

type Damage =
  | { kind: "truncate"; at: number }
  | { kind: "insert"; at: number; snippet: string }
  | { kind: "declare"; encoding: string };

const damage = fc.oneof(
  fc.record({ kind: fc.constant("truncate" as const), at: fc.noBias(fc.double({ min: 0, max: 1, noNaN: true })) }),
  fc.record({
    kind: fc.constant("insert" as const),
    at: fc.noBias(fc.double({ min: 0, max: 1, noNaN: true })),
    snippet: fc.constantFrom(...SNIPPETS),
  }),
  fc.record({ kind: fc.constant("declare" as const), encoding: fc.constantFrom(...ENCODINGS) }),
);

const DECLARATION = /^(\uFEFF?)<\?xml[^>]*\?>/;

const apply = (text: string, d: Damage): string => {
  if (d.kind === "declare") {
    const declaration = `<?xml version="1.0" encoding="${d.encoding}"?>`;
    return DECLARATION.test(text) ? text.replace(DECLARATION, `$1${declaration}`) : declaration + text;
  }
  const at = Math.floor(d.at * text.length);
  return d.kind === "truncate" ? text.slice(0, at) : text.slice(0, at) + d.snippet + text.slice(at);
};

const caseArb = fc.tuple(fc.noBias(fc.nat()), fc.array(damage, { minLength: 1, maxLength: 3 }));

describe("xml-parser — the interface module adds nothing to the reader and loses nothing (mutated real parts)", () => {
  it("has seeds", () => {
    expect(seeds.length).toBeGreaterThan(20);
  });

  it("XmlParser.parse(text) is buildTree(text), but ASCII text under another declared encoding", () => {
    let retried = 0;
    let windows1252 = 0;
    fc.assert(
      fc.property(caseArb, ([pick, damages]) => {
        const text = damages.reduce(apply, seeds[pick % seeds.length]!);
        const mine = run(() => part.parse(text));
        const tree = run(() => buildTree(text, "document", "UTF-8", MAX_XML_NODES, MAX_XML_DEPTH));
        if (tree === `refused: ${MISMATCH}` && /^[\x00-\x7f]*$/.test(text)) {
          // ASCII text: read as the bytes it was (its refusal, if any, is then the reader's own on those bytes).
          retried++;
          const bytes = new TextEncoder().encode(text);
          expect(mine).toBe(
            run(() => buildTree(text, "document", decodeXmlPart(bytes).encoding, MAX_XML_NODES, MAX_XML_DEPTH)),
          );
          return;
        }
        if (tree === `refused: ${MISMATCH}` && !text.includes("\uFFFD")) {
          // Accented text under a Windows-1252 declaration (« prefer UTF-8 », review of 26/10/2026): read as it stands.
          let declared: string | undefined;
          try {
            declared = decodeXmlPart(new TextEncoder().encode(text.slice(0, DECLARATION_MAX_BYTES))).encoding;
          } catch {
            declared = undefined;
          }
          if (declared === "windows-1252") {
            windows1252++;
            expect(mine).toBe(run(() => buildTree(text, "document", "windows-1252", MAX_XML_NODES, MAX_XML_DEPTH)));
            return;
          }
        }
        expect(mine).toBe(tree);
      }),
      { numRuns: fuzzRuns(2_000, process.env) },
    );
    expect(retried, "the ASCII rule was never reached: the fuzz proves nothing about it").toBeGreaterThan(0);
    expect(windows1252, "the Windows-1252 rule was never reached: the fuzz proves nothing about it").toBeGreaterThan(0);
  });

  it("XmlParser.parse(bytes) is decode.ts then buildTree", () => {
    fc.assert(
      fc.property(caseArb, ([pick, damages]) => {
        const bytes = new TextEncoder().encode(damages.reduce(apply, seeds[pick % seeds.length]!));
        const expected = run(() => {
          const { text, encoding } = decodeXmlPart(bytes);
          return buildTree(text, "document", encoding, MAX_XML_NODES, MAX_XML_DEPTH);
        });
        expect(run(() => part.parse(bytes))).toBe(expected);
      }),
      { numRuns: fuzzRuns(1_000, process.env) },
    );
  });

  it("XmlFragmentParser.parse(text) is buildTree(text, fragment) at MAX_STORAGE_XHTML_DEPTH", () => {
    fc.assert(
      fc.property(caseArb, ([pick, damages]) => {
        const text = damages.reduce(apply, seeds[pick % seeds.length]!);
        expect(run(() => fragment.parse(text))).toBe(
          run(() => buildTree(text, "fragment", "UTF-8", MAX_XML_NODES, MAX_STORAGE_XHTML_DEPTH)),
        );
      }),
      { numRuns: fuzzRuns(1_000, process.env) },
    );
  });
});
