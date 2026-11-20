/**
 * Namespace well-formedness of an Office part (W3C Namespaces in XML 1.0, Third Edition), CHECKS ONLY. The tree keeps
 * every name as written; the rewriting of prefixes stays in `ooxml/xml.ts#canonicalize`, our own code.
 * - A part that is namespace-well-formed is read.
 * - A part that breaks a namespace constraint is well-formed XML 1.0: its refusal is ours, with its own literal.
 * - Storage XHTML (fragment mode) is never checked: Confluence's `ac:` and `ri:` prefixes are never declared there.
 */
import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { buildTree, type XmlMode } from "./tree";

const QNAME = "xml: malformed qualified name";
const REPEATED = "xml: attributes repeated under their namespace";
const REASONS = [QNAME, REPEATED];

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
const W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const refusalOf = (source: string, mode: XmlMode = "document"): string => {
  try {
    buildTree(source, mode);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return (error as ReaderError).message;
  }
  throw new Error("expected a refusal");
};

const REFUSED: readonly (readonly [string, string, string])[] = [
  // §4 QName: Prefix ':' LocalPart, both NCNames — one colon at most, neither side empty, LocalPart starting as an NCName.
  ["two colons in an element name", `<a:b:c xmlns:a="u"/>`, QNAME],
  ["a colon first in an element name", `<:a/>`, QNAME],
  ["a colon last in an element name", `<a:/>`, QNAME],
  ["two colons in an attribute name", `<r xmlns:a="u" a:b:c="1"/>`, QNAME],
  ["a colon first in an attribute name", `<r :x="1"/>`, QNAME],
  ["a local part starting with a digit (a valid XML Name, not an NCName)", `<p:1a xmlns:p="u"/>`, QNAME],
  ["a local part starting with '-'", `<r xmlns:p="u" p:-a="1"/>`, QNAME],
  ["a local part starting with '.'", `<p:.a xmlns:p="u"/>`, QNAME],
  ["a local part starting with a combining mark", `<p:\u0301a xmlns:p="u"/>`, QNAME],
  ["a declaration attribute whose local part (the prefix) is no NCName", `<r xmlns:1p="u"/>`, QNAME],
  // §6.3 NSC Attributes Unique, matched on the product's canonical names (`ooxml/xml.ts#canonicalize`): two
  // attributes that become one key lose a value there.
  [
    "two prefixes, one canonical namespace, one local name",
    `<a xmlns:p="${W}" xmlns:q="${W}" p:x="1" q:x="2"/>`,
    REPEATED,
  ],
  [
    "the same, one prefix bound on an ancestor",
    `<r xmlns:p="${W}"><w:a xmlns:w="${W}" w:val="1" p:val="2"/></r>`,
    REPEATED,
  ],
  [
    "Transitional w:val and Strict s:val: one canonical name",
    `<a xmlns:w="${W}" xmlns:s="${W_STRICT}" w:val="1" s:val="2"/>`,
    REPEATED,
  ],
  [
    "a spreadsheetml prefix is stripped: s:val and val are one key",
    `<a xmlns:s="${SHEET}" s:val="1" val="2"/>`,
    REPEATED,
  ],
];

describe("xml namespaces — refused with our literal: a name no reader recognises, or a value the product would lose", () => {
  it.each(REFUSED)("%s", (_label, source, reason) => {
    expect(refusalOf(source)).toBe(reason);
  });

  it("each reason is a literal namespaces.ts raises", () => {
    const source = readFileSync(new URL("./namespaces.ts", import.meta.url), "utf8");
    const raised = new Set([...source.matchAll(/xmlRefusal\(\s*"([^"]+)"/g)].map((m) => `xml: ${m[1]}`));
    expect(REASONS.filter((reason) => !raised.has(reason))).toEqual([]);
  });
});

/** Namespace-well-formed parts, and what is read from each: the exact tree. */
const ACCEPTED: readonly (readonly [string, string, unknown])[] = [
  [
    "the prefix xml, bound without a declaration",
    `<a xml:space="preserve" xml:lang="fr">t</a>`,
    [{ a: [{ "#text": "t" }], ":@": { "xml:space": "preserve", "xml:lang": "fr" } }],
  ],
  [
    "the prefix xml declared to its own name",
    `<a xmlns:xml="${XML_NS}" xml:space="preserve"/>`,
    [{ a: [], ":@": { "xmlns:xml": "http://www.w3.org/XML/1998/namespace", "xml:space": "preserve" } }],
  ],
  [
    "an empty default namespace (§6.2), which 94 versioned parts write",
    `<a xmlns="u"><b xmlns=""><c/></b></a>`,
    [{ a: [{ b: [{ c: [] }], ":@": { xmlns: "" } }], ":@": { xmlns: "u" } }],
  ],
  [
    "a declaration applying to its own element's name and attributes",
    `<w:p xmlns:w="${W}" w:rsidR="00"/>`,
    [
      {
        "w:p": [],
        ":@": { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:rsidR": "00" },
      },
    ],
  ],
  [
    "a prefix redeclared to another name on a child, back to the first after it",
    `<p:a xmlns:p="u"><p:b xmlns:p="v"/><p:c p:x="1"/></p:a>`,
    [
      {
        "p:a": [
          { "p:b": [], ":@": { "xmlns:p": "v" } },
          { "p:c": [], ":@": { "p:x": "1" } },
        ],
        ":@": { "xmlns:p": "u" },
      },
    ],
  ],
  [
    "one local name under two namespaces",
    `<r xmlns:p="u" xmlns:q="v" p:x="1" q:x="2"/>`,
    [{ r: [], ":@": { "xmlns:p": "u", "xmlns:q": "v", "p:x": "1", "q:x": "2" } }],
  ],
  [
    "an unprefixed attribute is in no namespace, even beside the default's prefix",
    `<r xmlns="u" xmlns:p="u" x="1" p:x="2"/>`,
    [{ r: [], ":@": { xmlns: "u", "xmlns:p": "u", x: "1", "p:x": "2" } }],
  ],
  [
    "prefixes starting with xml other than xml and xmlns (§3: not a fatal error)",
    `<xmlfoo:a xmlns:xmlfoo="u" xmlfoo:x="1"/>`,
    [{ "xmlfoo:a": [], ":@": { "xmlns:xmlfoo": "u", "xmlfoo:x": "1" } }],
  ],
  ["an element named xmlns without a prefix", `<xmlns/>`, [{ xmlns: [] }]],
  [
    "a non-ASCII prefix and local part",
    `<é:à xmlns:é="u" é:ü="1"/>`,
    [{ "é:à": [], ":@": { "xmlns:é": "u", "é:ü": "1" } }],
  ],
  [
    "a local part with digits, '-', '.' after its first character",
    `<p:a1-b.c xmlns:p="u"/>`,
    [{ "p:a1-b.c": [], ":@": { "xmlns:p": "u" } }],
  ],
  [
    "a processing instruction target with a colon (dropped by both callers; §7 not enforced: only what would publish wrong is refused, 25/10/2026)",
    `<a><?p:i d?></a>`,
    [{ a: [] }],
  ],
  // Read, though not namespace-valid: nothing would publish wrong (product owner's rule, 25/10/2026).
  ["an element prefix bound nowhere", `<p:a/>`, [{ "p:a": [] }]],
  ["an attribute prefix bound nowhere", `<a p:x="1"/>`, [{ a: [], ":@": { "p:x": "1" } }]],
  [
    "a prefix bound on a closed element, out of scope",
    `<r><a xmlns:p="u"></a><b p:x="1"/></r>`,
    [
      {
        r: [
          { a: [], ":@": { "xmlns:p": "u" } },
          { b: [], ":@": { "p:x": "1" } },
        ],
      },
    ],
  ],
  ["the prefix xmlns declared", `<a xmlns:xmlns="u"/>`, [{ a: [], ":@": { "xmlns:xmlns": "u" } }]],
  ["an element named with the prefix xmlns", `<xmlns:a/>`, [{ "xmlns:a": [] }]],
  ["the prefix xml bound to another name", `<a xmlns:xml="u"/>`, [{ a: [], ":@": { "xmlns:xml": "u" } }]],
  [
    "another prefix bound to the name of xml",
    `<a xmlns:p="${XML_NS}"/>`,
    [{ a: [], ":@": { "xmlns:p": "http://www.w3.org/XML/1998/namespace" } }],
  ],
  [
    "the default namespace bound to the name of xml",
    `<a xmlns="${XML_NS}"/>`,
    [{ a: [], ":@": { xmlns: "http://www.w3.org/XML/1998/namespace" } }],
  ],
  [
    "a prefix bound to the name of xmlns",
    `<a xmlns:p="${XMLNS_NS}"/>`,
    [{ a: [], ":@": { "xmlns:p": "http://www.w3.org/2000/xmlns/" } }],
  ],
  ['a prefix undeclared (xmlns:p="")', `<a xmlns:p=""/>`, [{ a: [], ":@": { "xmlns:p": "" } }]],
  [
    "two unbound prefixes, one local name: two keys in the product",
    `<a p:x="1" q:x="2"/>`,
    [{ a: [], ":@": { "p:x": "1", "q:x": "2" } }],
  ],
  [
    "two prefixes of one UNKNOWN namespace: the product keeps both names",
    `<a xmlns:p="u" xmlns:q="u" p:x="1" q:x="2"/>`,
    [{ a: [], ":@": { "xmlns:p": "u", "xmlns:q": "u", "p:x": "1", "q:x": "2" } }],
  ],
  [
    "two prefixes of two canonical namespaces: two keys",
    `<a xmlns:p="${W}" xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/main" p:x="1" q:x="2"/>`,
    [
      {
        a: [],
        ":@": {
          "xmlns:p": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "xmlns:q": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "p:x": "1",
          "q:x": "2",
        },
      },
    ],
  ],
  [
    "a markup-compatibility part as Word writes it",
    `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body><w:p w14:paraId="1"/></w:body></w:document>`,
    [
      {
        "w:document": [{ "w:body": [{ "w:p": [], ":@": { "w14:paraId": "1" } }] }],
        ":@": {
          "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
          "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
          "mc:Ignorable": "w14",
        },
      },
    ],
  ],
];

describe("xml namespaces — what a namespace-well-formed part holds is read", () => {
  it.each(ACCEPTED)("%s", (_label, source, tree) => {
    expect(buildTree(source, "document")).toEqual(tree);
  });

  it("two attributes of one local name stay two keys, as written, whether their prefixes are unbound, share a namespace or sit in two", () => {
    const drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main";
    expect(buildTree(`<a xmlns:p="${W}" xmlns:q="${drawingml}" p:x="1" q:x="2"/>`, "document")).toEqual([
      { a: [], ":@": { "xmlns:p": W, "xmlns:q": drawingml, "p:x": "1", "q:x": "2" } },
    ]);
    expect(buildTree(`<a p:x="1" q:x="2"/>`, "document")).toEqual([{ a: [], ":@": { "p:x": "1", "q:x": "2" } }]);
    expect(buildTree(`<a xmlns:p="u" xmlns:q="u" p:x="1" q:x="2"/>`, "document")).toEqual([
      { a: [], ":@": { "xmlns:p": "u", "xmlns:q": "u", "p:x": "1", "q:x": "2" } },
    ]);
  });

  it("the names stay as written: no prefix is resolved or rewritten in the tree", () => {
    expect(buildTree(`<p:a xmlns:p="${W}" p:x="1"/>`, "document")).toEqual([
      { "p:a": [], ":@": { "xmlns:p": W, "p:x": "1" } },
    ]);
  });
});

describe("xml namespaces — storage XHTML (fragment mode) is not checked", () => {
  it("Confluence's undeclared ac: and ri: prefixes are read", () => {
    expect(
      buildTree(`<p>x</p><ac:image ac:width="10"><ri:attachment ri:filename="a.png"/></ac:image>`, "fragment"),
    ).toEqual([
      { p: [{ "#text": "x" }] },
      { "ac:image": [{ "ri:attachment": [], ":@": { "ri:filename": "a.png" } }], ":@": { "ac:width": "10" } },
    ]);
  });
});
