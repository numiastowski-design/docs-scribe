/**
 * The `preserveOrder` tree the readers walk, built from the tokenizer's events, with the structure checked (XML 1.0
 * §2.1 production [1] document, §3 WFC Element Type Match). What a valid Office part holds is read; what no valid part
 * can hold (tags that do not pair up, a second root, text outside the root) is refused with our own literal.
 * NOT here: the reading of whole published pages (golden pages, real workbooks rendered), which is held by the
 * publisher's application: a page is its renderer's output, not this package's.
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { cdata, element, serializeXml, text as plainText, type XmlNode } from "../../lib/xml";
import { MAX_XML_DEPTH } from "../ooxml/xml";
import { buildTree } from "./tree";
import type { XNode } from "./xml-fixtures.test.support";

const refusalOf = (source: string, mode: "document" | "fragment" = "document"): ReaderError => {
  try {
    buildTree(source, mode);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return error as ReaderError;
  }
  throw new Error("expected a refusal");
};

describe("xml tree — the preserveOrder shape the readers walk", () => {
  it("an element is { name: children, ':@'?: attributes }, attributes in the order written, ':@' after the name", () => {
    const tree = buildTree(`<a x="1" y='2'><b/>t<c z="&amp;"></c></a>`, "document");
    expect(JSON.stringify(tree)).toBe(
      JSON.stringify([{ a: [{ b: [] }, { "#text": "t" }, { c: [], ":@": { z: "&" } }], ":@": { x: "1", y: "2" } }]),
    );
  });

  it("no ':@' key at all when an element has no attribute", () => {
    const [root] = buildTree(`<a><b></b></a>`, "document");
    expect(Object.keys(root!)).toEqual(["a"]);
    expect(Object.keys((root!.a as XNode[])[0]!)).toEqual(["b"]);
  });

  it("text runs: joined across a comment, split by a PI, each CDATA a node of its own even empty", () => {
    expect(buildTree(`<a>x<!--c-->y<?p d?>z<![CDATA[]]><![CDATA[w]]>v</a>`, "document")).toEqual([
      { a: [{ "#text": "xy" }, { "#text": "z" }, { "#text": "" }, { "#text": "w" }, { "#text": "v" }] },
    ]);
  });

  it("drops what both callers drop: the declaration, every PI, root-level white space; keeps white space inside", () => {
    expect(
      buildTree(
        `\uFEFF<?xml version="1.0" encoding="UTF-8"?>\r\n<?mso-application progid="Word"?>\n<a> <b/> </a>\n<!--end-->\n`,
        "document",
      ),
    ).toEqual([{ a: [{ "#text": " " }, { b: [] }, { "#text": " " }] }]);
  });

  // U+FEFF and U+1680 are name characters (§2.3 [4]), not white space: the tree keeps them in names and in character data
  // (`tokenizer.test.ts` pins the same four sources at the event level).
  it("names and character data holding U+FEFF or U+1680 reach the tree as written", () => {
    expect(buildTree(`<a b\uFEFFc="1"/>`, "document")).toEqual([{ a: [], ":@": { "b\uFEFFc": "1" } }]);
    expect(buildTree(`<\uFEFFa/>`, "document")).toEqual([{ "\uFEFFa": [] }]);
    expect(buildTree(`<a \u1680b="1"/>`, "document")).toEqual([{ a: [], ":@": { "\u1680b": "1" } }]);
    expect(buildTree(`<a>x\uFEFFy\u1680</a>`, "document")).toEqual([{ a: [{ "#text": "x\uFEFFy\u1680" }] }]);
  });

  it("white space beside an attribute is kept untrimmed, and a name nested in itself stays three levels", () => {
    const word = `<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t xml:space="preserve"> a </w:t></w:r></w:p></w:body></w:document>`;
    expect(buildTree(word, "document")).toEqual([
      {
        "w:document": [
          { "w:body": [{ "w:p": [{ "w:r": [{ "w:t": [{ "#text": " a " }], ":@": { "xml:space": "preserve" } }] }] }] },
        ],
        ":@": { "xmlns:w": "urn:w" },
      },
    ]);
    expect(buildTree(`<a><a><a/></a></a>`, "document")).toEqual([{ a: [{ a: [{ a: [] }] }] }]);
  });

  // The depth guard stops a part long before the call stack could matter; 100 000 levels are refused
  // with our literal (`guards.test.ts` holds the exact boundary), and the deepest tree the product reads is built.
  it("nests with an explicit stack, never by recursion: the deepest part read builds, 100 000 levels are refused with our literal", () => {
    const depth = MAX_XML_DEPTH + 2;
    let node = buildTree(`${"<a>".repeat(depth - 1)}<a/>${"</a>".repeat(depth - 1)}`, "document")[0];
    let levels = 0;
    while (node) {
      levels++;
      node = (node.a as XNode[])[0];
    }
    expect(levels).toBe(depth);
    expect(refusalOf(`${"<a>".repeat(100_000)}${"</a>".repeat(100_000)}`).message).toBe(
      "xml: elements nested past the depth limit",
    );
  });
});

describe("xml tree — what is not well-formed is refused, never reproduced (XML 1.0 §2.1 [1], §3 WFC Element Type Match)", () => {
  const cases: [string, string][] = [
    [`<a></b>`, "xml: end tag does not match its start tag"],
    [`<a><b></a></b>`, "xml: end tag does not match its start tag"],
    [`<a></A>`, "xml: end tag does not match its start tag"],
    [`</a>`, "xml: end tag without a start tag"],
    [`<a/></a>`, "xml: end tag without a start tag"],
    [`<a>`, "xml: unclosed element"],
    [`<a><b>`, "xml: unclosed element"],
    [`<a/><b/>`, "xml: more than one root element"],
    [`<a></a><a></a>`, "xml: more than one root element"],
    [`x<a/>`, "xml: text outside the root element"],
    [`<a/>x`, "xml: text outside the root element"],
    [`<a/>&amp;`, "xml: text outside the root element"],
    [`<![CDATA[x]]><a/>`, "xml: text outside the root element"],
    [`<a/><![CDATA[ ]]>`, "xml: text outside the root element"],
    [`x`, "xml: text outside the root element"],
  ];

  it.each(cases)("%j is refused with its own literal", (source, message) => {
    const error = refusalOf(source);
    expect(error.code).toBe("EXTRACT_FAILED");
    expect(error.message).toBe(message);
  });

  it("refuses before building past the fault: a mismatch deep in a long part is refused, never returned half-built", () => {
    expect(refusalOf(`<a>${"<b/>".repeat(10_000)}<c></d></a>`).message).toBe(
      "xml: end tag does not match its start tag",
    );
  });
});

// Middle ground (25/10/2026: « il faut un juste milieu entre trop restrictif et trop permissif »): refuse what
// would publish wrong (text lost, gibberish), read as empty what holds no content. A part without any element and
// without any text is not well-formed (§2.1 [1]), but reading it as empty loses nothing; refusing it would
// turn « part ignored » into « document refused ». Text with no element around it stays refused
// (above): reading it would drop that text in silence. Real case: one 0-byte `customXml/item2.xml` (fixtures, corpus).
describe("xml tree — a part with no element and no text is read as empty", () => {
  it.each([
    [""],
    [" \n"],
    ["\uFEFF"],
    ['<?xml version="1.0"?>'],
    ["<!--c-->"],
    ["<?pi x?>"],
    ['\uFEFF<?xml version="1.0"?>\r\n<!--c-->\n'],
  ])("%j gives an empty tree", (source) => {
    expect(buildTree(source, "document")).toEqual([]);
  });
});

// The storage XHTML `src/conservation/xlsx-page-cells.ts` parses is OUR engine's output (`lib/xml.ts#serializeXml`): sibling
// blocks, no single root. Fragment mode keeps every root and drops the text between them, as that caller does.
describe("xml tree — fragment mode, for the storage XHTML of xlsx-page-cells.ts", () => {
  it("several roots, the text between them dropped; tags still have to pair up", () => {
    expect(buildTree(`<p>a</p> x <table><tbody><tr><td colspan="2">1</td></tr></tbody></table>`, "fragment")).toEqual([
      { p: [{ "#text": "a" }] },
      { table: [{ tbody: [{ tr: [{ td: [{ "#text": "1" }], ":@": { colspan: "2" } }] }] }] },
    ]);
    expect(buildTree(``, "fragment")).toEqual([]);
    expect(buildTree(`<?xml version="1.0"?><?p d?><p>a</p>`, "fragment")).toEqual([{ p: [{ "#text": "a" }] }]);
    expect(refusalOf(`<p></td>`, "fragment").message).toBe("xml: end tag does not match its start tag");
    expect(refusalOf(`<p>`, "fragment").message).toBe("xml: unclosed element");
    expect(refusalOf(`</p>`, "fragment").message).toBe("xml: end tag without a start tag");
  });

  it("what serializeXml writes (the four escapes, a CDATA macro body, spans) reads as written", () => {
    const node = (tag: string, attrs: Record<string, string>, children: XmlNode[]): XmlNode =>
      element(tag, attrs, children);
    const text = (value: string, isCdata = false): XmlNode => (isCdata ? cdata(value) : plainText(value));
    const xhtml = [
      node("ac:adf-extension-frame", {}, [
        node("table", { "data-layout": "default" }, [
          node("tbody", {}, [
            node("tr", {}, [
              node("td", { rowspan: "2", colspan: "3" }, [text(`a & b < c > d "e"`)]),
              node("th", {}, []),
            ]),
            node("tr", {}, [node("td", {}, [text("1 234,5")])]),
          ]),
        ]),
      ]),
      node("ac:structured-macro", { "ac:name": "code" }, [
        node("ac:plain-text-body", {}, [text("x <y> ]] & z", true)]),
      ]),
      node("p", {}, [text("é\u00a0😀")]),
    ]
      .map(serializeXml)
      .join("");
    expect(buildTree(xhtml, "fragment")).toEqual([
      {
        "ac:adf-extension-frame": [
          {
            table: [
              {
                tbody: [
                  {
                    tr: [{ td: [{ "#text": `a & b < c > d "e"` }], ":@": { rowspan: "2", colspan: "3" } }, { th: [] }],
                  },
                  { tr: [{ td: [{ "#text": "1 234,5" }] }] },
                ],
              },
            ],
            ":@": { "data-layout": "default" },
          },
        ],
      },
      { "ac:structured-macro": [{ "ac:plain-text-body": [{ "#text": "x <y> ]] & z" }] }], ":@": { "ac:name": "code" } },
      { p: [{ "#text": "é\u00a0😀" }] },
    ]);
  });
});
