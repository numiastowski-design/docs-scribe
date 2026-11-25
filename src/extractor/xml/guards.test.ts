/**
 * The guards on WORK and MEMORY. Every part is hostile; what is bounded is what is actually allocated or walked,
 * checked BEFORE the allocation it protects.
 * - Nesting depth: `MAX_XML_DEPTH`, ONE definition, imported from `ooxml/xml.ts`: an element with content is read down
 *   to level MAX_XML_DEPTH + 1, an empty-element tag one level further (the table below).
 * - Nodes of one part (`MAX_XML_NODES`) and attributes of one tag (`MAX_XML_ATTRIBUTES`): our own bounds.
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { MAX_XML_DEPTH } from "../ooxml/xml";
import { MAX_XML_ATTRIBUTES, tokenize } from "./tokenizer";
import { buildTree, MAX_XML_NODES, type XmlMode } from "./tree";
import type { XNode } from "./xml-fixtures.test.support";

const DEPTH = "xml: elements nested past the depth limit";
const NODES = "xml: more nodes than the limit";
const ATTRIBUTES = "xml: more attributes on one element than the limit";

const refusalOf = (run: () => unknown): ReaderError => {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return error as ReaderError;
  }
  throw new Error("expected a refusal");
};

/** The element nesting of a tree, an empty element counting as a level; walked without recursion. */
const depthOf = (nodes: readonly XNode[]): number => {
  let deepest = 0;
  const pending: [XNode, number][] = nodes.map((n) => [n, 1]);
  while (pending.length > 0) {
    const [node, depth] = pending.pop()!;
    const tag = Object.keys(node).find((k) => k !== ":@")!;
    if (tag === "#text") continue;
    deepest = Math.max(deepest, depth);
    for (const kid of node[tag] as XNode[]) pending.push([kid, depth + 1]);
  }
  return deepest;
};

/** `levels` elements nested, the innermost written as `leaf`. */
const nest = (levels: number, leaf: string): string => "<a>".repeat(levels - 1) + leaf + "</a>".repeat(levels - 1);

describe("xml guards — nesting depth, MAX_XML_DEPTH imported from ooxml/xml.ts", () => {
  // Ported from docx-reader.test.ts, « review 27/09 night, finding 7 »: document and body, two levels per content
  // control, then paragraph, run, text.
  const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
  const controls = (count: number): string =>
    `<w:document ${W_NS}><w:body>` +
    "<w:sdt><w:sdtContent>".repeat(count) +
    "<w:p><w:r><w:t>deep</w:t></w:r></w:p>" +
    "</w:sdtContent></w:sdt>".repeat(count) +
    "</w:body></w:document>";

  it("a document nested ~500 deep is read (245 controls: 495 levels)", () => {
    expect(depthOf(buildTree(controls(245), "document"))).toBe(495);
  });

  it("past the ceiling the part is refused with our literal, never a RangeError: 260 and 20 000 controls", () => {
    for (const count of [260, 20_000]) {
      const error = refusalOf(() => buildTree(controls(count), "document"));
      expect(error.code).toBe("EXTRACT_FAILED");
      expect(error.message).toBe(DEPTH);
    }
  });

  it.each<[string, number, string, "read" | "refused"]>([
    ["an element with content at level MAX_XML_DEPTH + 1", 1, "<a>x</a>", "read"],
    ["an empty pair at level MAX_XML_DEPTH + 1", 1, "<a></a>", "read"],
    ["an empty-element tag at level MAX_XML_DEPTH + 2", 2, "<a/>", "read"],
    ["an empty pair at level MAX_XML_DEPTH + 2", 2, "<a></a>", "refused"],
    ["an element with content at level MAX_XML_DEPTH + 2", 2, "<a>x</a>", "refused"],
    ["an empty-element tag at level MAX_XML_DEPTH + 3", 3, "<a/>", "refused"],
  ])("the boundary: %s", (_label, past, leaf, outcome) => {
    const source = nest(MAX_XML_DEPTH + past, leaf);
    if (outcome === "read") expect(depthOf(buildTree(source, "document"))).toBe(MAX_XML_DEPTH + past);
    else expect(refusalOf(() => buildTree(source, "document")).message).toBe(DEPTH);
  });

  it("holds in fragment mode too, with the same constant", () => {
    expect(depthOf(buildTree(nest(MAX_XML_DEPTH + 1, "<a/>"), "fragment"))).toBe(MAX_XML_DEPTH + 1);
    expect(refusalOf(() => buildTree(nest(MAX_XML_DEPTH + 2, "<a></a>"), "fragment")).message).toBe(DEPTH);
  });

  it("refuses as soon as the level is opened, not at the end of the text: a million start tags never closed", () => {
    expect(refusalOf(() => buildTree("<a>".repeat(1_000_000), "document")).message).toBe(DEPTH);
  });
});

describe("xml guards — what one part may allocate: MAX_XML_NODES, elements, attributes and text nodes together", () => {
  // Measured on 25/10/2026 (elements + attributes + text nodes under the root): largest versioned part 6 852 409
  // (an xlsx sheet of 50 M characters), largest corpus part 3 769 837.
  it("is a starting value above every real part measured", () => {
    expect(MAX_XML_NODES).toBeGreaterThan(6_852_409);
  });

  const SMALL = 1_000;
  const flat = (nodes: number): string => "<r>" + "<a/>".repeat(nodes - 1) + "</r>";

  it("at the bound the part is read; one node more is refused before it exists", () => {
    expect(buildTree(flat(SMALL), "document", "UTF-8", SMALL)[0]!.r).toHaveLength(SMALL - 1);
    expect(refusalOf(() => buildTree(flat(SMALL + 1), "document", "UTF-8", SMALL)).message).toBe(NODES);
  });

  it("attributes and text nodes count against the same bound; root-level white space and comments do not", () => {
    // Root + 333 × (element + attribute + text) = 1 000.
    const mixed = "<r>" + '<a b="1"/>x'.repeat(333) + "</r>";
    expect(buildTree(`\n<!--c-->${mixed}\n`, "document", "UTF-8", SMALL)[0]!.r).toHaveLength(666);
    expect(refusalOf(() => buildTree(mixed.replace("<r>", '<r c="1">'), "document", "UTF-8", SMALL)).message).toBe(
      NODES,
    );
    expect(refusalOf(() => buildTree(mixed.replace("<r>", "<r>x<?p?>"), "document", "UTF-8", SMALL)).message).toBe(
      NODES,
    );
    expect(
      refusalOf(() => buildTree(mixed.replace("<r>", "<r><![CDATA[]]>"), "document", "UTF-8", SMALL)).message,
    ).toBe(NODES);
  });

  it("a hostile part at the package's per-part ceiling, one empty element per 4 characters, is refused", () => {
    expect(refusalOf(() => buildTree(flat(16_000_000), "document")).message).toBe(NODES);
  }, 120_000);
});

describe("xml guards — the attributes of ONE tag, bounded while the tag is read", () => {
  const tag = (count: number): string => "<a" + Array.from({ length: count }, (_, k) => ` a${k}="1"`).join("") + "/>";

  it("is a starting value above every real tag measured (41 in the fixtures, 39 in the corpus, 25/10/2026)", () => {
    expect(MAX_XML_ATTRIBUTES).toBeGreaterThan(41);
  });

  it("MAX_XML_ATTRIBUTES are read; one more is refused before the element exists", () => {
    expect(Object.keys(buildTree(tag(MAX_XML_ATTRIBUTES), "document")[0]![":@"] as object)).toHaveLength(
      MAX_XML_ATTRIBUTES,
    );
    let opened = false;
    const error = refusalOf(() =>
      tokenize(tag(MAX_XML_ATTRIBUTES + 1), {
        open: () => {
          opened = true;
        },
        close() {},
        text() {},
      }),
    );
    expect(error.message).toBe(ATTRIBUTES);
    expect(opened).toBe(false);
  });
});

describe("xml guards — hostile parts at the package's per-part ceiling (64 MB) end in a refusal or a linear read, never a hang", () => {
  const MB64 = 64_000_000;
  const cases: [string, () => string, string, XmlMode][] = [
    ["64 MB of '<a '", () => "<a ".repeat(MB64 / 3), "xml: malformed name", "document"],
    ["a comment never closed", () => "<a><!--" + "x".repeat(MB64), "xml: unterminated comment", "document"],
    [
      "a CDATA section never closed",
      () => "<a><![CDATA[" + "x".repeat(MB64),
      "xml: unterminated CDATA section",
      "document",
    ],
    ["a PI never closed", () => "<a><?pi " + "x".repeat(MB64), "xml: unterminated processing instruction", "document"],
    [
      "an attribute value never closed",
      () => '<a b="' + "x".repeat(MB64),
      "xml: unterminated attribute value",
      "document",
    ],
    [
      "a character reference of 64 M digits",
      () => "<a>&#" + "9".repeat(MB64) + ";</a>",
      "xml: character reference out of range",
      "document",
    ],
    [
      "one tag repeating the same attribute 666 667 times (refused at the second)",
      () => "<a" + ' b="1"'.repeat(MB64 / 6 / 16) + "/>",
      "xml: repeated attribute",
      "document",
    ],
    ["16 M start tags never closed", () => "<a>".repeat(MB64 / 3), DEPTH, "fragment"],
  ];

  it.each(cases)(
    "%s",
    (_label, make, message, mode) => {
      expect(refusalOf(() => buildTree(make(), mode)).message).toBe(message);
    },
    60_000,
  );

  it("one tag of a million distinct attributes is refused by the attribute bound, before any element", () => {
    const source = "<a" + Array.from({ length: 1_000_000 }, (_, k) => ` a${k.toString(36)}="1"`).join("") + "/>";
    expect(refusalOf(() => buildTree(source, "document")).message).toBe(ATTRIBUTES);
  }, 60_000);

  it("64 MB of leading zeros in one character reference is read in one pass (§4.1 allows them)", () => {
    expect(buildTree("<a>&#" + "0".repeat(MB64) + "65;</a>", "document")).toEqual([{ a: [{ "#text": "A" }] }]);
  }, 60_000);
});
