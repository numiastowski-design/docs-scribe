/**
 * The interface module. Its two callers build `new XmlParser(options)` once and call `.parse(text)`; `xml-parser.ts`
 * answers that call shape with our own reader. The reader's own options are captured from its module, never copied;
 * the fragment caller lives outside this package, so its options are the ones the test support writes (`productParsers`).
 *
 * FROZEN TREES: one sha256 per versioned DOCUMENT, over the trees our reader builds from the bytes of all its XML parts
 * (`xml-trees.fixture.json`, section "documents"), so that any change of what is read names the document. The file is
 * rewritten only by a deliberate run (`XML_TREES_UPDATE=1 npx vitest run <this file>`, which then fails on
 * purpose; run again without the flag), when a versioned document was added, removed or replaced; read the diff of the
 * file before committing it. Parts of 1 M characters and more are not digested (the one of about 50 MB included), as in the
 * test that reads each part from its text and from its bytes.
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { MAX_XML_DEPTH } from "../ooxml/xml";
import { MAX_STORAGE_XHTML_DEPTH, XmlFragmentParser, XmlParser } from "./xml-parser";
import {
  expectPinned,
  fixtureXmlParts,
  JsonDigest,
  publicOfficeXmlParts,
  productParsers,
  type XNode,
} from "./xml-fixtures.test.support";

const refusalOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    expect((error as ReaderError).code).toBe("EXTRACT_FAILED");
    return (error as ReaderError).message;
  }
  throw new Error("expected a refusal");
};

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const PART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve"> a &amp; b &#233;</w:t></w:r><?mso x?><!--c--></w:p></w:body></w:document>`;

const nest = (levels: number, leaf: string): string => "<a>".repeat(levels) + leaf + "</a>".repeat(levels);

/** How many elements deep the first chain of elements goes (the tree is walked, never recursed). */
const depthOf = (nodes: readonly XNode[]): number => {
  let depth = 0;
  let level = nodes;
  for (;;) {
    const element = level.find((node) => !("#text" in node));
    if (element === undefined) return depth;
    depth++;
    level = element[Object.keys(element).find((key) => key !== ":@")!] as XNode[];
  }
};

describe("xml-parser — the callers' call shape, answered by our reader", () => {
  it("XmlParser, built with ooxml/xml.ts's own options, reads an Office part", async () => {
    const { docxReaderOptions } = await productParsers();
    const ours = new XmlParser(docxReaderOptions).parse(PART);
    expect(ours).toEqual([
      {
        "w:document": [
          {
            "w:body": [
              { "w:p": [{ "w:r": [{ "w:t": [{ "#text": " a & b é" }], ":@": { "xml:space": "preserve" } }] }] },
            ],
          },
        ],
        ":@": { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
      },
    ]);
  });

  it("XmlFragmentParser, built with a fragment caller's options, reads several roots", async () => {
    const { fragmentCallerOptions } = await productParsers();
    const page = `<ac:adf-extension-frame><table><tbody><tr><td colspan="2">1</td></tr></tbody></table></ac:adf-extension-frame> <p>x &quot;y&quot;</p>`;
    const ours = new XmlFragmentParser(fragmentCallerOptions).parse(page);
    expect(ours).toEqual([
      {
        "ac:adf-extension-frame": [
          { table: [{ tbody: [{ tr: [{ td: [{ "#text": "1" }], ":@": { colspan: "2" } }] }] }] },
        ],
      },
      { p: [{ "#text": 'x "y"' }] },
    ]);
    expect(ours).toHaveLength(2);
  });

  it("every versioned XML part under 1 M characters is read, from its text and from its bytes, to the same tree", async () => {
    const { docxReaderOptions } = await productParsers();
    const parts = await fixtureXmlParts((part) => part.text.length < 1_000_000);
    expect(parts.length).toBeGreaterThan(6000);
    const ours = new XmlParser(docxReaderOptions);
    for (const part of parts) {
      let fromText: unknown;
      let fromBytes: unknown;
      expect(() => (fromText = ours.parse(part.text)), part.id).not.toThrow();
      expect(() => (fromBytes = ours.parse(part.bytes!)), part.id).not.toThrow();
      expect(fromBytes, part.id).toEqual(fromText);
    }
  }, 300_000);

  it("every versioned document reads, part by part from its bytes, to the frozen trees (xml-trees.fixture.json): a change names the document", async () => {
    const { docxReaderOptions } = await productParsers();
    const reader = new XmlParser(docxReaderOptions);
    const digests = new Map<string, JsonDigest>();
    let parts = 0;
    // Every part under 1 M characters, as the test above: streamed, each tree is hashed and dropped. A part our reader refuses is pinned by its literal.
    const { documents } = await publicOfficeXmlParts((part) => {
      if (part.text.length >= 1_000_000) return;
      parts++;
      let reading: unknown;
      try {
        reading = reader.parse(part.bytes!);
      } catch (error) {
        if (!(error instanceof ReaderError)) throw error;
        reading = { refused: error.message };
      }
      const digest = digests.get(part.document) ?? new JsonDigest();
      digests.set(part.document, digest.add(part.id, reading));
    });
    expect(parts).toBeGreaterThan(6000);
    const emptyDocument = new JsonDigest().hex();
    expectPinned(
      "documents",
      Object.fromEntries(documents.map((document) => [document, digests.get(document)?.hex() ?? emptyDocument])),
    );
  }, 600_000);

  it("only the callers' own options exist: any other value or key is a caller's bug, a plain Error, never a document refusal", async () => {
    const { docxReaderOptions, fragmentCallerOptions } = await productParsers();
    for (const changed of [
      { preserveOrder: false },
      { trimValues: true },
      { htmlEntities: false },
      { attributeNamePrefix: "@_" },
      { parseAttributeValue: true },
      { ignoreDeclaration: true },
      { maxNestedTags: 0 },
      { processEntities: false },
    ]) {
      const options = { ...docxReaderOptions, ...changed } as never;
      expect(() => new XmlParser(options), JSON.stringify(changed)).toThrow(/^xml-parser: /);
      expect(() => new XmlParser(options)).not.toThrow(ReaderError);
      expect(
        () => new XmlFragmentParser({ ...fragmentCallerOptions, ...changed } as never),
        JSON.stringify(changed),
      ).toThrow(/^xml-parser: /);
    }
    // An Office part is always read under the product's own depth bound: it must be passed.
    expect(() => new XmlParser(fragmentCallerOptions as never)).toThrow(
      "xml-parser: maxNestedTags is required for an Office part",
    );
  });

  it("a refusal is our own ReaderError with a literal, never document content", async () => {
    const parser = new XmlParser((await productParsers()).docxReaderOptions);
    expect(refusalOf(() => parser.parse(`<w:document ${W}><w:body><w:p></w:body></w:document>`))).toBe(
      "xml: end tag does not match its start tag",
    );
    expect(refusalOf(() => parser.parse(`<!DOCTYPE a [<!ENTITY secret "x">]><a>&secret;</a>`))).toBe(
      "xml: document type or markup declaration",
    );
    expect(refusalOf(() => parser.parse(`<a>Mallory</a><b/>`))).toBe("xml: more than one root element");
  });

  it("a part with no element and no text reads as empty (middle ground, 25/10/2026)", async () => {
    const parser = new XmlParser((await productParsers()).docxReaderOptions);
    expect(parser.parse(`<?xml version="1.0"?>\r\n<!--x-->`)).toEqual([]);
    expect(parser.parse(new Uint8Array(0))).toEqual([]);
  });
});

describe("xml-parser — the bytes of a part: decoded by decode.ts", () => {
  const options = async (): Promise<never> => (await productParsers()).docxReaderOptions as never;
  const utf16le = (text: string): Uint8Array => {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = 0xff;
    out[1] = 0xfe;
    for (let i = 0; i < text.length; i++) {
      out[2 + i * 2] = text.charCodeAt(i) & 0xff;
      out[3 + i * 2] = text.charCodeAt(i) >> 8;
    }
    return out;
  };

  it("a UTF-16 part with its byte-order mark is read (XML 1.0 §4.3.3); the same text as UTF-8 bytes gives the same tree", async () => {
    const parser = new XmlParser(await options());
    const xml = `<?xml version="1.0" encoding="UTF-16"?><w:document ${W}><w:body><w:t>é€</w:t></w:body></w:document>`;
    expect(parser.parse(utf16le(xml))).toEqual(parser.parse(new TextEncoder().encode(xml.replace("UTF-16", "UTF-8"))));
  });

  it("a declared legacy encoding is read as declared from the bytes (option B); its text already decoded as UTF-8 is refused, never read as gibberish", async () => {
    const parser = new XmlParser(await options());
    const xml = `<?xml version="1.0" encoding="windows-1252"?><a>caf\u00e9 \u20ac</a>`;
    const bytes = Uint8Array.from([...xml].map((c) => (c === "\u20ac" ? 0x80 : c.charCodeAt(0))));
    expect(parser.parse(bytes)).toEqual([{ a: [{ "#text": "café €" }] }]);
    const productText = new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
    expect(refusalOf(() => parser.parse(productText))).toBe("xml: declared encoding does not match the bytes");
  });

  // Review 26/10/2026, item 5: `parse(string)` is not what the product calls (`parseXml` only ever hands bytes), but the
  // tests do; it follows the « prefer UTF-8 » rule of `decode.ts` instead of contradicting it.
  it("text already decoded, accented, under a Windows-1252 declaration is read as the UTF-8 it is; with U+FFFD in it, or under another encoding, still refused", async () => {
    const parser = new XmlParser(await options());
    const xml = `<?xml version="1.0" encoding="windows-1252"?><a>café €</a>`;
    expect(parser.parse(xml)).toEqual([{ a: [{ "#text": "café €" }] }]);
    expect(parser.parse(xml)).toEqual(parser.parse(new TextEncoder().encode(xml)));
    expect(refusalOf(() => parser.parse(xml.replace("café", "caf�")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => parser.parse(xml.replace("windows-1252", "shift_jis")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => parser.parse(xml.replace("windows-1252", "UTF-16")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
  });

  it("text whose declaration names an ASCII-compatible encoding, and holds only ASCII, is read: every such encoding decodes ASCII bytes alike", async () => {
    const parser = new XmlParser(await options());
    const xml = `<?xml version="1.0" encoding="Windows-1252"?><Registry><Item Id="a&amp;b"/></Registry>`;
    expect(parser.parse(xml)).toEqual([{ Registry: [{ Item: [], ":@": { Id: "a&b" } }] }]);
    // Named UTF-16 over ASCII text: a contradiction (XML 1.0 §4.3.3), refused as by decode.ts.
    expect(refusalOf(() => parser.parse(xml.replace("Windows-1252", "UTF-16")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
  });

  it("the stateful encoding keeps its meaning: ESC sequences of a declared ISO-2022-JP are read as the bytes they were, not refused as a forbidden character", async () => {
    // Measured on the whole-text path before the head-only declaration read (26/10/2026): the escape sequence is swallowed.
    const parser = new XmlParser(await options());
    expect(parser.parse(`<?xml version="1.0" encoding="iso-2022-jp"?><a>\u001b(B</a>`)).toEqual([{ a: [] }]);
    expect(parser.parse(`<?xml version="1.0" encoding="iso-2022-jp"?><a>x</a>`)).toEqual([{ a: [{ "#text": "x" }] }]);
  });

  it("a long ASCII text under a declared Windows-1252 is read whole after its declaration alone was decoded", async () => {
    const parser = new XmlParser(await options());
    const body = "x".repeat(50_000);
    expect(parser.parse(`<?xml version="1.0" encoding="windows-1252"?><a>${body}</a>`)).toEqual([
      { a: [{ "#text": body }] },
    ]);
  });

  it("bytes illegal in UTF-8 are refused, never replaced by U+FFFD", async () => {
    const parser = new XmlParser(await options());
    expect(refusalOf(() => parser.parse(Uint8Array.from([0x3c, 0x61, 0x3e, 0xff, 0x3c, 0x2f, 0x61, 0x3e])))).toBe(
      "xml: bytes not legal in their encoding",
    );
  });
});

describe("xml-parser — the depth each caller gets", () => {
  it("an Office part: MAX_XML_DEPTH, the option ooxml/xml.ts passes (513 levels read, 514 refused)", async () => {
    const options = (await productParsers()).docxReaderOptions;
    expect(options.maxNestedTags).toBe(MAX_XML_DEPTH);
    const parser = new XmlParser(options);
    expect(depthOf(parser.parse(nest(MAX_XML_DEPTH + 1, "x")))).toBe(MAX_XML_DEPTH + 1);
    expect(refusalOf(() => parser.parse(nest(MAX_XML_DEPTH + 2, "x")))).toBe(
      "xml: elements nested past the depth limit",
    );
  });

  it("storage XHTML: a caller that passes no maxNestedTags gets MAX_STORAGE_XHTML_DEPTH", async () => {
    const options = (await productParsers()).fragmentCallerOptions;
    expect(options.maxNestedTags).toBeUndefined();
    const parser = new XmlFragmentParser(options);
    for (const [levels, leaf] of [
      [MAX_STORAGE_XHTML_DEPTH + 1, "x"],
      [MAX_STORAGE_XHTML_DEPTH + 1, "<a/>"],
    ] as const) {
      expect(depthOf(parser.parse(nest(levels, leaf)))).toBe(levels + (leaf === "x" ? 0 : 1));
    }
    for (const [levels, leaf] of [
      [MAX_STORAGE_XHTML_DEPTH + 2, "x"],
      [MAX_STORAGE_XHTML_DEPTH + 2, "<a/>"],
    ] as const) {
      expect(refusalOf(() => parser.parse(nest(levels, leaf)))).toBe("xml: elements nested past the depth limit");
    }
  });
});
