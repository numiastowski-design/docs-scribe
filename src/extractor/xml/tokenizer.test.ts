import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { decodeXmlPart } from "./decode";
import { buildTree } from "./tree";
import { fixtureXmlParts } from "./xml-fixtures.test.support";
import {
  nonLiteralReasonCallsIn,
  raisesRefusal,
  READER_SOURCE_NAMES,
  readerSourceText,
  refusalReasonsIn,
  unrecognisedReferencesIn,
} from "./xml-refusals.test.support";
import { eventsOf } from "./xml-shadow.test.support";

const refusal = (source: string): ReaderError => {
  try {
    eventsOf(source);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return error as ReaderError;
  }
  throw new Error("expected a refusal");
};

describe("xml tokenizer — tags, attributes, text (XML 1.0 §2.3, §3.1)", () => {
  it("reports start, end and empty-element tags, attributes in order, text between them", () => {
    expect(eventsOf(`<a x="1" y='2'><b/>t</a >`)).toEqual([
      [
        "open",
        "a",
        [
          ["x", "1"],
          ["y", "2"],
        ],
      ],
      ["open", "b", []],
      ["close", "b"],
      ["text", "t"],
      ["close", "a"],
    ]);
  });

  it("accepts white space around = and inside tags (Eq, S)", () => {
    expect(eventsOf(`<a\n x = "1"\t/>`)).toEqual([
      ["open", "a", [["x", "1"]]],
      ["close", "a"],
    ]);
  });

  it("keeps white-space-only text, never reports empty text", () => {
    expect(eventsOf(`<a> <b></b></a>`)).toEqual([
      ["open", "a", []],
      ["text", " "],
      ["open", "b", []],
      ["close", "b"],
      ["close", "a"],
    ]);
  });

  it("names may hold a colon, digits after the first character, non-ASCII letters (Name production, §2.3)", () => {
    expect(eventsOf(`<w:t é.1-x="v"/>`)).toEqual([
      ["open", "w:t", [["é.1-x", "v"]]],
      ["close", "w:t"],
    ]);
  });

  // U+FEFF and U+1680 are name characters (§2.3 [4]), not white space: ours keeps them in the name; in character data too.
  it("names holding U+FEFF or U+1680 keep them: in an attribute name, as the first character of an element name", () => {
    expect(eventsOf(`<a b\uFEFFc="1"/>`)).toEqual([
      ["open", "a", [["b\uFEFFc", "1"]]],
      ["close", "a"],
    ]);
    expect(eventsOf(`<\uFEFFa/>`)).toEqual([
      ["open", "\uFEFFa", []],
      ["close", "\uFEFFa"],
    ]);
    expect(eventsOf(`<a \u1680b="1"/>`)).toEqual([
      ["open", "a", [["\u1680b", "1"]]],
      ["close", "a"],
    ]);
  });

  it("character data holding U+FEFF or U+1680 is kept untrimmed", () => {
    expect(eventsOf(`<a>x\uFEFFy\u1680</a>`)).toEqual([
      ["open", "a", []],
      ["text", "x\uFEFFy\u1680"],
      ["close", "a"],
    ]);
  });
});

describe("xml tokenizer — references (§4.1, §4.6) and line ends (§2.11)", () => {
  it("decodes the five predefined entities and character references in text and attribute values", () => {
    expect(eventsOf(`<a v="&amp;&lt;&gt;&quot;&apos;&#233;&#x1F600;">&amp;&lt;&#xe9;&#0000065;</a>`)).toEqual([
      ["open", "a", [["v", `&<>"'é😀`]]],
      ["text", "&<éA"],
      ["close", "a"],
    ]);
  });

  it("turns CR LF and a lone CR into LF, in text and in attribute values; a &#13; stays a CR", () => {
    expect(eventsOf(`<a v="1\r\n2\r3">x\r\ny\rz&#13;</a>`)).toEqual([
      ["open", "a", [["v", "1\n2\n3"]]],
      ["text", "x\ny\nz\r"],
      ["close", "a"],
    ]);
  });

  it("does not normalise attribute white space (an open limit against §3.3.3)", () => {
    expect(eventsOf(`<a v="\tx"/>`)).toEqual([
      ["open", "a", [["v", "\tx"]]],
      ["close", "a"],
    ]);
  });
});

describe("xml tokenizer — XML declaration and byte-order mark (§2.8, §4.3.3)", () => {
  it("reads the declaration in its eight versioned spellings, after an optional BOM, and reports nothing for it", () => {
    for (const decl of [
      `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?>`,
      `<?xml version="1.0" encoding="utf-8"?>`,
      `<?xml version="1.0" encoding="utf-8" standalone="yes"?>`,
      `<?xml version="1.0" standalone="no"?>`,
      `<?xml version="1.0" encoding="UTF-8"?>`,
      `<?xml version="1.0"?>`,
    ]) {
      expect(eventsOf(`\uFEFF${decl}\r\n<a/>`)).toEqual([
        ["open", "a", []],
        ["close", "a"],
      ]);
    }
  });

  it("refuses a declaration that names another encoding, or is malformed", () => {
    expect(refusal(`<?xml version="1.0" encoding="UTF-16"?><a/>`).message).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusal(`<?xml encoding="UTF-8"?><a/>`).message).toBe("xml: malformed XML declaration");
    expect(refusal(`<?xml version="1.0" standalone="maybe"?><a/>`).message).toBe("xml: malformed XML declaration");
    expect(refusal(`<?xml version="1.0"`).message).toBe("xml: malformed XML declaration");
  });

  it("checks each pseudo-attribute as it is read, before reading the next one (a bad encoding is refused even in an unterminated declaration)", () => {
    expect(refusal(`<?xml version="1.0" encoding="UTF-16" standalone="yes" `).message).toBe(
      "xml: declared encoding does not match the bytes",
    );
  });

  it("refuses a repeated pseudo-attribute at its second occurrence (that the rest is not collected is read in the code, not observed here)", () => {
    expect(refusal(`<?xml version="1.0"${' version="1.0"'.repeat(3)} encoding="UTF-16"?><a/>`).message).toBe(
      "xml: malformed XML declaration",
    );
  });
});

describe("xml tokenizer — what is not well-formed is refused with our own literal, never document content", () => {
  const cases: [string, string][] = [
    [`<a>&nbsp;</a>`, "xml: entity reference other than the five predefined ones"],
    [`<a>&secret;</a>`, "xml: entity reference other than the five predefined ones"],
    [`<a>a & b</a>`, "xml: entity reference other than the five predefined ones"],
    [`<a>&#0;</a>`, "xml: character reference to a forbidden character"],
    [`<a>&#xD800;</a>`, "xml: character reference to a forbidden character"],
    [`<a>&#x110000;</a>`, "xml: character reference out of range"],
    [`<a>&#1114112;</a>`, "xml: character reference out of range"],
    [`<a>&#${"9".repeat(100_000)};</a>`, "xml: character reference out of range"],
    [`<a>&#X41;</a>`, "xml: malformed character reference"],
    [`<a>&#x41</a>`, "xml: malformed character reference"],
    [`<a>&#;</a>`, "xml: malformed character reference"],
    [`<a>]]></a>`, "xml: ']]>' in character data"],
    [`<a>\u0001</a>`, "xml: forbidden character"],
    [`<a>\uD800x</a>`, "xml: forbidden character"],
    [`<a v="\uFFFE"/>`, "xml: forbidden character"],
    [`<a v="<"/>`, "xml: '<' in an attribute value"],
    [`<a v=1/>`, "xml: attribute value without quotes"],
    [`<a v/>`, "xml: attribute without a value"],
    [`<a v="1"w="2"/>`, "xml: attributes not separated by white space"],
    [`<a v="1" v="2"/>`, "xml: repeated attribute"],
    [`<a v="1`, "xml: unterminated attribute value"],
    [`<a v="1"`, "xml: unterminated tag"],
    [`<1a/>`, "xml: malformed name"],
    [`<a 1v="x"/>`, "xml: malformed name"],
    [`< a/>`, "xml: malformed name"],
    [`<a></ a>`, "xml: malformed name"],
    [`<a></a x>`, "xml: malformed end tag"],
    [`<__proto__/>`, "xml: name reserved by JavaScript objects"],
    [`<a constructor="x"/>`, "xml: name reserved by JavaScript objects"],
    [`<a toString="x"/>`, "xml: name reserved by JavaScript objects"],
    [`<toString/>`, "xml: name reserved by JavaScript objects"],
    [`<a valueOf="1"/>`, "xml: name reserved by JavaScript objects"],
    [`<a hasOwnProperty="1"/>`, "xml: name reserved by JavaScript objects"],
    [`<?xml version="1.${"0".repeat(70)}"?><a/>`, "xml: XML declaration value too long"],
  ];

  const labelled = cases.map(
    ([source, message]) => [source.length > 40 ? `${source.slice(0, 40)}…` : source, source, message] as const,
  );
  it.each(labelled)("%j is refused", (_label, source, message) => {
    const error = refusal(source);
    expect(error.code).toBe("EXTRACT_FAILED");
    expect(error.message).toBe(message);
  });
});

describe("xml tokenizer — a no-break space is character data like any other", () => {
  it("keeps U+00A0 whether it is written as a reference or as the character itself: not trimmed, not dropped", () => {
    expect(eventsOf(`<a>&#160;</a>`)).toEqual([
      ["open", "a", []],
      ["text", "\u00a0"],
      ["close", "a"],
    ]);
    expect(eventsOf(`<a>x&#xA0;y</a>`)).toEqual([
      ["open", "a", []],
      ["text", "x\u00a0y"],
      ["close", "a"],
    ]);
    expect(eventsOf("<a>\u00a0</a>")).toEqual([
      ["open", "a", []],
      ["text", "\u00a0"],
      ["close", "a"],
    ]);
  });
});

describe("xml reader — every XML part of the versioned fixtures is read", () => {
  it("every part under 1 M characters is read without a refusal, from its bytes decoded as the product decodes them", async () => {
    const parts = await fixtureXmlParts((part) => part.text.length < 1_000_000);
    // Floor re-measured on the public documents (25/11/2026): 3 100 parts.
    expect(parts.length).toBeGreaterThan(3_099);
    for (const part of parts) {
      const decoded = decodeXmlPart(part.bytes!);
      expect(() => buildTree(decoded.text, "document", decoded.encoding), part.id).not.toThrow();
    }
  }, 300_000);
});

describe("raisesRefusal and refusalReasonsIn (the check that the closed lists of the fuzzes rely on)", () => {
  it("is true for a reason written as xmlRefusal(...) and for one handed to readUntil(..., ...)", () => {
    expect(raisesRefusal("xml: unterminated comment")).toBe(true);
    expect(raisesRefusal("xml: unterminated CDATA section")).toBe(true);
  });

  it("is false for a made-up reason, for the prefix alone, and for text that does not start with it", () => {
    expect(raisesRefusal("xml: a reason nobody ever wrote")).toBe(false);
    expect(raisesRefusal("xml: ")).toBe(false);
    expect(raisesRefusal("unterminated comment")).toBe(false);
  });

  it("returns the reasons of the calls that a small source writes, and no other literal", () => {
    const source = [
      '// throw xmlRefusal("in a line comment");',
      '/* throw xmlRefusal("in a block comment"); this.readUntil("]]>", "in a block comment too"); */',
      "const text = 'throw xmlRefusal(\"inside a string\")';",
      'log(1, "last argument of another call");',
      'const quoted = /"[^"]*/; throw xmlRefusal("after a regex literal holding a quote");',
      'this.readUntil("?>",\n  "wrapped over two lines");',
      "throw xmlRefusal('single quoted');",
      "throw xmlRefusal(`template without substitution`);",
      String.raw`throw xmlRefusal("tab\there");`,
    ].join("\n");
    expect([...refusalReasonsIn(source)].sort()).toEqual(
      [
        "after a regex literal holding a quote",
        "single quoted",
        "tab\there",
        "template without substitution",
        "wrapped over two lines",
      ].sort(),
    );
    expect(nonLiteralReasonCallsIn(source)).toEqual([]);
  });

  it("reports a call whose reason is not a literal instead of ignoring it", () => {
    const source =
      'throw xmlRefusal(reason); throw xmlRefusal("a " + b); this.readUntil("]", `x${y}`); throw xmlRefusal("fine");';
    expect(nonLiteralReasonCallsIn(source)).toEqual([
      "xmlRefusal(reason)",
      'xmlRefusal("a " + b)',
      'this.readUntil("]", `x${y}`)',
    ]);
    expect([...refusalReasonsIn(source)]).toEqual(["fine"]);
  });

  it("drops a reason whose raise is replaced by a comment quoting it, in the real tree.ts, and finds it in the untouched text", () => {
    const raise = 'throw xmlRefusal("unclosed element");';
    const untouched = readerSourceText("tree.ts");
    expect(untouched.split(raise)).toHaveLength(2);
    expect(refusalReasonsIn(untouched, "tree.ts")).toContain("unclosed element");
    const commented = untouched.replace(raise, `// was: ${raise}`);
    expect(commented).not.toBe(untouched);
    expect(refusalReasonsIn(commented, "tree.ts")).not.toContain("unclosed element");
    expect(raisesRefusal("xml: unclosed element")).toBe(true);
  });

  it("finds every literal that the real sources write on a line of code", () => {
    for (const name of READER_SOURCE_NAMES) {
      const raw = readerSourceText(name);
      const onCodeLines = raw.split("\n").filter((line) => !/^\s*(\/\/|\*|\/\*)/.test(line));
      const written = onCodeLines.flatMap((line) =>
        [...line.matchAll(/xmlRefusal\("([^"]*)"\)/g)].map((match) => match[1]!),
      );
      expect(written.length, `${name}: no xmlRefusal("…") on a code line`).toBeGreaterThan(0);
      for (const reason of written) expect(refusalReasonsIn(raw, name), `${name}: ${reason}`).toContain(reason);
    }
  });

  it("pins the one non-literal reason in the real sources: readUntil's own forward of its parameter", () => {
    const sites = READER_SOURCE_NAMES.flatMap((name) =>
      nonLiteralReasonCallsIn(readerSourceText(name), name).map((call) => `${name}: ${call}`),
    );
    expect(sites).toEqual(["tokenizer.ts: xmlRefusal(unterminated)"]);
  });

  it("reports each way of naming xmlRefusal or readUntil that the walk of calls cannot follow", () => {
    const forms = [
      '(xmlRefusal)("a");',
      '(xmlRefusal as F)("a");',
      'obj["xmlRefusal"]("a");',
      'xmlRefusal.call(null, "a");',
      'import { xmlRefusal as refuse } from "./tokenizer";',
      "const r = xmlRefusal;",
      'this.readUntil.call(this, "]", "a");',
    ];
    for (const form of forms) expect(unrecognisedReferencesIn(form), form).not.toEqual([]);
    const recognised =
      'import { xmlRefusal } from "./t"; export { xmlRefusal }; const xmlRefusal = (r: string) => r; throw xmlRefusal("a"); this.readUntil("]", "b");';
    expect(unrecognisedReferencesIn(recognised)).toEqual([]);
  });

  it("finds no reference the walk cannot follow in the real sources", () => {
    for (const name of READER_SOURCE_NAMES)
      expect(unrecognisedReferencesIn(readerSourceText(name), name), name).toEqual([]);
  });
});
