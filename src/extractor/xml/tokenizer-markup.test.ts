/**
 * Comments (XML 1.0 §2.5), processing instructions (§2.6), CDATA sections (§2.7), and every other `<!` refused (§2.8
 * doctypedecl and markup declarations; no DTD is ever read). What a valid Office part holds is read; what no valid part
 * can hold is refused with our own literal.
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { tokenize } from "./tokenizer";
import { eventsOf } from "./xml-shadow.test.support";

const refusalOf = (source: string): ReaderError => {
  try {
    eventsOf(source);
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return error as ReaderError;
  }
  throw new Error("expected a refusal");
};

describe("xml tokenizer — comments (XML 1.0 §2.5)", () => {
  it("drops a comment and reads the text on both sides as ONE run", () => {
    expect(eventsOf(`<a>x<!--c-->y<!---->z</a>`)).toEqual([
      ["open", "a", []],
      ["text", "xyz"],
      ["close", "a"],
    ]);
    expect(eventsOf(`<a>&amp;<!-- <b> & -->&lt;</a>`)).toEqual([
      ["open", "a", []],
      ["text", "&<"],
      ["close", "a"],
    ]);
  });

  it("reads each of these comment cases as exactly these events", () => {
    const open = ["open", "a", []] as const;
    const close = ["close", "a"] as const;
    expect(eventsOf(`<a>&amp;<!--c-->lt;</a>`)).toEqual([open, ["text", "&lt;"], close]);
    expect(eventsOf(`<a>\r<!--c-->\n</a>`)).toEqual([open, ["text", "\n\n"], close]);
    expect(eventsOf(`<a>]]<!--c-->></a>`)).toEqual([open, ["text", "]]>"], close]);
    expect(eventsOf(`<a> <!--c--> </a>`)).toEqual([open, ["text", "  "], close]);
    expect(eventsOf(`<a><!--c--></a>`)).toEqual([open, close]);
    expect(eventsOf(`<a>x<!--c--><b/>y</a>`)).toEqual([
      open,
      ["text", "x"],
      ["open", "b", []],
      ["close", "b"],
      ["text", "y"],
      close,
    ]);
    expect(eventsOf(`<a><b/><!--c--><b/></a>`)).toEqual([
      open,
      ["open", "b", []],
      ["close", "b"],
      ["open", "b", []],
      ["close", "b"],
      close,
    ]);
    expect(eventsOf(`<a>&#160;<!--x-->&#xA0;</a>`)).toEqual([open, ["text", "\u00a0\u00a0"], close]);
  });

  it("accepts comments before and after the root element (Misc), and after the declaration", () => {
    expect(eventsOf(`<?xml version="1.0"?><!--a--><a/><!--b-->`)).toEqual([
      ["open", "a", []],
      ["close", "a"],
    ]);
    expect(eventsOf(`<?xml version="1.0" encoding="UTF-8"?>\r\n<!--a--><a/><!--b-->`)).toEqual([
      ["open", "a", []],
      ["close", "a"],
    ]);
  });

  it("refuses '--' inside a comment, a comment ending in '--->', an unterminated comment, a forbidden character", () => {
    expect(refusalOf(`<a><!--a--b--></a>`).message).toBe("xml: '--' inside a comment");
    expect(refusalOf(`<a><!--a---></a>`).message).toBe("xml: '--' inside a comment");
    expect(refusalOf(`<a><!--x`).message).toBe("xml: unterminated comment");
    expect(refusalOf(`<a><!--\u0001--></a>`).message).toBe("xml: forbidden character");
  });

  it("refuses a reference split by a comment (`&<!---->amp;`: a bare `&` in character data, §2.4)", () => {
    expect(refusalOf(`<a>&<!--c-->amp;</a>`).message).toBe("xml: entity reference other than the five predefined ones");
  });
});

describe("xml tokenizer — processing instructions (§2.6)", () => {
  it("reads and drops a processing instruction; it splits the text around it", () => {
    expect(eventsOf(`<?mso-contentType?><a>x<?pi some data?>y<?xml-stylesheet href="s"?></a>`)).toEqual([
      ["open", "a", []],
      ["text", "x"],
      ["text", "y"],
      ["close", "a"],
    ]);
  });

  it("reads each of these processing-instruction cases as exactly these events", () => {
    const open = ["open", "a", []] as const;
    const close = ["close", "a"] as const;
    expect(eventsOf(`<?xml version="1.0"?><?mso-contentType?><a/>`)).toEqual([open, close]);
    expect(eventsOf(`<a><?p2?></a>`)).toEqual([open, close]);
    expect(eventsOf(`<a><?pi a\r\nb?></a>`)).toEqual([open, close]);
    expect(eventsOf(`<a>x<!--c--><?p?>y</a>`)).toEqual([open, ["text", "x"], ["text", "y"], close]);
  });

  it("reports the target and the data to a visitor that asks for them, line ends normalised", () => {
    const seen: [string, string][] = [];
    tokenize(`<a><?pi  one ? two\r\n?><?p2?></a>`, {
      open() {},
      close() {},
      text() {},
      processingInstruction: (target, data) => seen.push([target, data]),
    });
    expect(seen).toEqual([
      ["pi", "one ? two\n"],
      ["p2", ""],
    ]);
  });

  // The 06/10 version refused a target named like a JavaScript object property. A PI target is a valid Name that never
  // becomes a key of ours.
  it("reads a target named like a JavaScript object property", () => {
    expect(eventsOf(`<a><?toString?></a>`)).toEqual([
      ["open", "a", []],
      ["close", "a"],
    ]);
  });

  it("refuses the reserved target xml anywhere but the declaration, a target glued to its data, an unterminated one", () => {
    expect(refusalOf(`<a><?xml version="1.0"?></a>`).message).toBe("xml: reserved processing instruction target");
    expect(refusalOf(`<a/><?XmL x?>`).message).toBe("xml: reserved processing instruction target");
    expect(refusalOf(` <?xml version="1.0"?><a/>`).message).toBe("xml: reserved processing instruction target");
    expect(refusalOf(`<!--c--><?xml version="1.0"?><a/>`).message).toBe("xml: reserved processing instruction target");
    expect(refusalOf(`<a><?pi"x"?></a>`).message).toBe("xml: malformed processing instruction");
    expect(refusalOf(`<a><?pi x`).message).toBe("xml: unterminated processing instruction");
    expect(refusalOf(`<a><?pi`).message).toBe("xml: unterminated processing instruction");
  });
});

describe("xml tokenizer — CDATA sections (§2.7)", () => {
  it("reports a CDATA section raw, as a text node of its own even when empty, line ends normalised", () => {
    expect(eventsOf(`<a>x<![CDATA[<b>&amp;]]\r\n]]><![CDATA[]]>y</a>`)).toEqual([
      ["open", "a", []],
      ["text", "x"],
      ["text", "<b>&amp;]]\n"],
      ["text", ""],
      ["text", "y"],
      ["close", "a"],
    ]);
  });

  it("reads each of these CDATA cases as exactly these events", () => {
    const open = ["open", "a", []] as const;
    const close = ["close", "a"] as const;
    expect(eventsOf(`<a><![CDATA[]]></a>`)).toEqual([open, ["text", ""], close]);
    expect(eventsOf(`<a>x<![CDATA[y]]>z</a>`)).toEqual([open, ["text", "x"], ["text", "y"], ["text", "z"], close]);
    expect(eventsOf(`<a>x<![CDATA[&#65;&amp;]]></a>`)).toEqual([open, ["text", "x"], ["text", "&#65;&amp;"], close]);
    expect(eventsOf(`<a><![CDATA[a\rb]]></a>`)).toEqual([open, ["text", "a\nb"], close]);
  });

  it("without a cdata method, a non-empty section reaches text()", () => {
    const seen: string[] = [];
    tokenize(`<a><![CDATA[<x>]]><![CDATA[]]></a>`, { open() {}, close() {}, text: (v) => seen.push(v) });
    expect(seen).toEqual(["<x>"]);
  });

  it("refuses an unterminated section and a forbidden character inside one", () => {
    expect(refusalOf(`<a><![CDATA[x`).message).toBe("xml: unterminated CDATA section");
    expect(refusalOf(`<a><![CDATA[\uFFFF]]></a>`).message).toBe("xml: forbidden character");
  });
});

describe("xml tokenizer — no DTD, no markup declaration is ever read (§2.8; OPC forbids DTDs)", () => {
  it.each([
    `<!DOCTYPE a><a/>`,
    `<?xml version="1.0"?><!DOCTYPE w:document [<!ENTITY who "Mallory">]><w:document>&who;</w:document>`,
    `<a><!ENTITY x "y"></a>`,
    `<a><![INCLUDE[x]]></a>`,
    `<a><!x></a>`,
    `<a><!-x--></a>`,
  ])("%j is refused before anything is declared", (source) => {
    const error = refusalOf(source);
    expect(error.code).toBe("EXTRACT_FAILED");
    expect(error.message).toBe("xml: document type or markup declaration");
  });

  // A DOCTYPE inside a comment is valid XML: ours reads the comment. No versioned part and no corpus part holds
  // `<!DOCTYPE` at all.
  it("a DOCTYPE inside a comment is read: the comment is dropped", () => {
    expect(eventsOf(`<a><!-- <!DOCTYPE x> --></a>`)).toEqual([
      ["open", "a", []],
      ["close", "a"],
    ]);
  });

  // PI data is any Char up to the first `?>` (§2.6); quotes inside it mean nothing.
  it("a quote in a processing instruction's data follows §2.6", () => {
    expect(eventsOf(`<a><?pi x="?>"?>y</a>`)).toEqual([
      ["open", "a", []],
      ["text", '"?>y'],
      ["close", "a"],
    ]);
    expect(eventsOf(`<a><?pi it's?>y</a>`)).toEqual([
      ["open", "a", []],
      ["text", "y"],
      ["close", "a"],
    ]);
  });
});
