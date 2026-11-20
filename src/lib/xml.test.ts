import { describe, expect, it } from "vitest";

import {
  cdata,
  decodeEntities,
  element,
  parseXmlFragment,
  serializeXml,
  serializeXmlFragment,
  stripInvalidXmlChars,
  text,
} from "./xml";

describe("serializeXml", () => {
  it("serializes an empty element as self-closing", () => {
    expect(serializeXml(element("ri:attachment", { "ri:filename": "a.png" }))).toBe(
      '<ri:attachment ri:filename="a.png"/>',
    );
  });

  it("serializes nested elements with text", () => {
    const tree = element("p", {}, [text("hello "), element("strong", {}, [text("world")])]);
    expect(serializeXml(tree)).toBe("<p>hello <strong>world</strong></p>");
  });

  it("escapes &, < and > in text", () => {
    expect(serializeXml(text("a & b < c > d"))).toBe("a &amp; b &lt; c &gt; d");
  });

  it('escapes & and " in attribute values', () => {
    expect(serializeXml(element("a", { href: 'x&y"z' }))).toBe('<a href="x&amp;y&quot;z"/>');
  });

  it("wraps a cdata text node in <![CDATA[...]]> unescaped", () => {
    expect(serializeXml(cdata("<not escaped> & stuff"))).toBe("<![CDATA[<not escaped> & stuff]]>");
  });

  it("serializes multiple attributes in construction order", () => {
    expect(serializeXml(element("ac:parameter", { "ac:name": "title", extra: "x" }))).toBe(
      '<ac:parameter ac:name="title" extra="x"/>',
    );
  });
});

describe("parseXmlFragment", () => {
  it("parses a self-closing element with attributes", () => {
    expect(parseXmlFragment('<ri:attachment ri:filename="a.png"/>')).toEqual([
      { type: "element", tag: "ri:attachment", attrs: [["ri:filename", "a.png"]], children: [] },
    ]);
  });

  it("parses nested elements and text", () => {
    const [node] = parseXmlFragment("<p>hello <strong>world</strong></p>");
    expect(node).toEqual(element("p", {}, [text("hello "), element("strong", {}, [text("world")])]));
  });

  it("decodes the five predefined XML entities and numeric references", () => {
    const [node] = parseXmlFragment("<p>&amp;&lt;&gt;&quot;&apos;&#65;&#x42;</p>");
    expect(node).toEqual(element("p", {}, [text("&<>\"'AB")]));
  });

  it("decodes HTML4 named entities (Confluence's page GET API uses these for accented text)", () => {
    const [node] = parseXmlFragment("<p>&eacute;&egrave;&agrave;&euro;&mdash;&Eacute;</p>");
    expect(node).toEqual(element("p", {}, [text("éèà€—É")]));
  });

  it("leaves an unrecognized named entity untouched rather than guessing", () => {
    const [node] = parseXmlFragment("<p>&notareal;</p>");
    expect(node).toEqual(element("p", {}, [text("&notareal;")]));
  });

  it("parses a CDATA section verbatim, without entity decoding", () => {
    const [node] = parseXmlFragment("<ac:plain-text-link-body><![CDATA[a & <b>]]></ac:plain-text-link-body>");
    expect(node).toEqual(element("ac:plain-text-link-body", {}, [cdata("a & <b>")]));
  });

  it("drops comments (text either side is not merged into one node, but reserializes identically)", () => {
    const parsed = parseXmlFragment("<p>a<!-- comment -->b</p>");
    expect(parsed).toEqual([element("p", {}, [text("a"), text("b")])]);
    expect(serializeXmlFragment(parsed)).toBe("<p>ab</p>");
  });

  it("parses sibling top-level nodes without a single root", () => {
    expect(parseXmlFragment("<p>a</p><p>b</p>")).toEqual([
      element("p", {}, [text("a")]),
      element("p", {}, [text("b")]),
    ]);
  });

  it("throws on a mismatched closing tag", () => {
    expect(() => parseXmlFragment("<p><strong>x</p></strong>")).toThrow(/mismatched closing tag/);
  });

  it("throws on an unterminated element", () => {
    expect(() => parseXmlFragment("<p>x")).toThrow(/unterminated element/);
  });

  it("round-trips serialize -> parse -> serialize for a representative storage-format snippet", () => {
    const original = serializeXmlFragment([
      element("h1", {}, [text("Title")]),
      element("ac:image", { "ac:alt": 'a "quoted" alt' }, [element("ri:attachment", { "ri:filename": "x.png" })]),
    ]);
    const reparsed = serializeXmlFragment(parseXmlFragment(original));
    expect(reparsed).toBe(original);
  });
});

/**
 * Added on 08/09/2026 after a MEASUREMENT, not as a precaution. The fixture
 * `13-Police-Sans-CMap.pdf` — a PDF whose font is an embedded subset
 * without `/ToUnicode` — made pdf.js (before its removal) return raw codes as
 * control characters, and **seventeen of them reached the published
 * XHTML**: that is, the body of a request to the Confluence API, which
 * expects valid XML. The storage format is not tolerant HTML.
 *
 * No numeric reference redeems these characters — `&#1;` is just as
 * invalid as the bare byte (XML 1.0 §2.2). The only conforming treatment is
 * removal, and it belongs to the serializer: it is the one in charge of
 * never producing an invalid document, whatever the origin of the string.
 */
describe("caractères interdits en XML 1.0 (§2.2)", () => {
  const CTRL_SOH = "\u0001";
  const CTRL_BS = "\u0008";
  const CTRL_BEL = "\u0007";

  it("les retire d'un nœud texte", () => {
    expect(serializeXml(text(`a${CTRL_SOH}b${CTRL_BS}c`))).toBe("abc");
  });

  it("les retire AUSSI d'une section CDATA, que l'échappement court-circuite", () => {
    expect(serializeXml(cdata(`a${CTRL_SOH}b`))).toBe("<![CDATA[ab]]>");
  });

  it("les retire d'une valeur d'attribut", () => {
    expect(serializeXml(element("p", { title: `a${CTRL_BEL}b` }))).toBe('<p title="ab"/>');
  });

  it("préserve tabulation, saut de ligne et retour chariot, qui sont VALIDES", () => {
    const valides = "a\tb\nc\rd";
    expect(stripInvalidXmlChars(valides)).toBe(valides);
  });

  it("ne touche à rien d'imprimable — accents, tiret cadratin, euro", () => {
    const propre = "Chiffre d'affaires — 7 815 k€";
    expect(stripInvalidXmlChars(propre)).toBe(propre);
  });

  it("retire les non-caractères U+FFFE et U+FFFF", () => {
    expect(stripInvalidXmlChars("a\ufffeb\uffffc")).toBe("abc");
  });

  it("n'altère pas l'échappement des trois caractères XML", () => {
    expect(serializeXml(text("a & b < c > d"))).toBe("a &amp; b &lt; c &gt; d");
  });

  /**
   * Found while fixing the control characters: the fuzz corpus of
   * `storage.property.test.ts` ALREADY contained isolated surrogates
   * (`String.fromCharCode(0xd800)`) and nothing had ever flagged it — our
   * own parser tolerates them, where Confluence's XML parser would
   * reject them. The range D800-DFFF is outside the `Char` production: only
   * the PAIR is valid.
   */
  describe("surrogates isolés", () => {
    const HIGH = "\ud800";
    const LOW = "\udfff";

    it("retire un haut sans bas", () => {
      expect(stripInvalidXmlChars(`a${HIGH}b`)).toBe("ab");
    });

    it("retire un bas sans haut", () => {
      expect(stripInvalidXmlChars(`a${LOW}b`)).toBe("ab");
    });

    it("retire deux hauts consécutifs, qui ne forment pas une paire", () => {
      expect(stripInvalidXmlChars(`a${HIGH}${HIGH}b`)).toBe("ab");
    });

    it("PRÉSERVE une paire valide — un emoji astral traverse intact", () => {
      const emoji = "\u{1f600}"; // U+1F600, i.e. D83D DE00
      expect(stripInvalidXmlChars(`a${emoji}b`)).toBe(`a${emoji}b`);
      expect(serializeXml(text(emoji))).toBe(emoji);
    });

    it("préserve la paire ET retire l'isolé quand les deux coexistent", () => {
      expect(stripInvalidXmlChars(`\u{1f600}${HIGH}`)).toBe("\u{1f600}");
    });
  });
});

/**
 * THE SAFETY NET, distinct from the removal. It claims nothing about the data: it
 * claims that the serializer cannot produce an invalid document. If
 * it ever fires, a serialization path went around
 * `stripInvalidXmlChars` — a defect in our code, not dirty data.
 */
const CTRL_SOH = "\u0001";
const HIGH_SURROGATE = "\ud800";

describe("serializeXmlFragment — filet de bonne formation", () => {
  it("ne lève pas sur un fragment ordinaire", () => {
    expect(() => serializeXmlFragment([element("p", {}, [text("bonjour")])])).not.toThrow();
  });

  it("ne lève pas sur un texte contenant des caractères interdits — le retrait passe avant", () => {
    // This is the intended behaviour: the fix cleans, the net is only for
    // the paths that would have escaped the fix.
    expect(serializeXmlFragment([element("p", {}, [text(`a${CTRL_SOH}b`)])])).toBe("<p>ab</p>");
  });

  it("lève, en nommant le point de code et la position, si un caractère interdit atteint la sortie", () => {
    // We fabricate the case that the removal cannot produce: an `element`
    // node whose NAME carries a forbidden character. Tag names
    // are deliberately not cleaned (they come from code, never from
    // data) — which makes them the only lever to put the net to the test.
    expect(() => serializeXmlFragment([element(`p${CTRL_SOH}x`)])).toThrow(
      /invalid XML 1\.0 character U\+0001 at offset 2/,
    );
  });

  it("le message ne contient aucun contenu client — seulement le code et la position", () => {
    try {
      serializeXmlFragment([element(`p${CTRL_SOH}x`, {}, [text("SECRET-CLIENT")])]);
      expect.unreachable("attendu une levée");
    } catch (error) {
      expect((error as Error).message).not.toContain("SECRET-CLIENT");
      expect((error as Error).message).toContain("U+0001");
    }
  });

  it("détecte aussi un surrogate isolé, pas seulement les C0", () => {
    expect(() => serializeXmlFragment([element(`p${HIGH_SURROGATE}x`)])).toThrow(/U\+D800/);
  });
});

/**
 * TOLERANCES OF THE HOME-MADE PARSER — locked in the state they are in, not
 * fixed (method laid down on 08/09/2026).
 *
 * These tests document holes. They fail if someone one day believes they have
 * plugged them — like the number-reformatting test in
 * `pipeline/translation-invariants.test.ts`. Each tolerance carries its production path,
 * or the absence of a path, which decides whether it must be dealt with.
 */
describe("tolérances de parseXmlFragment par rapport à XML 1.0 §2.2", () => {
  const CTRL_SOH = "\u0001";

  it("DÉCODE une référence numérique interdite, fabriquant le caractère à partir d'une entrée légale", () => {
    // No production path parses what Confluence returns any more
    // (`stripBanner` disappeared on 30/09/2026): `canonicalizeStorage` only reads
    // our own output. The tolerance is recorded, not exploited.
    const [node] = parseXmlFragment("<p>a&#1;b</p>");
    const value = node?.type === "element" && node.children[0]?.type === "text" ? node.children[0].value : "";
    expect(value).toBe(`a${CTRL_SOH}b`);
  });

  it("décode aussi une référence à un surrogate isolé", () => {
    const [node] = parseXmlFragment("<p>a&#xD800;b</p>");
    const value = node?.type === "element" && node.children[0]?.type === "text" ? node.children[0].value : "";
    expect(value).toBe("a\ud800b");
  });

  it("accepte un caractère interdit BRUT dans le texte", () => {
    expect(() => parseXmlFragment(`<p>a${CTRL_SOH}b</p>`)).not.toThrow();
  });

  it("accepte un caractère interdit dans une valeur d'attribut et dans un CDATA", () => {
    expect(() => parseXmlFragment(`<p t="a${CTRL_SOH}b"/>`)).not.toThrow();
    expect(() => parseXmlFragment(`<p><![CDATA[a${CTRL_SOH}b]]></p>`)).not.toThrow();
  });

  it("accepte une esperluette nue, un attribut sans guillemets et un attribut dupliqué", () => {
    // No output path: the serializer emits a tree, it cannot
    // produce these forms. Recorded, not fixed — hardening the parser for
    // an unreachable case adds code and risk for zero protection.
    expect(() => parseXmlFragment("<p>a & b</p>")).not.toThrow();
    expect(() => parseXmlFragment("<p t=x/>")).not.toThrow();
    expect(() => parseXmlFragment('<p t="1" t="2"/>')).not.toThrow();
  });

  it("REJETTE bien les fautes structurelles — c'est sa seule garantie", () => {
    expect(() => parseXmlFragment("<p>x")).toThrow(/unterminated element/);
    expect(() => parseXmlFragment("<p><b>x</p></b>")).toThrow(/mismatched closing tag/);
  });
});

// Review HIGH of 07/10/2026: the entity table is a plain object, so « &constructor; » found an inherited member and
// String.fromCodePoint threw — the page check abandoned. An unknown name is kept as written.
describe("decodeEntities — an entity named like an object member", () => {
  it.each(["constructor", "toString", "hasOwnProperty", "__proto__"])("🔴 « &%s; » is kept as written", (name) => {
    expect(decodeEntities(`a &${name}; b`)).toBe(`a &${name}; b`);
  });
});
