/**
 * A fuzz of our tokenizer and tree on REAL parts of the versioned fixtures, damaged, against rules written without any
 * other parser; its name keeps « differential » from the time each case was compared with another parser, and is not
 * renamed, by project rule.
 *
 * Damage: truncation, one character replaced, a slice removed or duplicated, an XML-significant snippet inserted
 * anywhere (`<`, `&`, a quote, `]]>`, a bad character reference, a CR, a stray end tag, a comment, a DOCTYPE, a CDATA
 * section, broken or whole), a comment, a processing instruction or a CDATA section inserted inside character data, one
 * whole tag removed, duplicated or renamed, text or a second root put outside the root element. The rules held:
 * - ours refuses only with our own `ReaderError("EXTRACT_FAILED", "xml: …")`; anything else is a crash, and fails;
 * - when ours accepts, a second, independent reading of the tokenizer's events finds the structure well-formed: tags
 *   pair up, at most one root (none only in a part with no text: read as empty, middle ground of 25/10/2026,
 *   `tree.test.ts`), no character data outside it. Without it, a tree that forgot a check would pass;
 * - when ours accepts, the text events of the tokenizer are the ones a SCANNER of the damaged text gives from the
 *   written rules alone (`expectedTexts`; the rules are pinned in `tokenizer.test.ts` and `tokenizer-markup.test.ts`): a
 *   comment joins the text around it, a processing instruction or a tag splits it, a CDATA section is a text of its
 *   own, CR LF and a lone CR become LF, a character reference becomes its character; and when the scanner finds
 *   something those rules refuse (a bare `&`, `]]>`, a character reference to U+0000, a surrogate or past U+10FFFF, a
 *   forbidden character, `--` in a comment, any `<!` but a comment or CDATA), ours must have refused. The scanner reads
 *   no tree and no event of ours; it is itself checked against our reading of every undamaged seed.
 * Ours refuses what no valid Office part holds (each case is written in the header of `tokenizer.ts`); a refusal is never
 * compared with another reader's.
 *
 * Seeds: for each kind of part (its name with digits dropped), the smallest one under 8 KiB that our reader reads, so a
 * run costs milliseconds. fast-check only chooses WHERE and HOW to damage them.
 */
import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { fixtureXmlParts } from "./xml-fixtures.test.support";
import { tokenize } from "./tokenizer";
import { eventsOf, readerAccepts } from "./xml-shadow.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(60_000, process.env) });

const SEED_MAX_CHARS = 8 * 1024;

let seeds: string[] = [];

beforeAll(async () => {
  // Streamed: only a part that could become its kind's smallest (under the cap, read by ours) is retained.
  const best = new Map<string, number>();
  const kept = await fixtureXmlParts((part) => {
    if (part.text.length > SEED_MAX_CHARS) return false;
    const kind = part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, "");
    const known = best.get(kind);
    if (known !== undefined && known <= part.text.length) return false;
    if (!readerAccepts(part.text)) return false;
    best.set(kind, part.text.length);
    return true;
  });
  // The filter keeps, per kind, a strictly shrinking run: the last part kept of a kind is its smallest.
  const smallest = new Map<string, string>();
  for (const part of kept) smallest.set(part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, ""), part.text);
  seeds = [...smallest.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, text]) => text);
}, 180_000);

const SNIPPETS = [
  "<",
  ">",
  "&",
  "&amp",
  "&#x0;",
  "&#xD800;",
  "&#1114112;",
  "&nbsp;",
  "&#X41;",
  '"',
  "'",
  "]]>",
  "\r",
  "\r\n",
  "</x>",
  "<x>",
  "<!-- c -->",
  "<!DOCTYPE a>",
  "<?pi?>",
  "\uFFFE",
  "\u0001",
  "=",
  "/>",
  // The markup the tokenizer reads: whole and broken.
  "<![CDATA[x]]>",
  "<![CDATA[",
  "<![CDATA[]]>",
  "<!--",
  "-->",
  "--",
  "<?",
  "?>",
  "<?pi data?>",
  '<?xml version="1.0"?>',
  '<!ENTITY e "x">',
  "<!x>",
];

type Damage =
  | { kind: "truncate"; where: number }
  | { kind: "replace"; where: number; by: number }
  | { kind: "remove"; where: number; length: number }
  | { kind: "duplicate"; where: number; length: number }
  | { kind: "insert"; where: number; snippet: number }
  | { kind: "insertInText"; where: number; snippet: number }
  | { kind: "tag"; where: number; action: "remove" | "duplicate" | "rename" }
  | { kind: "outside"; atEnd: boolean; snippet: number };

/** What may stand outside the root element (white space, a comment, a PI) and what may not. */
const OUTSIDE = [" ", "\r\n", "<!--c-->", "<?pi?>", "x", "&amp;", "<x/>", "<x></x>", "<![CDATA[ ]]>", "</x>"];

/** The [start, end) of every tag of `seed` (start, end or empty-element tag; never a comment, PI or CDATA). */
const tagSpans = (seed: string): [number, number][] => {
  const out: [number, number][] = [];
  for (let i = seed.indexOf("<"); i >= 0; i = seed.indexOf("<", i + 1)) {
    const next = seed[i + 1];
    if (next === "!" || next === "?") continue;
    const end = seed.indexOf(">", i);
    if (end < 0) break;
    out.push([i, end + 1]);
  }
  return out;
};

/** Markup that must NOT change the character data around it but by its own rule: a comment joins the
 *  text on both sides, an instruction and a CDATA section split it. */
const IN_TEXT = ["<!-- c -->", "<!---->", "<?pi?>", "<?pi d?>", "<![CDATA[x]]>", "<![CDATA[]]>", "&amp;<!--c-->lt;"];

/** The positions of `seed` strictly inside character data (both neighbours outside any markup). */
const textPositions = (seed: string): number[] => {
  const out: number[] = [];
  let inMarkup = false;
  for (let i = 0; i < seed.length; i++) {
    const c = seed[i];
    if (c === "<") inMarkup = true;
    else if (c === ">") inMarkup = false;
    else if (!inMarkup && i > 0 && seed[i - 1] !== ">" && seed[i - 1] !== ";" && c !== "&") out.push(i);
  }
  return out;
};

const fraction = fc.double({ min: 0, max: 1, noNaN: true });
const damage: fc.Arbitrary<Damage> = fc.oneof(
  fc.record({ kind: fc.constant("truncate" as const), where: fraction }),
  fc.record({ kind: fc.constant("replace" as const), where: fraction, by: fc.integer({ min: 0, max: 0x7f }) }),
  fc.record({ kind: fc.constant("remove" as const), where: fraction, length: fc.integer({ min: 1, max: 40 }) }),
  fc.record({ kind: fc.constant("duplicate" as const), where: fraction, length: fc.integer({ min: 1, max: 40 }) }),
  // Weighted: an inserted snippet is what reaches the constructs a reader must refuse.
  { arbitrary: fc.record({ kind: fc.constant("insert" as const), where: fraction, snippet: fc.nat() }), weight: 4 },
  // Markup put INSIDE character data, where joining or splitting the text is visible.
  {
    arbitrary: fc.record({ kind: fc.constant("insertInText" as const), where: fraction, snippet: fc.nat() }),
    weight: 2,
  },
  // The structure the tree checks.
  {
    arbitrary: fc.record({
      kind: fc.constant("tag" as const),
      where: fraction,
      action: fc.constantFrom("remove" as const, "duplicate" as const, "rename" as const),
    }),
    weight: 2,
  },
  fc.record({ kind: fc.constant("outside" as const), atEnd: fc.boolean(), snippet: fc.nat() }),
);

const apply = (seed: string, d: Damage): string => {
  const at = "where" in d ? Math.min(seed.length, Math.floor(d.where * seed.length)) : 0;
  switch (d.kind) {
    case "truncate":
      return seed.slice(0, at);
    case "replace":
      return seed.slice(0, at) + String.fromCharCode(d.by) + seed.slice(at + 1);
    case "remove":
      return seed.slice(0, at) + seed.slice(at + d.length);
    case "duplicate":
      return seed.slice(0, at + d.length) + seed.slice(at, at + d.length) + seed.slice(at + d.length);
    case "insert":
      return seed.slice(0, at) + SNIPPETS[d.snippet % SNIPPETS.length]! + seed.slice(at);
    case "insertInText": {
      const positions = textPositions(seed);
      if (positions.length === 0) return seed;
      const where = positions[Math.min(positions.length - 1, Math.floor(d.where * positions.length))]!;
      return seed.slice(0, where) + IN_TEXT[d.snippet % IN_TEXT.length]! + seed.slice(where);
    }
    case "tag": {
      const spans = tagSpans(seed);
      if (spans.length === 0) return seed;
      const [start, end] = spans[Math.min(spans.length - 1, Math.floor(d.where * spans.length))]!;
      if (d.action === "remove") return seed.slice(0, start) + seed.slice(end);
      if (d.action === "duplicate") return seed.slice(0, end) + seed.slice(start, end) + seed.slice(end);
      // Rename: a name character inserted after `<` or `</`, so that a start and its end tag no longer match.
      const at = seed[start + 1] === "/" ? start + 2 : start + 1;
      return seed.slice(0, at) + "z" + seed.slice(at);
    }
    case "outside": {
      const snippet = OUTSIDE[d.snippet % OUTSIDE.length]!;
      return d.atEnd ? seed + snippet : seed.replace(/^(\uFEFF?<\?xml[^>]*>)?/, (decl) => decl + snippet);
    }
  }
};

/** The independent reading: the tokenizer's events alone, judged against §2.1 [1] and §3 WFC Element Type Match. */
const wellStructured = (text: string): boolean => {
  const open: string[] = [];
  let roots = 0;
  let ok = true;
  const outside = (value: string, cdata: boolean): void => {
    if (open.length === 0 && (cdata || /[^ \t\r\n]/.test(value))) ok = false;
  };
  tokenize(text, {
    open(name, _attributes, selfClosing) {
      if (open.length === 0) roots++;
      if (!selfClosing) open.push(name);
    },
    close(name) {
      if (open.pop() !== name) ok = false;
    },
    text: (value) => outside(value, false),
    cdata: (value) => outside(value, true),
  });
  return ok && open.length === 0 && roots <= 1;
};

/** What character data, a comment and a CDATA section may not hold (§2.2 Char; pinned in `tokenizer.test.ts` and
 *  `tokenizer-markup.test.ts`): a control other than TAB, LF, CR, a lone surrogate, U+FFFE, U+FFFF. */
const FORBIDDEN =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

const normalisedLineEnds = (value: string): string => value.replace(/\r\n?/g, "\n");

const PREDEFINED: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const REFERENCE = /&(?:#([0-9]+)|#x([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));/y;

/** One run of raw character data, decoded by the written rules; `undefined` when a written rule refuses it. */
const decodedRun = (raw: string): string | undefined => {
  if (raw.includes("]]>") || FORBIDDEN.test(raw)) return undefined;
  const text = normalisedLineEnds(raw);
  let out = "";
  let from = 0;
  for (let amp = text.indexOf("&"); amp >= 0; amp = text.indexOf("&", from)) {
    out += text.slice(from, amp);
    REFERENCE.lastIndex = amp;
    const m = REFERENCE.exec(text);
    if (m === null) return undefined;
    if (m[3] !== undefined) out += PREDEFINED[m[3]]!;
    else {
      const code = m[1] !== undefined ? Number(m[1]) : parseInt(m[2]!, 16);
      if (code === 0 || (code >= 0xd800 && code <= 0xdfff) || code > 0x10ffff) return undefined;
      out += String.fromCodePoint(code);
    }
    from = amp + m[0].length;
  }
  return out + text.slice(from);
};

/** The index of the `>` that ends the tag starting at `start`, a `>` inside a quoted value not counting; -1 when none. */
const tagEnd = (source: string, start: number): number => {
  let quote = "";
  for (let i = start + 1; i < source.length; i++) {
    const c = source[i]!;
    if (quote !== "") {
      if (c === quote) quote = "";
    } else if (c === '"' || c === "'") quote = c;
    else if (c === ">") return i;
  }
  return -1;
};

/**
 * The text events inside the root that the written rules give for `source`: a SECOND reading, a scanner that builds
 * no tree and uses no event of ours. `undefined` when a written rule says the part must be refused.
 */
const expectedTexts = (source: string): string[] | undefined => {
  const out: string[] = [];
  let depth = 0;
  let run = "";
  const flush = (): void => {
    if (run !== "" && depth > 0) out.push(run);
    run = "";
  };
  let i = source.startsWith("\uFEFF") ? 1 : 0;
  while (i < source.length) {
    if (source[i] !== "<") {
      const next = source.indexOf("<", i);
      const end = next < 0 ? source.length : next;
      const decoded = decodedRun(source.slice(i, end));
      if (decoded === undefined) return undefined;
      run += decoded;
      i = end;
    } else if (source.startsWith("<!--", i)) {
      const end = source.indexOf("-->", i + 4);
      if (end < 0) return undefined;
      const body = source.slice(i + 4, end);
      if (body.includes("--") || body.endsWith("-") || FORBIDDEN.test(body)) return undefined;
      i = end + 3;
    } else if (source.startsWith("<?", i)) {
      const end = source.indexOf("?>", i + 2);
      if (end < 0) return undefined;
      flush();
      i = end + 2;
    } else if (source.startsWith("<![CDATA[", i)) {
      const end = source.indexOf("]]>", i + 9);
      if (end < 0) return undefined;
      const body = source.slice(i + 9, end);
      if (FORBIDDEN.test(body)) return undefined;
      flush();
      if (depth > 0) out.push(normalisedLineEnds(body));
      i = end + 3;
    } else if (source.startsWith("<!", i)) {
      return undefined;
    } else {
      const end = tagEnd(source, i);
      if (end < 0) return undefined;
      flush();
      if (source[i + 1] === "/") depth--;
      else if (source[end - 1] !== "/") depth++;
      i = end + 1;
    }
  }
  flush();
  return out;
};

/** The text events the tokenizer reports (inside the root), in order. */
const textsOfEvents = (source: string): string[] =>
  eventsOf(source).flatMap((event) => (event[0] === "text" ? [event[1]] : []));

describe("xml tokenizer — fuzz on damaged real parts", () => {
  it("found a seed for each common kind of part", () => {
    expect(seeds.length).toBeGreaterThan(20);
  });

  it("the text scanner gives, for every undamaged seed, the text events our tokenizer reports", () => {
    for (const seed of seeds) expect(expectedTexts(seed)).toEqual(textsOfEvents(seed));
  });

  it("the text scanner applies the written rules: what comment, PI and CDATA do to the text around them, and what is refused", () => {
    expect(expectedTexts(`<a>x<!--c-->y<?p?>z<![CDATA[]]><![CDATA[w\r\nv]]>&amp;&#65;&#x42;\r\n\r</a>`)).toEqual([
      "xy",
      "z",
      "",
      "w\nv",
      "&AB\n\n",
    ]);
    expect(expectedTexts(`<a>&amp;<!--c-->lt;</a>`)).toEqual(["&lt;"]);
    expect(expectedTexts(`<a>\r<!--c-->\n</a>`)).toEqual(["\n\n"]);
    expect(expectedTexts(`<a>]]<!--c-->></a>`)).toEqual(["]]>"]);
    expect(expectedTexts(`<a v="x>y"/>t`)).toEqual([]);
    for (const refused of [
      `<a>&</a>`,
      `<a>&nbsp;</a>`,
      `<a>&#X41;</a>`,
      `<a>&#0;</a>`,
      `<a>&#xD800;</a>`,
      `<a>&#1114112;</a>`,
      `<a>]]></a>`,
      `<a>\u0001</a>`,
      `<a><!--a--b--></a>`,
      `<a><!--x`,
      `<a><!DOCTYPE a></a>`,
      `<a><![CDATA[\uFFFF]]></a>`,
    ]) {
      expect(expectedTexts(refused), refused).toBeUndefined();
    }
  });

  it("no crash; a part ours accepts is well-structured and holds the text the written rules give", () => {
    let compared = 0;
    fc.assert(
      fc.property(fc.nat(), damage, (which, d) => {
        const text = apply(seeds[which % seeds.length]!, d);
        if (!readerAccepts(text)) return; // a crash of ours propagates and fails
        expect(wellStructured(text), `accepted after a ${d.kind} damage`).toBe(true);
        const expected = expectedTexts(text);
        expect(expected, `ours accepted what a written rule refuses, after a ${d.kind} damage`).toBeDefined();
        expect(textsOfEvents(text), `text read after a ${d.kind} damage`).toEqual(expected);
        compared++;
      }),
      { numRuns: fuzzRuns(2_000, process.env) },
    );
    // The oracle must reach texts it compares, or it proves nothing.
    expect(compared, "no accepted part had its texts compared").toBeGreaterThan(0);
  });
});
