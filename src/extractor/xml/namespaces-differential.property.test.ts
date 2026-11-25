/**
 * A fuzz of the NAMESPACE checks on REAL parts of the versioned fixtures, their namespaces damaged: a declaration
 * dropped, emptied or aliased, a prefix renamed, a reserved prefix or name declared, a colon put into a name; its name
 * keeps « differential » from the time each case was compared with another parser, and is not renamed, by project rule.
 * Only two damages are refused (product owner's rule, 25/10/2026: refuse only what would publish wrong): a malformed
 * qualified name, and two attributes of one element with the same canonical name (a value lost by `canonicalize`).
 * The others (a dropped, emptied or reserved declaration, an unbound prefix) must be READ.
 *
 * Rules held:
 * - any exception but our own refusal is a crash, and fails;
 * - a SECOND, independent reading of the tokenizer's events (a naive checker that copies the whole scope at every
 *   element, as `ooxml/xml.ts#canonicalize` does) must agree in both directions that matter: when ours accepts, it
 *   finds the part namespace-well-formed; when ours refuses with a namespace literal, it finds it is not. Without it,
 *   a check forgotten in `namespaces.ts` would pass, and so would a check that refuses too much;
 * - any OTHER refusal (the damage can break the XML itself, a name for instance) must be one of the literals of
 *   `OTHER_REFUSALS`, a closed list written from runs: a new, unwritten refusal turns the test red. What the second
 *   reading says of such a part is not asserted (a part ours refused for another reason may be namespace-well-formed).
 *
 * Seeds: for each kind of part (its name with digits dropped), the smallest one under 8 KiB that our reader reads and
 * which declares a namespace. fast-check only chooses WHERE and HOW to damage them.
 */
import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { renamed } from "../ooxml/xml";
import { tokenize } from "./tokenizer";
import { buildTree } from "./tree";
import { fixtureXmlParts } from "./xml-fixtures.test.support";
import { raisesRefusal } from "./xml-refusals.test.support";
import { readerAccepts } from "./xml-shadow.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(60_000, process.env) });

const SEED_MAX_CHARS = 8 * 1024;
const NAMESPACE_REASONS = ["xml: malformed qualified name", "xml: attributes repeated under their namespace"];
/** The other refusals ours gives on these damaged parts, from runs of this fuzz (3 runs of the whole test, each of 20 000 cases at FUZZ_RUNS_MULTIPLIER=10); each is a literal of `tokenizer.ts` or `tree.ts`. */
const OTHER_REFUSALS: readonly string[] = ["xml: repeated attribute"];
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

let seeds: string[] = [];

beforeAll(async () => {
  const best = new Map<string, number>();
  const kept = await fixtureXmlParts((part) => {
    if (part.text.length > SEED_MAX_CHARS || !part.text.includes("xmlns")) return false;
    const kind = part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, "");
    const known = best.get(kind);
    if (known !== undefined && known <= part.text.length) return false;
    if (!readerAccepts(part.text)) return false;
    best.set(kind, part.text.length);
    return true;
  });
  const smallest = new Map<string, string>();
  for (const part of kept) smallest.set(part.id.slice(part.id.indexOf("!") + 1).replace(/[0-9]+/g, ""), part.text);
  seeds = [...smallest.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, text]) => text);
}, 180_000);

/** The independent reading: is every name a QName, and do the attributes of each element keep distinct canonical names?
 *  A scope map copied at every element, regular expressions for the names. Only called on parts our tokenizer accepts (it throws otherwise, and the case is skipped). */
const namespaceWellFormed = (text: string): boolean => {
  const scopes: Map<string, string>[] = [new Map<string, string>()];
  const qname = /^(?:([^:]+):)?([^:]+)$/;
  const badStart = /^[-.0-9·̀-ͯ‿⁀]/;
  let ok = true;
  const wellNamed = (name: string): boolean => {
    const m = qname.exec(name);
    return m !== null && !badStart.test(m[2]!);
  };
  tokenize(text, {
    open(name, attributes, selfClosing) {
      const scope = new Map(scopes.at(-1));
      if (!wellNamed(name)) ok = false;
      for (const [attribute, value] of attributes) {
        if (!wellNamed(attribute)) ok = false;
        else if (attribute === "xmlns") scope.set("", value);
        else if (attribute.startsWith("xmlns:")) scope.set(attribute.slice(6), value);
      }
      const seen = new Set<string>();
      for (const [attribute] of attributes) {
        const canonical = renamed(attribute, scope, true);
        if (seen.has(canonical)) ok = false;
        seen.add(canonical);
      }
      if (!selfClosing) scopes.push(scope);
    },
    close() {
      scopes.pop();
    },
    text() {},
  });
  return ok;
};

type Damage =
  | {
      kind: "dropDeclaration" | "emptyDeclaration" | "renameAttributePrefix" | "aliasAttribute" | "colonInName";
      where: number;
    }
  | { kind: "reserved"; where: number; which: number };

const RESERVED = [
  `xmlns:xml="u"`,
  `xmlns:xmlns="u"`,
  `xmlns:p="${XMLNS_NS}"`,
  `xmlns="${XML_NS}"`,
  `xmlns:xml="${XML_NS}"`,
  `xmlns=""`,
];

/** The `[start, end)` of each start tag of `seed` (never an end tag, comment, PI or CDATA). */
const startTags = (seed: string): [number, number][] => {
  const out: [number, number][] = [];
  for (let i = seed.indexOf("<"); i >= 0; i = seed.indexOf("<", i + 1)) {
    if ("!?/".includes(seed[i + 1] ?? "")) continue;
    const end = seed.indexOf(">", i);
    if (end < 0) break;
    out.push([i, end + 1]);
  }
  return out;
};

const pick = <T>(items: readonly T[], where: number): T | undefined =>
  items[Math.min(items.length - 1, Math.floor(where * items.length))];

/** One damage inside one chosen start tag; the seed unchanged when the tag offers nothing to damage. */
const apply = (seed: string, d: Damage): string => {
  const tags = startTags(seed);
  const tag = pick(tags, d.where);
  if (!tag) return seed;
  const [start, end] = tag;
  const head = seed.slice(start, end);
  const declaration = /\sxmlns(?::[^\s=]+)?\s*=\s*("[^"]*"|'[^']*')/.exec(head);
  const prefixedAttribute = /\s([^\s=:<]+):([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/.exec(head);
  let damaged = head;
  switch (d.kind) {
    case "dropDeclaration":
      if (declaration) damaged = head.replace(declaration[0], "");
      break;
    case "emptyDeclaration":
      if (declaration) damaged = head.replace(declaration[0], declaration[0].replace(/("[^"]*"|'[^']*')$/, '""'));
      break;
    case "renameAttributePrefix":
      if (prefixedAttribute && prefixedAttribute[1] !== "xmlns")
        damaged = head.replace(
          prefixedAttribute[0],
          ` zq${prefixedAttribute[1]}:${prefixedAttribute[2]}=${prefixedAttribute[3]}`,
        );
      break;
    case "aliasAttribute": {
      // The same local name again, under an alias of the same namespace, found by the regex on the whole seed.
      if (!prefixedAttribute || prefixedAttribute[1] === "xmlns") break;
      const uri = new RegExp(`xmlns:${prefixedAttribute[1]}\\s*=\\s*"([^"]*)"`).exec(seed)?.[1];
      if (uri === undefined) break;
      damaged = head.replace(/(\/?>)$/, ` xmlns:zalias="${uri}" zalias:${prefixedAttribute[2]}="v"$1`);
      break;
    }
    case "colonInName":
      damaged = head.replace(
        /^<([^\s/>]+)/,
        (_m, name: string) => `<${name.includes(":") ? name.replace(":", "::") : `${name}:`}`,
      );
      break;
    case "reserved":
      damaged = head.replace(/(\/?>)$/, ` ${RESERVED[d.which % RESERVED.length]!}$1`);
      break;
  }
  return seed.slice(0, start) + damaged + seed.slice(end);
};

const fraction = fc.double({ min: 0, max: 1, noNaN: true });
const damage: fc.Arbitrary<Damage> = fc.oneof(
  fc.record({
    kind: fc.constantFrom(
      "dropDeclaration" as const,
      "emptyDeclaration" as const,
      "renameAttributePrefix" as const,
      "aliasAttribute" as const,
      "colonInName" as const,
    ),
    where: fraction,
  }),
  fc.record({ kind: fc.constant("reserved" as const), where: fraction, which: fc.nat() }),
);

/** Our refusal's message, or `undefined` when the part is read; a ReaderError that is not one of ours (`xml: …`) and any
 *  other throw are crashes and propagate. */
const refusalOf = (text: string): string | undefined => {
  try {
    buildTree(text, "document");
    return undefined;
  } catch (error) {
    if (error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: "))
      return error.message;
    throw error;
  }
};

describe("xml namespaces — fuzz of the namespace checks on real parts, namespaces damaged", () => {
  it("found a seed for each common kind of part", () => {
    expect(seeds.length).toBeGreaterThan(20);
  });

  it("accepts only what a second reading finds namespace-well-formed, refuses with a namespace literal only what it finds not, and refuses otherwise only with a written literal", () => {
    fc.assert(
      fc.property(fc.nat(), damage, (which, d) => {
        const text = apply(seeds[which % seeds.length]!, d);
        const refusal = refusalOf(text); // a crash of ours propagates and fails
        if (refusal === undefined) expect(namespaceWellFormed(text), `accepted after a ${d.kind} damage`).toBe(true);
        else if (NAMESPACE_REASONS.includes(refusal))
          expect(namespaceWellFormed(text), `${refusal} after a ${d.kind} damage`).toBe(false);
        else expect(OTHER_REFUSALS, `a refusal that is not written, after a ${d.kind} damage`).toContain(refusal);
      }),
      { numRuns: fuzzRuns(2_000, process.env) },
    );
  });
});

describe("xml namespaces — the closed lists of refusals", () => {
  it("every entry of OTHER_REFUSALS is still a literal our reader raises", () => {
    expect(OTHER_REFUSALS.filter((reason) => !raisesRefusal(reason))).toEqual([]);
  });

  it("every entry of NAMESPACE_REASONS is still a literal our reader raises", () => {
    expect(NAMESPACE_REASONS.filter((reason) => !raisesRefusal(reason))).toEqual([]);
  });
});
