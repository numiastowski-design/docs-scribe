/**
 * The one test of our XML reader that still uses `fast-xml-parser`, as an independent SECOND READING of damaged input:
 * on real parts of the versioned fixtures, damaged as text and as bytes, whenever both readers accept the same damaged
 * part they must build the same tree. Nothing else in the suite second-reads what our reader accepts on damaged input;
 * the other tests of the reader (literal assertions, `xml-trees.fixture.json`, the other `*.property.test.ts` files) do
 * not use `fast-xml-parser`. LIMIT: the reference is given the bytes as UTF-8 and, where that makes it disagree, OUR decoded
 * text (the `decodedAsUtf8` finding below), so this file never second-reads the DECODING of UTF-16 or of a legacy
 * encoding; that is held by literals in `xml/decode.test.ts`.
 *
 * What is held, on REAL parts (seeds: parts both readers already read alike) damaged in document mode (text and bytes)
 * and, on the golden storage XHTML, in fragment mode (the XHTML that `src/conservation/xlsx-page-cells.ts` reads):
 * - the reference is `fast-xml-parser` built with the options the callers really pass (captured from the module that
 *   reads the Office parts; the fragment parser accepts one set of options only), behind the product's DTD pre-check (`ooxml/xml.ts#parseXml`), in a comparison view that drops `?...` nodes
 *   and the text outside the root;
 * - the outcome of each case is one of: same tree, both refuse, ours alone refuses, the reference alone refuses, the
 *   reference loses (ours refuses, with a reason of `NOT_IN_A_VALID_OFFICE_PART`, what the reference reads), different
 *   trees;
 * - allowed: same, both refuse, the reference loses; « ours alone » only with a reason written in `WRITTEN_OURS_ALONE`;
 *   « the reference alone » and « different » only when a written finding RECOGNISES the case (`excuse`), never assumed:
 *   given what the finding says the reference gets wrong, the reference must then give OUR tree;
 * - the findings themselves are tested (they recognise their own cases and nothing else), and the tallies keep the fuzz
 *   from passing vacuously.
 * Crashes, structure, legal bytes and time are held without `fast-xml-parser` in the reader's own `xml/xml-differential.property.test.ts`: not
 * repeated here. A « crash » (any exception but our own refusal) still fails here, since it propagates.
 * fast-check only chooses the seed, WHERE and HOW to damage it. Positions are drawn without bias (`fc.noBias`): with the
 * default bias most damage lands at position 0, in the XML declaration.
 *
 * NOT covered: « the reference loses » is granted on the LITERAL of a listed refusal reason, never on proof that the
 * damaged text is malformed. If our reader refused a damaged text that stayed WELL-FORMED with a listed reason, this fuzz
 * would count it as a win.
 */
import { readFileSync } from "node:fs";

import { XMLParser as ReferenceParser } from "fast-xml-parser";
import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { publicFixturePath, publicFixtures } from "../public-fixtures.test.support";
import { decodeXmlPart } from "./decode";
import { buildTree, type XmlMode } from "./tree";
import { damaged, damages, damageText, textDamage, utf8, utf16Bytes, type Damage } from "./xml-damage.test.support";
import { fixtureXmlParts, type FixturePart, type XNode } from "./xml-fixtures.test.support";
import { raisesRefusal } from "./xml-refusals.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(120_000, process.env) });

/** A seed is a small part (one in eight, by id) or any part holding a comment or a processing instruction, capped. */
const SMALL_SEED_MAX_CHARS = 40_000;
const MARKUP_SEED_MAX_CHARS = 200_000;
/** A few LARGE real parts too. */
const LARGE_SEED_MIN_CHARS = 1_000_000;
const LARGE_SEED_MAX_CHARS = 2_000_000;
const LARGE_SEED_COUNT = 3;

/** Refusals of WELL-FORMED XML that ours may give alone, each a limit written where it is raised. */
const WRITTEN_OURS_ALONE: Readonly<Record<string, string>> = {
  "xml: name reserved by JavaScript objects":
    "`__proto__`, `constructor`, `toString`... never become object keys (the reference renames some to `__name`)",
  "xml: malformed qualified name":
    "a name the product's readers would never recognise (Namespaces in XML §4), its content lost",
  "xml: attributes repeated under their namespace": "two attributes `canonicalize` gives one name, a value lost",
  "xml: unknown encoding declared": "a declared encoding no decoder knows is refused, visible",
  "xml: XML declaration value too long": "our 64-character ceiling on a declaration value",
  "xml: XML declaration too long": "our 4 KB ceiling on the opening declaration",
  "xml: more attributes on one element than the limit": "MAX_XML_ATTRIBUTES",
  "xml: more nodes than the limit": "MAX_XML_NODES",
};

/** The refusals of ours that count as « the reference loses » when the reference reads the part: what no valid Office part
 *  can hold. Each is a violation of XML 1.0 well-formedness; the byte-level ones are bytes illegal in their encoding or a
 *  declaration the bytes contradict (§4.3.3 fatal errors); the structural ones are tags that do not pair up, a second root,
 *  text outside the root (§2.1 [1], §3 WFC Element Type Match). A declared encoding we cannot decode (« unknown encoding
 *  declared ») is NOT in the list: the part may be valid XML, so that refusal stays visible as « ours alone ». A DTD
 *  (« document type or markup declaration ») is in the list: refused for safety (entity expansion), as the product refuses
 *  it; ECMA-376 Part 2 §6.2.5 b) forbids it in the package's own parts, Part 1 says nothing for the others. Any other
 *  refusal of ours stays « ours alone »: names reserved by JavaScript objects, a refusal of VALID names written as a limit,
 *  and the ceilings on a declaration. */
const NOT_IN_A_VALID_OFFICE_PART: readonly string[] = [
  "xml: '--' inside a comment",
  "xml: '<' in an attribute value",
  "xml: ']]>' in character data",
  "xml: attribute value without quotes",
  "xml: attribute without a value",
  "xml: attributes not separated by white space",
  "xml: bytes not legal in their encoding",
  "xml: character reference out of range",
  "xml: character reference to a forbidden character",
  "xml: declared encoding does not match the bytes",
  "xml: document type or markup declaration",
  "xml: end tag does not match its start tag",
  "xml: end tag without a start tag",
  "xml: entity reference other than the five predefined ones",
  "xml: forbidden character",
  "xml: malformed XML declaration",
  "xml: malformed character reference",
  "xml: malformed end tag",
  "xml: malformed name",
  "xml: malformed processing instruction",
  "xml: malformed tag",
  "xml: more than one root element",
  "xml: repeated attribute",
  "xml: reserved processing instruction target",
  "xml: text outside the root element",
  "xml: UTF-16 without a byte-order mark",
  "xml: unterminated CDATA section",
  "xml: unterminated attribute value",
  "xml: unterminated comment",
  "xml: unterminated processing instruction",
  "xml: unclosed element",
  "xml: unterminated tag",
];

type Outcome = "same" | "bothRefuse" | "oursAlone" | "referenceAlone" | "referenceLoses" | "different";

type OurReading = { readonly tree: string; readonly encoding: string } | { readonly refused: string };

const isOurRefusal = (error: unknown): error is ReaderError =>
  error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: ");

/** Our whole reader on bytes (document mode) or on text (fragment mode); a crash propagates. */
const ours = (input: Uint8Array | string, mode: XmlMode): OurReading => {
  try {
    if (typeof input === "string") return { tree: JSON.stringify(buildTree(input, mode)), encoding: "UTF-8" };
    const { text, encoding } = decodeXmlPart(input);
    return { tree: JSON.stringify(buildTree(text, mode, encoding)), encoding };
  } catch (error) {
    if (isOurRefusal(error)) return { refused: error.message };
    throw error;
  }
};

// ——— The reference ———

type ReferenceOptions = ConstructorParameters<typeof ReferenceParser>[0];

interface ProductReference {
  /** The options `ooxml/xml.ts` passes to its parser, as constructed. */
  readonly documentOptions: ReferenceOptions;
  /** The raw `preserveOrder` output of the reference for this text behind the product's `parseXml` (its DTD guard and its
   *  canonicalisation included); `undefined` when `parseXml` throws, as it does for what the docx, pptx and xlsx readers refuse. */
  readonly productReading: (text: string) => XNode[] | undefined;
}

/**
 * The options a caller of the fragment parser writes. They are the ONLY ones `XmlFragmentParser` accepts: it throws on
 * any other key or value (`xml-parser.ts#assertSupportedOptions`, held by `xml-parser.test.ts`), so a caller outside
 * this package cannot pass others, and they need no capture.
 */
const fragmentOptions: ReferenceOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
  parseTagValue: false,
  htmlEntities: true,
};

/**
 * The reader of the Office parts keeps its parser and its options module-private: its options are therefore CAPTURED,
 * by loading the module afresh while OUR parser classes (`xml-parser.ts`) are replaced by recorders. A recorder keeps
 * the options it is given and answers with `fast-xml-parser` built with those very options. A copy of the options
 * would drift in silence; a capture cannot.
 */
const captureProductReference = async (): Promise<ProductReference> => {
  const actual = await vi.importActual<typeof import("./xml-parser")>("./xml-parser");
  const seen: ReferenceOptions[] = [];
  let lastOutput: unknown;
  class Recording {
    private readonly reference: ReferenceParser;
    constructor(options: ReferenceOptions) {
      seen.push(options);
      this.reference = new ReferenceParser(options);
    }
    parse(xmlData: string | Uint8Array): unknown {
      lastOutput = this.reference.parse(xmlData);
      return lastOutput;
    }
  }
  vi.resetModules();
  vi.doMock("./xml-parser", () => ({
    ...actual,
    XmlParser: Recording,
    XmlFragmentParser: Recording,
  }));
  try {
    const built = (): number => seen.length;
    const xml = await import("../ooxml/xml");
    if (built() !== 1) throw new Error(`ooxml/xml.ts built ${built()} parsers; the capture expects one`);
    const productReading = (text: string): XNode[] | undefined => {
      lastOutput = undefined;
      try {
        xml.parseXml(text);
      } catch {
        return undefined; // the package layer turns any throw into « damaged »
      }
      return lastOutput as XNode[];
    };
    return { documentOptions: seen[0], productReading };
  } finally {
    vi.doUnmock("./xml-parser");
    vi.resetModules();
  }
};

const tagOf = (node: XNode): string | undefined => Object.keys(node).find((k) => k !== ":@");

/** Drops `?...` nodes at every level and text nodes at the root, as both product callers drop them. */
const comparisonView = (nodes: readonly XNode[], root = true): XNode[] =>
  nodes.flatMap((node): XNode[] => {
    const tag = tagOf(node);
    if (tag === undefined || tag.startsWith("?")) return [];
    if (tag === "#text") return root ? [] : [node];
    const out: XNode = { [tag]: comparisonView(node[tag] as XNode[], false) };
    if (node[":@"] !== undefined) out[":@"] = node[":@"];
    return [out];
  });

/** The reference alone, with the given options, in the comparison view; `undefined` when it throws. */
const referenceView = (text: string, options: ReferenceOptions): string | undefined => {
  try {
    return JSON.stringify(comparisonView(new ReferenceParser(options).parse(text) as XNode[]));
  } catch {
    return undefined;
  }
};

let reference: ProductReference;

/** The product's reading (its DTD guard included) in the comparison view; `undefined` when it refuses. */
const productView = (text: string): string | undefined => {
  const nodes = reference.productReading(text);
  return nodes === undefined ? undefined : JSON.stringify(comparisonView(nodes));
};

/** What ours and the reference made of the same case. */
const classify = (mine: OurReading, theirs: string | undefined): Outcome => {
  if ("refused" in mine) {
    if (theirs === undefined) return "bothRefuse";
    return NOT_IN_A_VALID_OFFICE_PART.includes(mine.refused) ? "referenceLoses" : "oursAlone";
  }
  if (theirs === undefined) return "referenceAlone";
  return mine.tree === theirs ? "same" : "different";
};

/** The outcome of one document part. Ours reads `bytes` when given (decoded by `decode.ts`), else `text`. The reference is
 *  fed `text`, which callers build as the bytes decoded as UTF-8: a stand-in, since the product itself hands the BYTES to
 *  `ooxml/xml.ts#parseXml`, which decodes them with `decode.ts`. A part that is not plain UTF-8 therefore reaches the
 *  reference differently from how it reaches ours; the `decodedAsUtf8` finding recognises that case. */
const documentOutcome = (text: string, bytes?: Uint8Array): Outcome =>
  classify(ours(bytes ?? text, "document"), productView(text));

// ——— The written findings ———

const PI_SPAN = /<\?[\s\S]*?\?>/g;
/** Every processing instruction but the declaration emptied of its data, its target kept. */
const withoutPiData = (text: string): string =>
  text.replace(PI_SPAN, (pi) => (/^<\?xml[\s?]/.test(pi) ? pi : `${/^<\?[^\s?]*/.exec(pi)![0]}?>`));

/** The product's DTD guard (`ooxml/xml.ts#DECLARES_DTD`), matched the same way; the product tests it on the text decoded
 *  by `decode.ts`, a comment or a CDATA section included. Here it is tested on the text given to the reference. */
const DECLARES_DTD = /<!DOCTYPE/i;

type Finding = "doctypeInMarkup" | "piDataAsAttributes" | "decodedAsUtf8" | "nameCharReadAsSpace";

interface ExcuseInput {
  /** The text the reference is first given: the bytes decoded as UTF-8 (the product itself decodes them with `decode.ts`). */
  readonly productText: string;
  /** The bytes ours reads; absent in fragment mode (storage XHTML is text). */
  readonly bytes?: Uint8Array;
  readonly ourTree: string;
  /** Reads a text as the reference does, in the comparison view; `undefined` when it refuses. */
  readonly reference: (text: string) => string | undefined;
  /** The reference alone, without the product's DTD guard (document mode only). */
  readonly unguarded?: (text: string) => string | undefined;
}

/** The two XML 1.0 NameChars that JavaScript calls white space (§2.3 [4]: #x37F-#x1FFF and #xFDF0-#xFFFD). */
const NAME_CHARS_READ_AS_SPACE = /[\u1680\uFEFF]/g;
/** A quoted string of our serialized tree holding one of them and followed by a colon: meant for an element or attribute
 *  name, but it is a text filter, so it may fire on a text value too (a false positive only adds one more reading to try). */
const NAME_WITH_SUCH_CHAR = /"[^"]*[\u1680\uFEFF][^"]*":/;
/** Each of them replaced by a plain name character, a leading byte-order mark kept. The replacement is made over the WHOLE
 *  text, character data and attribute values included, not only names (the same one is made in our tree, so they stay equal). */
const plainNameChars = (text: string): string =>
  text.startsWith("\uFEFF")
    ? `\uFEFF${text.slice(1).replace(NAME_CHARS_READ_AS_SPACE, "_")}`
    : text.replace(NAME_CHARS_READ_AS_SPACE, "_");

interface Reading {
  readonly findings: readonly Finding[];
  readonly text: string;
  readonly tree: string;
}

/** Each reading a written finding (or several) says the reference SHOULD have made, beside the plain one. */
const readingsUnderFindings = ({ productText, bytes, ourTree }: ExcuseInput): Reading[] => {
  let readings: Reading[] = [{ findings: [], text: productText, tree: ourTree }];
  if (bytes) {
    const decoded = decodeXmlPart(bytes).text;
    if (decoded !== productText) readings.push({ findings: ["decodedAsUtf8"], text: decoded, tree: ourTree });
  }
  readings = readings.flatMap((r) => {
    const stripped = withoutPiData(r.text);
    return stripped === r.text
      ? [r]
      : [r, { findings: [...r.findings, "piDataAsAttributes"], text: stripped, tree: r.tree }];
  });
  if (!NAME_WITH_SUCH_CHAR.test(ourTree)) return readings;
  return readings.flatMap((r) => [
    r,
    {
      findings: [...r.findings, "nameCharReadAsSpace"],
      text: plainNameChars(r.text),
      tree: r.tree.replace(NAME_CHARS_READ_AS_SPACE, "_"),
    },
  ]);
};

/**
 * The written findings that explain why the reference refuses or differs (joined by `+` when several hold), or
 * `undefined`. Each finding is a VALID part the reference reads wrong (written for the humans, never copied, never
 * silently fixed), and is RECOGNISED, never assumed: the reference, given what the finding says it gets wrong, must then
 * give our tree.
 * - `doctypeInMarkup`: a `<!DOCTYPE` inside a comment, a CDATA section or a PI's data is valid XML; the product's DTD
 *   guard (a regular expression on the decoded text, `ooxml/xml.ts#DECLARES_DTD`) refuses it. Recognised when the reference WITHOUT the guard gives our tree.
 * - `piDataAsAttributes`: the reference reads a PI's data as pseudo-attributes, so a quote keeps `?>` from ending it, and a
 *   pseudo-attribute `__proto__` throws; XML 1.0 §2.6 [16] makes the data opaque. Recognised when, every PI's data taken
 *   out, the reference gives our tree.
 * - `decodedAsUtf8`: the product hands a part's BYTES to `ooxml/xml.ts#parseXml`, which decodes them with `decode.ts`
 *   (BOM, declaration, UTF-16, legacy encodings); this file feeds the reference the bytes decoded as UTF-8, so a part in
 *   UTF-16 or in a declared legacy encoding reaches it as gibberish. Recognised when the reference, given the text as
 *   `decode.ts` decodes the bytes, gives our tree. That makes OUR decoding the reference's input: it is not second-read here.
 * - `nameCharReadAsSpace`: U+FEFF and U+1680 are NameChars (§2.3 [4]); the reference takes them for white space, so a name
 *   holding one is cut, emptied (`<\uFEFFa/>` gives an element named ``) or loses an attribute (`b\uFEFFc="1"` gives `c`).
 *   Recognised when, both replaced by `_`, the reference gives our tree. The guard and the replacement act on the whole
 *   text, so they also touch character data and attribute values.
 */
const excuse = (input: ExcuseInput): string | undefined => {
  const { reference: read, unguarded } = input;
  for (const r of readingsUnderFindings(input)) {
    if (r.findings.length > 0 && read(r.text) === r.tree) return r.findings.join("+");
    if (unguarded && DECLARES_DTD.test(r.text) && unguarded(r.text) === r.tree)
      return [...r.findings, "doctypeInMarkup"].join("+");
  }
  return undefined;
};

// ——— Seeds ———

interface Seed {
  readonly id: string;
  readonly text: string;
}

let documentSeeds: Seed[] = [];
let largeSeeds: Seed[] = [];
let fragmentSeeds: Seed[] = [];

const HOLDS_MARKUP = /<!--|<!\[CDATA\[|<\?(?!xml[\s?])/;

beforeAll(async () => {
  reference = await captureProductReference();
  let small = 0;
  const large = (part: FixturePart): boolean =>
    part.text.length > LARGE_SEED_MIN_CHARS && part.text.length <= LARGE_SEED_MAX_CHARS;
  const isSeed = (part: FixturePart): boolean => {
    const markup = part.text.length <= MARKUP_SEED_MAX_CHARS && HOLDS_MARKUP.test(part.text);
    const pick = markup || large(part) || (part.text.length <= SMALL_SEED_MAX_CHARS && small++ % 8 === 0);
    // The damage starts from agreement: only parts both readers read alike are seeds.
    return pick && documentOutcome(part.text, part.bytes) === "same";
  };
  const kept = await fixtureXmlParts(isSeed);
  const seedOf = (part: FixturePart): Seed => ({ id: part.id, text: part.text });
  documentSeeds = kept.filter((part) => !large(part)).map(seedOf);
  largeSeeds = kept
    .filter(large)
    .sort((a, b) => b.text.length - a.text.length || a.id.localeCompare(b.id))
    .slice(0, LARGE_SEED_COUNT)
    .map(seedOf);
  fragmentSeeds = publicFixtures(".xhtml")
    .filter((path) => ["docx", "gdoc", "pdf", "pptx"].some((dir) => path.startsWith(`golden/${dir}/`)))
    .map((path) => ({ id: path, text: readFileSync(publicFixturePath(path), "utf8") }));
}, 300_000);

// ——— The rules ———

/** The index of a seed, drawn without bias (`fc.nat()` leans toward the smallest values). */
const seedIndex = (seeds: readonly Seed[]): fc.Arbitrary<number> => fc.noBias(fc.nat({ max: seeds.length - 1 }));

const describeDamages = (ds: readonly Damage[]): string => ds.map((d) => d.kind).join("+");

interface Tally {
  readonly outcomes: Record<string, number>;
  readonly findings: Record<string, number>;
}

const newTally = (): Tally => ({ outcomes: {}, findings: {} });
const bump = (counts: Record<string, number>, key: string): void => {
  counts[key] = (counts[key] ?? 0) + 1;
};

/** The rules for one outcome; `explain` names the finding when the reference refuses alone or differs. */
const judge = (
  outcome: Outcome,
  mine: OurReading,
  explain: () => string | undefined,
  where: string,
  tally: Tally,
): void => {
  bump(tally.outcomes, outcome);
  if (outcome === "oursAlone") {
    const reason = "refused" in mine ? mine.refused : "";
    expect(Object.hasOwn(WRITTEN_OURS_ALONE, reason), `ours alone without a written reason: ${reason} (${where})`).toBe(
      true,
    );
  }
  if (outcome !== "referenceAlone" && outcome !== "different") return;
  const finding = explain();
  if (finding === undefined) expect.fail(`${outcome}, no written finding (${where})`);
  bump(tally.findings, finding);
};

/** One damaged document case: ours and the reference on the same damaged part, then the rules. */
const checkDocumentCase = (seed: Seed, ds: readonly Damage[], tally: Tally): void => {
  const bytes = damaged(seed.text, ds);
  const productText = utf8.decode(bytes); // what the reference is first given (see `documentOutcome`)
  const where = `${seed.id}, ${describeDamages(ds)}`;
  const mine = ours(bytes, "document");
  const outcome = classify(mine, productView(productText));
  const explain = (): string | undefined =>
    "tree" in mine
      ? excuse({
          productText,
          bytes,
          ourTree: mine.tree,
          reference: productView,
          unguarded: (text) => referenceView(text, reference.documentOptions),
        })
      : undefined;
  judge(outcome, mine, explain, where, tally);
};

describe("xml reader — comparative fuzz against fast-xml-parser on damaged real parts", () => {
  it("found many seeds, the parts with markup among them", () => {
    expect(documentSeeds.length).toBeGreaterThan(400);
    expect(documentSeeds.filter((seed) => HOLDS_MARKUP.test(seed.text)).length).toBeGreaterThan(5);
    // A floor: adding a golden must not turn this test red.
    expect(fragmentSeeds.length).toBeGreaterThanOrEqual(6);
  });

  it("document mode, text and bytes damaged: only written refusals of ours, only recognised findings of the reference", () => {
    const tally = newTally();
    fc.assert(
      fc.property(seedIndex(documentSeeds), damages, (which, ds) =>
        checkDocumentCase(documentSeeds[which]!, ds, tally),
      ),
      { numRuns: fuzzRuns(1_000, process.env) },
    );
    // The fuzz must reach agreement and refusals of both kinds, or it proves nothing.
    expect(tally.outcomes.same ?? 0).toBeGreaterThan(0);
    expect(tally.outcomes.referenceLoses ?? 0).toBeGreaterThan(0);
    expect(tally.outcomes.bothRefuse ?? 0).toBeGreaterThan(0);
  });

  it("large real parts (1 to 2 M characters), damaged: the same rules", () => {
    expect(largeSeeds.length).toBe(LARGE_SEED_COUNT);
    const tally = newTally();
    fc.assert(
      fc.property(seedIndex(largeSeeds), damages, (which, ds) => checkDocumentCase(largeSeeds[which]!, ds, tally)),
      // endOnFailure: a shrink on 1 to 2 M characters would reach the timeout before the real message.
      { numRuns: fuzzRuns(12, process.env), endOnFailure: true },
    );
  });

  it("fragment mode (storage XHTML of xlsx-page-cells.ts), text damaged: the same rules", () => {
    const tally = newTally();
    const read = (text: string): string | undefined => referenceView(text, fragmentOptions);
    fc.assert(
      fc.property(seedIndex(fragmentSeeds), fc.array(textDamage, { minLength: 1, maxLength: 3 }), (which, ds) => {
        const seed = fragmentSeeds[which]!;
        const text = ds.reduce(damageText, seed.text);
        const where = `${seed.id}, ${describeDamages(ds)}`;
        const mine = ours(text, "fragment");
        const explain = (): string | undefined =>
          "tree" in mine ? excuse({ productText: text, ourTree: mine.tree, reference: read }) : undefined;
        judge(classify(mine, read(text)), mine, explain, where, tally);
      }),
      { numRuns: fuzzRuns(500, process.env) },
    );
    expect(tally.outcomes.same ?? 0).toBeGreaterThan(0);
    expect(tally.outcomes.referenceLoses ?? 0).toBeGreaterThan(0);
  });

  const readsAlike = (id: string, text: string): void => {
    expect(ours(text, "fragment"), id).toEqual({
      tree: referenceView(text, fragmentOptions),
      encoding: "UTF-8",
    });
  };

  it("the undamaged storage pages read the same in both readers (the damage starts from agreement)", () => {
    for (const seed of fragmentSeeds) readsAlike(seed.id, seed.text);
  });

  // The storage pages of the real workbooks, rendered as published, are compared in the application that renders them
  // (`src/conservation/xlsx-page-cells.reference.test.ts` there): this package renders no page.
});

describe("xml reader — the written findings recognise their own cases, and nothing else", () => {
  const documentExcuse = (text: string, bytes?: Uint8Array): string | undefined => {
    const mine = ours(bytes ?? new TextEncoder().encode(text), "document");
    if (!("tree" in mine)) throw new Error("expected ours to read the case");
    return excuse({
      productText: text,
      ...(bytes ? { bytes } : {}),
      ourTree: mine.tree,
      reference: productView,
      unguarded: (t) => referenceView(t, reference.documentOptions),
    });
  };

  it("a DOCTYPE inside a comment: refused by the product's guard alone", () => {
    const text = `<a><!--<!DOCTYPE x>-->t</a>`;
    expect(documentOutcome(text)).toBe("referenceAlone");
    expect(documentExcuse(text)).toBe("doctypeInMarkup");
  });

  it("a PI's data read as pseudo-attributes by the reference (XML 1.0 §2.6 [16]: opaque)", () => {
    for (const text of [
      `<?mTemplates xmlns="mso-contentType?><FormTemplates xmlns="urn:f"/>`,
      `<a><?pi x="?>"?>t</a>`,
      `<?pi __proto__="x"?><a/>`,
    ]) {
      expect(documentOutcome(text), text).not.toBe("same");
      expect(documentExcuse(text), text).toBe("piDataAsAttributes");
    }
  });

  it("a valid UTF-16 part, fed to the reference as UTF-8", () => {
    const text = `<?xml version="1.0" encoding="UTF-16"?><a>é</a>`;
    const bytes = utf16Bytes(`\uFEFF${text}`, false);
    expect(documentOutcome(utf8.decode(bytes), bytes)).not.toBe("same");
    expect(documentExcuse(utf8.decode(bytes), bytes)).toBe("decodedAsUtf8");
  });

  it("a name holding U+FEFF or U+1680, NameChars that the reference takes for white space", () => {
    for (const text of [`<a b\uFEFFc="1"/>`, `<\uFEFFa/>`, `<a \u1680b="1"/>`]) {
      expect(documentOutcome(text), JSON.stringify(text)).toBe("different");
      expect(documentExcuse(text), JSON.stringify(text)).toBe("nameCharReadAsSpace");
    }
    // In character data both keep it: nothing to excuse.
    expect(documentOutcome(`<a>x\uFEFFy\u1680</a>`)).toBe("same");
  });

  it("findings combine: a PI with a quote in a part declared ISO-8859-1", () => {
    const bytes = Uint8Array.from(`<?xml version="1.0" encoding="ISO-8859-1"?><a><?pi it's?>\u00e9</a>`, (c) =>
      c.charCodeAt(0),
    );
    expect(documentExcuse(utf8.decode(bytes), bytes)).toBe("decodedAsUtf8+piDataAsAttributes");
  });

  it("nothing else is excused: an unclosed comment, a tree that differs", () => {
    expect(
      excuse({ productText: `<a/><!--x`, ourTree: JSON.stringify([{ a: [] }]), reference: productView }),
    ).toBeUndefined();
    expect(
      excuse({
        productText: `<?pi?><a>x</a>`,
        ourTree: JSON.stringify([{ a: [{ "#text": "y" }] }]),
        reference: productView,
      }),
    ).toBeUndefined();
    const bytes = new TextEncoder().encode(`<a>x</a>`);
    expect(
      excuse({ productText: `<a>x</a>`, bytes, ourTree: JSON.stringify([{ a: [] }]), reference: productView }),
    ).toBeUndefined();
  });

  it("every entry of WRITTEN_OURS_ALONE and of NOT_IN_A_VALID_OFFICE_PART is still a literal our reader raises, and no reason is in both", () => {
    expect(Object.keys(WRITTEN_OURS_ALONE).filter((reason) => !raisesRefusal(reason))).toEqual([]);
    expect(NOT_IN_A_VALID_OFFICE_PART.filter((reason) => !raisesRefusal(reason))).toEqual([]);
    for (const reason of Object.keys(WRITTEN_OURS_ALONE))
      expect(NOT_IN_A_VALID_OFFICE_PART, reason).not.toContain(reason);
  });
});
