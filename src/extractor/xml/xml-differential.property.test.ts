/**
 * The GENERAL fuzz of our whole reader (`decode.ts`, then `tree.ts` with its guards and namespace checks) on REAL parts
 * of the versioned fixtures, damaged as TEXT and as BYTES; its name keeps « differential » from the time each case was
 * compared with another parser (that comparison now lives in `xml-reference-differential.property.test.ts`), and
 * is not renamed, by project rule. The other fuzzes each aim at one guard
 * (`tokenizer-`, `guards-`, `namespaces-differential.property.test.ts`, text only, the smallest part of each kind);
 * this one takes many more parts, damages the bytes the zip hands over (UTF-8 sequences, the mark, the declared
 * encoding, UTF-16), and also fuzzes the storage XHTML that `src/conservation/xlsx-page-cells.ts` reads (fragment mode).
 *
 * Rules held on every damaged case:
 * - any exception but our own refusal (`ReaderError`, `EXTRACT_FAILED`, `xml: …`) is a crash, and fails; and our refusal is
 *   one of the literals of `WRITTEN_REFUSALS`, a closed list written from runs, so that a new, unwritten refusal turns the test red;
 * - when ours accepts, two independent readings agree: the bytes are legal in the encoding found (a fatal decoder: no
 *   gibberish: independent of ours), and the tokenizer's events pair up under at most one root, with no text outside it
 *   (`wellStructured` reuses OUR `tokenize`: independent of `tree.ts` only);
 * - our reading of one case costs less than `SLOW_CASE_MS`: a slow case is a failure.
 * fast-check only chooses the seed, WHERE and HOW to damage it. Positions are drawn without bias (`fc.noBias`): with the
 * default bias most damage lands at position 0, in the XML declaration.
 */
import { readFileSync } from "node:fs";

import fc from "fast-check";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import { fuzzRuns, fuzzTestTimeout } from "../../lib/docs-fuzz-runs.test.support";
import { publicFixturePath, publicFixtures } from "../public-fixtures.test.support";
import { decodeXmlPart } from "./decode";
import { tokenize } from "./tokenizer";
import { buildTree, type XmlMode } from "./tree";
import { damageBytes, damaged, damages, damageText, textDamage, type Damage } from "./xml-damage.test.support";
import { fixtureXmlParts, type FixturePart } from "./xml-fixtures.test.support";
import { raisesRefusal } from "./xml-refusals.test.support";
import { readerAccepts } from "./xml-shadow.test.support";

vi.setConfig({ testTimeout: fuzzTestTimeout(120_000, process.env) });

/** A seed is a small part (one in eight, by id) or any part holding a comment or a processing instruction, capped. */
const SMALL_SEED_MAX_CHARS = 40_000;
const MARKUP_SEED_MAX_CHARS = 200_000;
/** A few LARGE real parts too, so that a walk quadratic in the size of a part shows as a slow case. */
const LARGE_SEED_MIN_CHARS = 1_000_000;
const LARGE_SEED_MAX_CHARS = 2_000_000;
const LARGE_SEED_COUNT = 3;
/** The time our reading of one case may take: 2 s, or 2.5 s per million characters above 800 000. Our reader costs about
 *  0.1 s per million characters under vitest (a part of 2 M characters: 0.1-0.3 s), so the budget is a wide margin, never
 *  a measurement (that one is taken on an esbuild bundle); a walk quadratic in the size of a 2 M-character part costs
 *  minutes. */
const SLOW_CASE_MS = 2_000;
const SLOW_CASE_MS_PER_MILLION_CHARS = 2_500;
const budgetMs = (chars: number): number =>
  Math.max(SLOW_CASE_MS, (chars / 1_000_000) * SLOW_CASE_MS_PER_MILLION_CHARS);

type OurReading = { readonly tree: string; readonly encoding: string } | { readonly refused: string };

/** Every refusal ours gave on these damaged parts, document and fragment mode, in 3 runs of this file at FUZZ_RUNS_MULTIPLIER=10 (37 literals, the same in each run);
 *  each is a literal of `tokenizer.ts`, `decode.ts`, `tree.ts` or `namespaces.ts`. A refusal outside the list is a new behaviour: write it here, or fix it. */
const WRITTEN_REFUSALS: readonly string[] = [
  "xml: '--' inside a comment",
  "xml: '<' in an attribute value",
  "xml: ']]>' in character data",
  "xml: UTF-16 without a byte-order mark",
  "xml: XML declaration too long",
  "xml: XML declaration value too long",
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
  "xml: malformed qualified name",
  "xml: malformed tag",
  "xml: more than one root element",
  "xml: name reserved by JavaScript objects",
  "xml: repeated attribute",
  "xml: reserved processing instruction target",
  "xml: text outside the root element",
  "xml: unclosed element",
  "xml: unknown encoding declared",
  "xml: unterminated CDATA section",
  "xml: unterminated attribute value",
  "xml: unterminated comment",
  "xml: unterminated processing instruction",
  "xml: unterminated tag",
];

const isOurRefusal = (error: unknown): error is ReaderError =>
  error instanceof ReaderError && error.code === "EXTRACT_FAILED" && error.message.startsWith("xml: ");

/** Our whole reader on bytes (document mode) or on text (fragment mode); a crash propagates. */
const ours = (input: Uint8Array | string, mode: XmlMode): OurReading => {
  try {
    if (typeof input === "string") return { tree: JSON.stringify(buildTree(input, mode)), encoding: "UTF-8" };
    const { text, encoding } = decodeXmlPart(input);
    return { tree: JSON.stringify(buildTree(text, mode, encoding)), encoding };
  } catch (error) {
    if (isOurRefusal(error)) {
      expect(WRITTEN_REFUSALS, "a refusal that is not written").toContain(error.message);
      return { refused: error.message };
    }
    throw error;
  }
};

/** The structure reading of an accepted text, independent of `tree.ts` only (it reuses OUR `tokenize`): tags pair up, at
 *  most one root (document), no text outside. */
const wellStructured = (text: string, encoding: string, mode: XmlMode): boolean => {
  const open: string[] = [];
  let roots = 0;
  let ok = true;
  const outside = (value: string, cdata: boolean): void => {
    if (mode === "document" && open.length === 0 && (cdata || /[^ \t\r\n]/.test(value))) ok = false;
  };
  tokenize(
    text,
    {
      open(name, _attributes, selfClosing) {
        if (open.length === 0) roots++;
        if (!selfClosing) open.push(name);
      },
      close(name) {
        if (open.pop() !== name) ok = false;
      },
      text: (value) => outside(value, false),
      cdata: (value) => outside(value, true),
    },
    encoding,
  );
  return ok && open.length === 0 && (mode === "fragment" || roots <= 1);
};

/** The independent byte reading: UTF-8 bytes ours accepted are legal UTF-8 (a fatal decoder), never gibberish. */
const legalUtf8 = (bytes: Uint8Array): boolean => {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
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
  let small = 0;
  const large = (part: FixturePart): boolean =>
    part.text.length > LARGE_SEED_MIN_CHARS && part.text.length <= LARGE_SEED_MAX_CHARS;
  const kept = await fixtureXmlParts((part: FixturePart) => {
    const markup = part.text.length <= MARKUP_SEED_MAX_CHARS && HOLDS_MARKUP.test(part.text);
    const pick = markup || large(part) || (part.text.length <= SMALL_SEED_MAX_CHARS && small++ % 8 === 0);
    return pick && readerAccepts(part.text, part.bytes);
  });
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
  accepted: number;
  refused: number;
}

/** Ours on one input, timed: a slow case fails. */
const timedOurs = (input: Uint8Array | string, mode: XmlMode, where: string): OurReading => {
  const started = performance.now();
  const mine = ours(input, mode);
  const spent = performance.now() - started;
  expect(spent, `slow case (${where})`).toBeLessThan(budgetMs(input.length));
  return mine;
};

/** One damaged document case: ours timed, then the independent readings when ours accepts. */
const checkDocumentCase = (seed: Seed, ds: readonly Damage[], tally: Tally): void => {
  const bytes = damaged(seed.text, ds);
  const where = `${seed.id}, ${describeDamages(ds)}`;
  const mine = timedOurs(bytes, "document", where);
  if ("refused" in mine) {
    tally.refused++;
    return;
  }
  tally.accepted++;
  if (mine.encoding === "UTF-8") expect(legalUtf8(bytes), `accepted bytes that are not UTF-8 (${where})`).toBe(true);
  expect(
    wellStructured(decodeXmlPart(bytes).text, mine.encoding, "document"),
    `accepted a broken structure (${where})`,
  ).toBe(true);
};

describe("xml reader — general fuzz on damaged real parts", () => {
  it("found many seeds, the parts with markup among them", () => {
    expect(documentSeeds.length).toBeGreaterThan(400);
    expect(documentSeeds.filter((seed) => HOLDS_MARKUP.test(seed.text)).length).toBeGreaterThan(5);
    // A floor, as in `tree.test.ts` (10/11/2026): adding a golden must not turn this test red.
    expect(fragmentSeeds.length).toBeGreaterThanOrEqual(6);
  });

  it("document mode, text and bytes damaged: no crash, accepted parts are legal and well-structured", () => {
    const tally: Tally = { accepted: 0, refused: 0 };
    fc.assert(
      fc.property(seedIndex(documentSeeds), damages, (which, ds) =>
        checkDocumentCase(documentSeeds[which]!, ds, tally),
      ),
      { numRuns: fuzzRuns(1_000, process.env) },
    );
    // The fuzz must reach both an accepted and a refused part, or it proves nothing.
    expect(tally.accepted).toBeGreaterThan(0);
    expect(tally.refused).toBeGreaterThan(0);
  });

  it("large real parts (1 to 2 M characters), damaged: the same rules, and no slow case", () => {
    expect(largeSeeds.length).toBe(LARGE_SEED_COUNT);
    const tally: Tally = { accepted: 0, refused: 0 };
    fc.assert(
      fc.property(seedIndex(largeSeeds), damages, (which, ds) => checkDocumentCase(largeSeeds[which]!, ds, tally)),
      // endOnFailure: a shrink on 1 to 2 M characters would reach the timeout before the real message.
      { numRuns: fuzzRuns(12, process.env), endOnFailure: true },
    );
  });

  it("fragment mode (storage XHTML of xlsx-page-cells.ts), text damaged: the same rules", () => {
    const tally: Tally = { accepted: 0, refused: 0 };
    fc.assert(
      fc.property(seedIndex(fragmentSeeds), fc.array(textDamage, { minLength: 1, maxLength: 3 }), (which, ds) => {
        const seed = fragmentSeeds[which]!;
        const text = ds.reduce(damageText, seed.text);
        const where = `${seed.id}, ${describeDamages(ds)}`;
        const mine = timedOurs(text, "fragment", where);
        if ("refused" in mine) {
          tally.refused++;
          return;
        }
        tally.accepted++;
        expect(wellStructured(text, "UTF-8", "fragment"), `accepted a broken structure (${where})`).toBe(true);
      }),
      { numRuns: fuzzRuns(500, process.env) },
    );
    expect(tally.accepted).toBeGreaterThan(0);
    expect(tally.refused).toBeGreaterThan(0);
  });

  it("the undamaged seeds are read by ours (the damage starts from an accepted part)", () => {
    for (const seed of fragmentSeeds) expect(ours(seed.text, "fragment"), seed.id).toHaveProperty("tree");
  });
});

describe("xml reader — the damage helpers keep what they do not damage", () => {
  it("the declared-encoding damage rewrites the declaration only: a byte 0x80-0x9F (a UTF-8 continuation) stays intact", () => {
    const original = new TextEncoder().encode(`<?xml version="1.0"?><a>’</a>`);
    expect([...original.subarray(original.indexOf(0xe2), original.indexOf(0xe2) + 3)]).toEqual([0xe2, 0x80, 0x99]);
    const declared = damageBytes(original, { kind: "declare", label: "windows-1252" });
    expect(Buffer.from(declared).toString("latin1")).toBe(
      Buffer.from(original)
        .toString("latin1")
        .replace(`<?xml version="1.0"?>`, `<?xml version="1.0" encoding="windows-1252"?>`),
    );
    const marked = damageBytes(Uint8Array.from([0xef, 0xbb, 0xbf, ...original]), {
      kind: "declare",
      label: "ISO-8859-1",
    });
    expect([...marked.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(Buffer.from(marked).includes(Buffer.from([0xe2, 0x80, 0x99]))).toBe(true);
  });
});

describe("xml reader — the closed list of written refusals", () => {
  it("every entry of WRITTEN_REFUSALS is still a literal our reader raises", () => {
    expect(WRITTEN_REFUSALS.filter((reason) => !raisesRefusal(reason))).toEqual([]);
  });
});
