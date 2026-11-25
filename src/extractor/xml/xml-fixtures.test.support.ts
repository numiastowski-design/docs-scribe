/** Test support: every XML part of the public Office documents of `fixtures/` (`publicFixtures`), decoded as UTF-8
 *  (our own reader decodes the bytes itself), the options the callers pass to our parser classes (the reader's own
 *  captured from its module, never copied; the fragment caller's written here), and the frozen digests of what our reader reads
 *  (`xml-trees.fixture.json`: `JsonDigest`, `expectPinned`).
 *  NOT here: the calls to our reader (`xml-shadow.test.support.ts`, the `xml/*.test.ts` files), the damage of the fuzzes
 *  (`xml-damage.test.support.ts`).
 *  Called by: the tests of `extractor/xml/`, `xml-reference-differential.property.test.ts`. */

import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { expect, vi } from "vitest";

import { byCodeUnit } from "doc-ir/internal";
import { PACKAGE_MAX_BYTES, PART_MAX_BYTES } from "../ooxml/package";
import { publicFixturePath, publicFixtures } from "../public-fixtures.test.support";
import { openZip } from "../zip/zip-package";
import type { FragmentParserOptions, PartParserOptions } from "./xml-parser";

export interface FixturePart {
  /** `<document path under fixtures/>!<part name>`: a technical id, never content. */
  readonly id: string;
  /** The document's path under `fixtures/`: a file name, never content. */
  readonly document: string;
  readonly text: string;
  /** The part's bytes, as the zip gave them: what our reader decodes itself (`decode.ts`). */
  readonly bytes?: Uint8Array;
}

const XML_PART = /\.(xml|rels|vml)$/i;
/** The package layer's own ceilings (`ooxml/package.ts`), one definition. */
const LIMITS = { partMaxBytes: PART_MAX_BYTES, packageMaxBytes: PACKAGE_MAX_BYTES };

export interface OfficeParts {
  /** Every Office document read, path under `fixtures/`. */
  readonly documents: readonly string[];
}

type Zip = ReturnType<typeof openZip>;

/** The archive at `path`, or `undefined` when the package layer cannot open it. */
const tryOpenZip = (path: string): Zip | undefined => {
  try {
    return openZip(new Uint8Array(readFileSync(path)), LIMITS);
  } catch {
    return undefined;
  }
};

/** The bytes of a part, or `undefined` when the package layer refuses to inflate it (it never reaches a parser). */
const tryReadBytes = async (zip: Zip, name: string): Promise<Uint8Array | undefined> => {
  try {
    return await zip.readBytes(name);
  } catch {
    return undefined;
  }
};

/** Hands every XML part of one archive to `handle`. */
const handleXmlParts = async (
  zip: Zip,
  document: string,
  utf8: InstanceType<typeof TextDecoder>,
  handle: (part: FixturePart) => void,
): Promise<void> => {
  for (const name of [...zip.names].sort(byCodeUnit)) {
    if (!XML_PART.test(name)) continue;
    const bytes = await tryReadBytes(zip, name);
    if (bytes) handle({ id: `${document}!${name}`, document, text: utf8.decode(bytes), bytes });
  }
};

/** Every XML part of every public Office document, each handed to `onPart` as it is read and NOT kept (a corpus does not
 *  fit in memory). Documents come in path order, the parts of one in name order. */
export const publicOfficeXmlParts = async (onPart: (part: FixturePart) => void): Promise<OfficeParts> => {
  const utf8 = new TextDecoder("utf-8", { ignoreBOM: true });
  const documents: string[] = [];
  for (const path of publicFixtures(".docx", ".pptx", ".xlsx").sort(byCodeUnit)) {
    const document = path.slice("fixtures/".length);
    documents.push(document);
    const zip = tryOpenZip(publicFixturePath(path));
    if (zip !== undefined) await handleXmlParts(zip, document, utf8, onPart);
  }
  return { documents };
};

let cache: Promise<FixturePart[]> | undefined;

/** The XML parts of the public fixtures, read once per test file (the first call's `keep` decides). With `keep`, a part
 *  is retained only when it returns true: the fixtures (about 171 MB of text) are streamed, never held whole. */
export const fixtureXmlParts = (keep?: (part: FixturePart) => boolean): Promise<FixturePart[]> =>
  (cache ??= (async () => {
    const kept: FixturePart[] = [];
    await publicOfficeXmlParts((part) => {
      if (keep === undefined || keep(part)) kept.push(part);
    });
    return kept;
  })());

export interface ProductParsers {
  /** The options `ooxml/xml.ts` passes to its parser, as constructed. */
  readonly docxReaderOptions: PartParserOptions;
  /** The options a caller of the fragment parser writes when it names no depth (`FRAGMENT_CALLER_OPTIONS`). */
  readonly fragmentCallerOptions: FragmentParserOptions;
}

/**
 * What a caller of `XmlFragmentParser` writes when it names no depth: the six fixed values, each the only one the
 * parser has, and no `maxNestedTags`. The caller lives outside this package (the application's reading of a published
 * page), so its options are WRITTEN here, not captured: a test of this package depends on nothing outside it. Nothing
 * can drift in silence for all that: the parser refuses any other value at construction, and the application holds
 * with its own test that it passes exactly these.
 */
const FRAGMENT_CALLER_OPTIONS: FragmentParserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
  parseTagValue: false,
  htmlEntities: true,
};

/**
 * The reader's own caller (`ooxml/xml.ts`) keeps its parser and its options module-private: its options are therefore
 * CAPTURED, by loading the module afresh while OUR parser classes (`xml-parser.ts`) are replaced by recorders that keep
 * the options they are given. A copy of the options would drift in silence; a capture cannot.
 */
const captureProductParsers = async (): Promise<ProductParsers> => {
  const actual = await vi.importActual<typeof import("./xml-parser")>("./xml-parser");
  const seen: unknown[] = [];
  /** The caller only builds its parser at load time; nothing parses during the capture. */
  class Recording {
    constructor(options: unknown) {
      seen.push(options);
    }
    parse(): never {
      throw new Error("the capture never parses");
    }
  }
  vi.resetModules();
  vi.doMock("./xml-parser", () => ({ ...actual, XmlParser: Recording, XmlFragmentParser: Recording }));
  try {
    const built = (): number => seen.length;
    await import("../ooxml/xml");
    if (built() !== 1) throw new Error(`ooxml/xml.ts built ${built()} parsers; the capture expects one`);
    return { docxReaderOptions: seen[0] as PartParserOptions, fragmentCallerOptions: FRAGMENT_CALLER_OPTIONS };
  } finally {
    vi.doUnmock("./xml-parser");
    vi.resetModules();
  }
};

let parsers: Promise<ProductParsers> | undefined;

/** The callers' options, the reader's own captured once per test file. */
export const productParsers = (): Promise<ProductParsers> => (parsers ??= captureProductParsers());

export type XNode = Record<string, unknown>;

export type XmlEvent =
  | readonly ["open", string, readonly (readonly [string, string])[]]
  | readonly ["close", string]
  | readonly ["text", string];

/**
 * Frozen references: what OUR reader builds from the real documents, pinned as sha256 digests in
 * `xml-trees.fixture.json` (sections of `id → digest`, sorted keys, one per line), so that a change of any tree turns
 * a test red and names the document. A digest is of the canonical JSON of the tree, streamed (the tree of a 50 M
 * character sheet never becomes one string).
 * The file is rewritten ONLY by a deliberate run: `XML_TREES_UPDATE=1` (the run then fails on purpose, see `expectPinned`).
 */
export const XML_TREES_FILE = join(dirname(fileURLToPath(import.meta.url)), "xml-trees.fixture.json");

/** A sha256 over labelled JSON values: `add(label, value)` hashes `label`, a newline, the value as `JSON.stringify` writes it. */
export class JsonDigest {
  private readonly hash = createHash("sha256");
  private buffer = "";

  private write(text: string): void {
    this.buffer += text;
    if (this.buffer.length > 1 << 16) {
      this.hash.update(this.buffer);
      this.buffer = "";
    }
  }

  private walk(value: unknown): void {
    if (Array.isArray(value)) {
      this.write("[");
      value.forEach((item, index) => {
        if (index > 0) this.write(",");
        this.walk(item);
      });
      this.write("]");
    } else if (value !== null && typeof value === "object") {
      this.write("{");
      Object.entries(value).forEach(([key, item], index) => {
        if (index > 0) this.write(",");
        this.write(`${JSON.stringify(key)}:`);
        this.walk(item);
      });
      this.write("}");
    } else this.write(JSON.stringify(value));
  }

  add(label: string, value: unknown): this {
    this.write(`${label}\n`);
    this.walk(value);
    this.write("\n");
    return this;
  }

  hex(): string {
    this.hash.update(this.buffer);
    this.buffer = "";
    return this.hash.digest("hex");
  }
}

type PinnedFile = Record<string, Record<string, string>>;

/** Code-unit order, never the locale's: the file is the same on every machine. */
const compareKeys = (a: string, b: string): number => {
  if (a === b) return 0;
  return a < b ? -1 : 1;
};

const sortedRecord = <T>(record: Readonly<Record<string, T>>): Record<string, T> =>
  Object.fromEntries(Object.entries(record).sort(([a], [b]) => compareKeys(a, b)));

const readPinned = (): PinnedFile =>
  existsSync(XML_TREES_FILE) ? (JSON.parse(readFileSync(XML_TREES_FILE, "utf8")) as PinnedFile) : {};

/** The one test file that owns the sections of the frozen file, and the command that regenerates them. */
const REGENERATE = "XML_TREES_UPDATE=1 npx vitest run src/extractor/xml/xml-parser.test.ts";

/** Compares `actual` (`id → digest`) with the frozen `section`; a difference names its ids: changed, read with no entry,
 *  entry with nothing read. With `XML_TREES_UPDATE=1` the section is rewritten and the test FAILS, so that a rewrite
 *  never passes silently: it must be run again without the flag. */
export const expectPinned = (section: string, actual: Readonly<Record<string, string>>): void => {
  if (process.env.XML_TREES_UPDATE === "1") {
    const file = readPinned();
    file[section] = sortedRecord(actual);
    writeFileSync(XML_TREES_FILE, `${JSON.stringify(sortedRecord(file), null, 2)}\n`);
    expect.fail(
      `section "${section}" of xml-trees.fixture.json was REWRITTEN (${Object.keys(actual).length} entries): read the diff of the file, then run this test again WITHOUT XML_TREES_UPDATE`,
    );
  }
  const pinned = readPinned()[section] ?? {};
  const ids = Object.keys(actual);
  expect(
    {
      changed: ids.filter((id) => id in pinned && pinned[id] !== actual[id]).sort(byCodeUnit),
      readWithNoEntry: ids.filter((id) => !(id in pinned)).sort(byCodeUnit),
      entryWithNothingRead: Object.keys(pinned)
        .filter((id) => !(id in actual))
        .sort(byCodeUnit),
    },
    `section "${section}" of xml-trees.fixture.json. Regenerating (${REGENERATE}) is legitimate ONLY when a fixture document was added, removed or replaced; never to make a reader change pass without looking at it`,
  ).toEqual({ changed: [], readWithNoEntry: [], entryWithNothingRead: [] });
};
