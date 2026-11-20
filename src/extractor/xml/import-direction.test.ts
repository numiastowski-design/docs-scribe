/**
 * One-way rule: no module of our XML reader (`extractor/xml/`) imports `ooxml/xml.ts`.
 *
 * WHY. `ooxml/xml.ts` builds its parser AT MODULE LOAD (`new XmlParser({...})`, from `xml/xml-parser.ts`). A reader
 * module importing `ooxml/xml.ts` back (`tree.ts` for `MAX_XML_DEPTH`, `namespaces.ts` for `renamed`) closes an import
 * cycle. Which module enters the cycle first then decides whether `ooxml/xml.ts` runs before `class XmlParser` has been
 * evaluated: ESM gives `ReferenceError: Cannot access 'XmlParser' before initialization`, the shipped esbuild CJS bundle
 * gives `TypeError: XmlParser is not a constructor`. Reproduced by the review of 26/10/2026 (Node 24.16 and esbuild, both
 * orders). The shipped bundle would survive by import ORDER only, while test files entering the other way crash at
 * import (19 of them when the cycle was measured, 26/10/2026). What the reader needs from the docx side lives in the
 * leaf `ooxml/prefixes.ts`; the last two tests pin that: it is the reader's only import INTO THIS PACKAGE that leaves
 * `xml/` (another package named by its package name is not followed), and it imports nothing at all, a stricter rule than the cycle needs.
 *
 * SCOPE. The reader's SOURCE modules, `xml/` and its sub-folders. Test files (`*.test.ts`, `*.test.support.ts`) are out:
 * some read a part through the product's caller (`ooxml/xml.ts`), and importing it is their point.
 * METHOD. The module specifiers of each file (static `import`/`export … from`, dynamic `import()`), resolved against
 * the file's directory, or, for a self-reference `docs-scribe/<path>`, through the `exports` map of this package: robust to
 * spelling (`../ooxml/xml`, `../ooxml/xml.ts`, `./../ooxml/xml`; a self-reference has no extension: `docs-scribe/extractor/ooxml/xml`).
 */

import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join, normalize, relative } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const XML_DIR = dirname(fileURLToPath(import.meta.url));
const EXTRACTOR_DIR = normalize(join(XML_DIR, ".."));
const PACKAGE_DIR = normalize(join(XML_DIR, "..", "..", ".."));
const FORBIDDEN = normalize(join(EXTRACTOR_DIR, "ooxml", "xml"));
const PREFIXES = normalize(join(EXTRACTOR_DIR, "ooxml", "prefixes"));

const isTestFile = (name: string): boolean => name.endsWith(".test.ts") || name.endsWith(".test.support.ts");

const readerModules = (): string[] =>
  readdirSync(XML_DIR, { recursive: true })
    .map(String)
    .filter((name) => name.endsWith(".ts") && !isTestFile(name));

/** Every module specifier a file names: `from "x"` (import, export, type-only) and `import("x")`. */
const SPECIFIER = /(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s+)(["'])([^"'\n]+)\1/g;

const specifiersOf = (source: string): string[] => [...source.matchAll(SPECIFIER)].map((match) => match[2]!);

const PACKAGE_NAME = "docs-scribe";
type ExportsMap = Readonly<Record<string, unknown>>;

const PACKAGE_EXPORTS = (JSON.parse(readFileSync(join(PACKAGE_DIR, "package.json"), "utf8")) as { exports: ExportsMap })
  .exports;

/** The file an `exports` entry names; a target that is not a string is not followed: this guard fails loudly rather than guess. */
const targetOf = (key: string, target: unknown): string => {
  if (typeof target !== "string")
    throw new Error(`exports entry "${key}" is not a string: this guard does not follow it`);
  return target;
};

/** What `docs-scribe/<subpath>` names in an `exports` map, as a file path, or `undefined` when no entry matches. Node's rules,
 *  as far as this package uses them: an exact key wins; else, among the `/*` keys that match, the LONGEST prefix wins, and
 *  EVERY `*` of its target takes the matched remainder. */
const selfReference = (specifier: string, exportsMap: ExportsMap): string | undefined => {
  const subpath = `.${specifier.slice(PACKAGE_NAME.length)}`;
  if (exportsMap[subpath] !== undefined) return normalize(join(PACKAGE_DIR, targetOf(subpath, exportsMap[subpath])));
  const matching = Object.keys(exportsMap)
    .filter((key) => key.endsWith("/*") && subpath.startsWith(key.slice(0, -1)))
    .sort((a, b) => b.length - a.length);
  const key = matching[0];
  if (key === undefined) return undefined;
  return normalize(join(PACKAGE_DIR, targetOf(key, exportsMap[key]).replaceAll("*", subpath.slice(key.length - 1))));
};

const withoutExtension = (path: string): string => path.replace(/\.(?:ts|js|mjs)$/, "");

/** The specifier as a path without extension, or `undefined` for any other package. */
const resolved = (fromFile: string, specifier: string): string | undefined => {
  if (specifier.startsWith(".")) return withoutExtension(normalize(join(dirname(fromFile), specifier)));
  if (specifier.startsWith(`${PACKAGE_NAME}/`)) {
    const target = selfReference(specifier, PACKAGE_EXPORTS);
    return target === undefined ? undefined : withoutExtension(target);
  }
  return undefined;
};

describe("xml reader — import direction", () => {
  it("lists the reader's source modules (a guard that reads nothing guards nothing)", () => {
    expect(readerModules()).toEqual(
      expect.arrayContaining(["tree.ts", "namespaces.ts", "tokenizer.ts", "decode.ts", "xml-parser.ts"]),
    );
  });

  // The forbidden module moved once (`docx-reader/xml.ts` became `ooxml/xml.ts`) and this guard kept aiming at the old
  // path: it stayed green while guarding nothing.
  it("the forbidden module exists where the guard aims", () => {
    expect(existsSync(`${FORBIDDEN}.ts`)).toBe(true);
  });

  // Since the sources live in the published package (19/11/2026), its `exports` lead to `dist/` alone: no self-reference
  // names a source any more, so the relative spelling is the only way in. The exports rules stay held on a literal map.
  it("a self-reference `docs-scribe/<path>` names no source through the published exports, and any other package is not resolved", () => {
    const from = join(XML_DIR, "tree.ts");
    expect(resolved(from, "docs-scribe/extractor/ooxml/xml")).toBeUndefined();
    expect(resolved(from, "docs-scribe/lib/xml")).toBeUndefined();
    expect(resolved(from, "../ooxml/xml")).toBe(FORBIDDEN);
    expect(resolved(from, "doc-ir/internal")).toBeUndefined();
  });

  it("the resolution of a self-reference follows the exports rules, on a literal map", () => {
    const at = (path: string): string => normalize(join(PACKAGE_DIR, path));
    const map: ExportsMap = {
      "./lib": "./src/lib/index.ts",
      "./lib/light": "./src/lib/light.ts",
      "./*": "./src/any/*.ts",
      "./extractor/*": "./src/extractor/*.ts",
      "./twice/*": "./src/*/*.ts",
    };
    const cases: readonly (readonly [string, string | undefined])[] = [
      ["docs-scribe/lib", at("src/lib/index.ts")], // an exact key
      ["docs-scribe/lib/light", at("src/lib/light.ts")], // an exact key beats the wildcards
      ["docs-scribe/other/a", at("src/any/other/a.ts")], // the only wildcard that matches
      ["docs-scribe/extractor/a", at("src/extractor/a.ts")], // the longest of two prefixes, the shorter listed first
      ["docs-scribe/twice/a", at("src/a/a.ts")], // every `*` of the target is replaced
    ];
    for (const [specifier, expected] of cases) expect(selfReference(specifier, map), specifier).toBe(expected);
    const reversed = Object.fromEntries(Object.entries(map).reverse());
    expect(
      selfReference("docs-scribe/extractor/a", reversed),
      "the longest prefix, whatever the order of the keys",
    ).toBe(at("src/extractor/a.ts"));
    expect(selfReference("docs-scribe/a", { "./x/*": "./src/*.ts" }), "no entry matches").toBeUndefined();
  });

  it("a conditional export is refused with an explicit message, not a TypeError", () => {
    expect(() => selfReference("docs-scribe/a", { "./a": { import: "./a.mjs", default: "./a.js" } })).toThrow(
      'exports entry "./a" is not a string',
    );
    expect(() => selfReference("docs-scribe/b/c", { "./b/*": { default: "./src/*.ts" } })).toThrow(
      'exports entry "./b/*" is not a string',
    );
  });

  it("no source module of extractor/xml/ imports ooxml/xml.ts", () => {
    const offenders = readerModules().flatMap((name) => {
      const file = join(XML_DIR, name);
      return specifiersOf(readFileSync(file, "utf8"))
        .filter((specifier) => resolved(file, specifier) === FORBIDDEN)
        .map((specifier) => `${relative(XML_DIR, file)} imports "${specifier}"`);
    });
    expect(offenders).toEqual([]);
  });

  it("the only import of a reader source module that leaves xml/ is the leaf ooxml/prefixes.ts", () => {
    const leaving = new Set(
      readerModules().flatMap((name) => {
        const file = join(XML_DIR, name);
        return specifiersOf(readFileSync(file, "utf8"))
          .map((specifier) => resolved(file, specifier))
          .filter((target): target is string => target !== undefined && relative(XML_DIR, target).startsWith(".."))
          .map((target) => relative(EXTRACTOR_DIR, target).replaceAll("\\", "/"));
      }),
    );
    expect([...leaving]).toEqual(["ooxml/prefixes"]);
  });

  it("the leaf ooxml/prefixes.ts imports nothing, so neither ooxml/xml.ts nor a module of xml/", () => {
    expect(existsSync(`${PREFIXES}.ts`)).toBe(true);
    expect(specifiersOf(readFileSync(`${PREFIXES}.ts`, "utf8"))).toEqual([]);
  });
});
