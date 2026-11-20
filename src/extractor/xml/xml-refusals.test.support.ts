/** Test support: the reading of the XML reader's own SOURCES through the TypeScript parser, to say whether a refusal
 *  message is still a literal raised in code (for the closed lists of the fuzzes): `READER_SOURCE_NAMES` and
 *  `readerSourceText` (which sources, their text), `refusalReasonsIn` and `nonLiteralReasonCallsIn` (the reasons of the
 *  calls the walk recognises), `unrecognisedReferencesIn` (the references it cannot follow), `raisesRefusal`. Kept apart
 *  from `xml-shadow.test.support.ts` so that only the tests that use it load the compiler.
 *  NOT here: running the reader on a part (`xml-shadow.test.support.ts`), the parts of the fixtures
 *  (`xml-fixtures.test.support.ts`), our reader itself (`tokenizer.ts`, `tree.ts`).
 *  Called by: `tokenizer.test.ts`, the closed lists of `namespaces-differential.property.test.ts` and
 *  `xml-differential.property.test.ts`, and `xml-reference-differential.property.test.ts`. */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import ts from "typescript";

/** The reader sources that raise refusals (`xml-parser.ts` only compares one message, it raises none). */
export const READER_SOURCE_NAMES = ["tokenizer.ts", "decode.ts", "tree.ts", "namespaces.ts"] as const;

/** The text of one reader source, from this folder. */
export const readerSourceText = (name: string): string =>
  readFileSync(join(dirname(fileURLToPath(import.meta.url)), name), "utf8");

type ReasonSite = { reason: string | undefined; call: string };

const REASON_ARGUMENT_INDEX: ReadonlyMap<string, number> = new Map([
  ["xmlRefusal", 0],
  ["readUntil", 1],
]);

/** The name a call is written with: `f(…)` and `object.f(…)` give `f`. */
const calleeName = (call: ts.CallExpression): string => {
  const { expression } = call;
  if (ts.isIdentifier(expression)) return expression.text;
  return ts.isPropertyAccessExpression(expression) ? expression.name.text : "";
};

/** The argument of a call that holds the refusal reason, when the call is one of the two shapes. */
const reasonArgument = (call: ts.CallExpression): ts.Expression | undefined => {
  const index = REASON_ARGUMENT_INDEX.get(calleeName(call));
  return index === undefined ? undefined : call.arguments[index];
};

/** The reason argument of each call to `xmlRefusal(reason)` and `readUntil(terminator, reason)` in one source, read
 *  through the TypeScript parser: a comment is no node, and a call written in a string is no call. `reason` is the
 *  literal's value when the argument is a string literal (any quote, or a template without substitution), else
 *  undefined. */
const reasonSitesIn = (source: string, fileName: string): ReasonSite[] => {
  const file = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true);
  const sites: ReasonSite[] = [];
  const visit = (node: ts.Node): void => {
    const argument = ts.isCallExpression(node) ? reasonArgument(node) : undefined;
    if (ts.isCallExpression(node) && argument !== undefined)
      sites.push({ reason: ts.isStringLiteralLike(argument) ? argument.text : undefined, call: node.getText(file) });
    ts.forEachChild(node, visit);
  };
  visit(file);
  return sites;
};

/** The reasons that one reader source raises as literals, in code only (see `reasonSitesIn`). */
export const refusalReasonsIn = (source: string, fileName = "source.ts"): Set<string> =>
  new Set(reasonSitesIn(source, fileName).flatMap((site) => (site.reason === undefined ? [] : [site.reason])));

/** The calls of one source whose reason is not a literal: the helper cannot read them, so a test lists them. */
export const nonLiteralReasonCallsIn = (source: string, fileName = "source.ts"): string[] =>
  reasonSitesIn(source, fileName).flatMap((site) => (site.reason === undefined ? [site.call] : []));

/** Whether the identifier is the name its own declaration gives (a variable, a function, a method). */
const isOwnDeclaration = (node: ts.Identifier): boolean => {
  const { parent } = node;
  return (
    (ts.isVariableDeclaration(parent) || ts.isFunctionDeclaration(parent) || ts.isMethodDeclaration(parent)) &&
    parent.name === node
  );
};

/** Whether an import or export specifier keeps the name (`{ xmlRefusal }`), as opposed to renaming it either way. */
const isSameNameSpecifier = (node: ts.Identifier): boolean =>
  (ts.isImportSpecifier(node.parent) || ts.isExportSpecifier(node.parent)) && node.parent.propertyName === undefined;

/** Whether the identifier is the callee of a call that `calleeName` reads: `f(…)` or `object.f(…)`. */
const isRecognisedCallee = (node: ts.Identifier): boolean => {
  const { parent } = node;
  if (ts.isCallExpression(parent)) return parent.expression === node;
  return (
    ts.isPropertyAccessExpression(parent) &&
    parent.name === node &&
    ts.isCallExpression(parent.parent) &&
    parent.parent.expression === parent
  );
};

/** The references to `xmlRefusal` and `readUntil` in one source that the walk of `reasonSitesIn` cannot follow: every
 *  use of either name that is not its own declaration, a same-name import or export, or the direct callee of a
 *  recognised call. A parenthesised or cast callee, `f.call(…)`, an alias (`as`, `const r = f`) and `obj["f"](…)` are
 *  reported, each as the text of the node around the name, so that a raise written that way cannot go unseen. */
export const unrecognisedReferencesIn = (source: string, fileName = "source.ts"): string[] => {
  const file = ts.createSourceFile(fileName, source, ts.ScriptTarget.Latest, true);
  const found: string[] = [];
  const visit = (node: ts.Node): void => {
    if (ts.isIdentifier(node) && REASON_ARGUMENT_INDEX.has(node.text)) {
      if (!isOwnDeclaration(node) && !isSameNameSpecifier(node) && !isRecognisedCallee(node))
        found.push(node.parent.getText(file));
    } else if (
      ts.isStringLiteralLike(node) &&
      REASON_ARGUMENT_INDEX.has(node.text) &&
      ts.isElementAccessExpression(node.parent) &&
      node.parent.argumentExpression === node
    ) {
      found.push(node.parent.getText(file));
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return found;
};

let reasonsRaised: Set<string> | undefined;

/** Whether `message` (`xml: …`) is still a literal our reader raises (see `refusalReasonsIn`). A closed list of refusals
 *  checks its entries with it, so that an entry the reader no longer raises cannot linger. */
export const raisesRefusal = (message: string): boolean => {
  reasonsRaised ??= new Set(READER_SOURCE_NAMES.flatMap((name) => [...refusalReasonsIn(readerSourceText(name), name)]));
  return message.startsWith("xml: ") && reasonsRaised.has(message.slice("xml: ".length));
};
