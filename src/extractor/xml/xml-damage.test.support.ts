/** Test support: the damage the fuzzes of our XML reader inflict on a real part, as `fast-check` arbitraries (`textDamage`,
 *  `damages`) and the functions that apply it to text (`damageText`) or to bytes (`damageBytes`, `damaged`,
 *  with `utf16Bytes` and the UTF-8 decoder `utf8`). A position is a fraction of the length, so one damage applies to a
 *  part of any size.
 *  NOT here: the seeds, the reading of a damaged part, the rules held on it (each fuzz writes its own), any other parser.
 *  Called by: `xml-differential.property.test.ts` (which also tests `damageBytes`) and
 *  `xml-reference-differential.property.test.ts`. */

import fc from "fast-check";

const HOSTILE_CHARS = [
  "<",
  ">",
  "&",
  '"',
  "'",
  "/",
  "=",
  ";",
  "#",
  "]",
  "!",
  "?",
  "-",
  ":",
  " ",
  "\r",
  "x",
  "\u0000",
  "\uFFFE",
  "\uD800",
  "\u0085",
  "\u2028",
  "\u1680",
];
const HOSTILE_SNIPPETS = [
  "&nbsp;",
  "&#0;",
  "&#x110000;",
  "&#X41;",
  "&#x9;",
  "&#13;",
  "&amp",
  "&lt;&gt;&quot;&apos;",
  "<!DOCTYPE a>",
  '<!ENTITY e "x">',
  "<![CDATA[x]]>",
  "<![CDATA[",
  "<![CDATA[]]>",
  "<!--x-->",
  "<!--a--b-->",
  "<!--<!DOCTYPE a>-->",
  "<?pi x?>",
  '<?pi a="?>"?>',
  "<?pi it's?>",
  '<?pi __proto__="x"?>',
  '<?xml version="1.0"?>',
  "]]>",
  "</w:p>",
  "<w:p>",
  "<w:p/>",
  ' xmlns:p=""',
  ' a="1"',
  ' a="1" a="2"',
  ' __proto__="x"',
  "<constructor/>",
  "<a:b:c/>",
  "<zz:q/>",
  ' w:val="1" w:val="2"',
  "<!--x",
  "<?pi",
  "\r\n",
  "\uFEFF",
  "\u1680",
];
/** UTF-8 that is not: a lone lead byte, a stray continuation, an overlong form, a surrogate, past U+10FFFF, NUL. */
const BAD_BYTES: readonly (readonly number[])[] = [
  [0xff],
  [0xc3],
  [0x80],
  [0xc0, 0xaf],
  [0xed, 0xa0, 0x80],
  [0xf4, 0x90, 0x80, 0x80],
  [0xe2, 0x82],
  [0x00],
];
const LABELS = [
  "UTF-8",
  "utf-8",
  "windows-1252",
  "ISO-8859-1",
  "us-ascii",
  "Shift_JIS",
  "UTF-16",
  "UTF-16LE",
  "x-user-defined",
  "no-such-encoding",
];

export type TextDamage =
  | { readonly kind: "truncate"; readonly at: number }
  | { readonly kind: "replace"; readonly at: number; readonly char: string }
  | { readonly kind: "delete"; readonly at: number; readonly length: number }
  | { readonly kind: "copy"; readonly from: number; readonly length: number; readonly to: number }
  | { readonly kind: "insert"; readonly at: number; readonly snippet: string }
  /** Inserted just after a `>`: in content, where a reference or markup is read as such. */
  | { readonly kind: "insertAfterTag"; readonly at: number; readonly snippet: string };

export type ByteDamage =
  | { readonly kind: "badBytes"; readonly at: number; readonly bytes: readonly number[] }
  | { readonly kind: "cutBytes"; readonly at: number }
  | { readonly kind: "toggleMark" }
  | { readonly kind: "declare"; readonly label: string }
  | { readonly kind: "utf16"; readonly bigEndian: boolean; readonly mark: boolean; readonly declare: boolean };

export type Damage = TextDamage | ByteDamage;

const position = fc.noBias(fc.integer({ min: 0, max: 999_999 })).map((n) => n / 1_000_000);

export const textDamage: fc.Arbitrary<TextDamage> = fc.oneof(
  position.map((at) => ({ kind: "truncate", at }) as const),
  fc.record({ at: position, char: fc.constantFrom(...HOSTILE_CHARS) }).map((m) => ({ kind: "replace", ...m }) as const),
  fc.record({ at: position, length: fc.integer({ min: 1, max: 64 }) }).map((m) => ({ kind: "delete", ...m }) as const),
  fc
    .record({ from: position, length: fc.integer({ min: 1, max: 200 }), to: position })
    .map((m) => ({ kind: "copy", ...m }) as const),
  fc
    .record({ at: position, snippet: fc.constantFrom(...HOSTILE_SNIPPETS) })
    .map((m) => ({ kind: "insert", ...m }) as const),
  {
    arbitrary: fc
      .record({ at: position, snippet: fc.constantFrom(...HOSTILE_SNIPPETS) })
      .map((m) => ({ kind: "insertAfterTag", ...m }) as const),
    weight: 2,
  },
);

const byteDamage: fc.Arbitrary<ByteDamage> = fc.oneof(
  fc.record({ at: position, bytes: fc.constantFrom(...BAD_BYTES) }).map((m) => ({ kind: "badBytes", ...m }) as const),
  position.map((at) => ({ kind: "cutBytes", at }) as const),
  fc.constant({ kind: "toggleMark" } as const),
  { arbitrary: fc.constantFrom(...LABELS).map((label) => ({ kind: "declare", label }) as const), weight: 2 },
  fc
    .record({ bigEndian: fc.boolean(), mark: fc.boolean(), declare: fc.boolean() })
    .map((m) => ({ kind: "utf16", ...m }) as const),
);

export const damages: fc.Arbitrary<Damage[]> = fc.array(
  fc.oneof({ arbitrary: textDamage, weight: 3 }, { arbitrary: byteDamage, weight: 1 }),
  { minLength: 1, maxLength: 3 },
);

const isByteDamage = (d: Damage): d is ByteDamage =>
  ["badBytes", "cutBytes", "toggleMark", "declare", "utf16"].includes(d.kind);

const at = (length: number, fraction: number): number => Math.min(length, Math.floor(fraction * length));

export const damageText = (text: string, d: TextDamage): string => {
  switch (d.kind) {
    case "truncate":
      return text.slice(0, at(text.length, d.at));
    case "replace":
      return text.slice(0, at(text.length, d.at)) + d.char + text.slice(at(text.length, d.at) + 1);
    case "delete":
      return text.slice(0, at(text.length, d.at)) + text.slice(at(text.length, d.at) + d.length);
    case "copy": {
      const from = at(text.length, d.from);
      const to = at(text.length, d.to);
      return text.slice(0, to) + text.slice(from, from + d.length) + text.slice(to);
    }
    case "insert":
      return text.slice(0, at(text.length, d.at)) + d.snippet + text.slice(at(text.length, d.at));
    case "insertAfterTag": {
      const closings: number[] = [];
      for (let k = text.indexOf(">"); k >= 0; k = text.indexOf(">", k + 1)) closings.push(k + 1);
      const where =
        closings.length === 0 ? 0 : closings[Math.min(closings.length - 1, Math.floor(d.at * closings.length))]!;
      return text.slice(0, where) + d.snippet + text.slice(where);
    }
  }
};

const DECLARATION = /^(\uFEFF?)(<\?xml[^>]*?>)?/;
const ENCODING_PSEUDO = /\sencoding\s*=\s*(["'])[^"']*\1/;

/** The text with its declaration naming `label` (a declaration added when there is none). */
const declaring = (text: string, label: string): string =>
  text.replace(DECLARATION, (_whole, mark: string, declaration: string | undefined) => {
    if (declaration === undefined) return `${mark}<?xml version="1.0" encoding="${label}"?>`;
    if (ENCODING_PSEUDO.test(declaration)) return mark + declaration.replace(ENCODING_PSEUDO, ` encoding="${label}"`);
    return (
      mark +
      declaration.replace(/^<\?xml\s+version\s*=\s*(["'])[^"']*\1/, (version) => `${version} encoding="${label}"`)
    );
  });

/** The parts end to end in one array, without spreading millions of bytes into a JS array. */
const concatBytes = (...parts: readonly Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
};

export const utf8 = new TextDecoder("utf-8", { ignoreBOM: true });

export const utf16Bytes = (text: string, bigEndian: boolean): Uint8Array => {
  const out = new Uint8Array(text.length * 2);
  for (let k = 0; k < text.length; k++) {
    const unit = text.charCodeAt(k);
    out[2 * k + (bigEndian ? 1 : 0)] = unit & 0xff;
    out[2 * k + (bigEndian ? 0 : 1)] = unit >> 8;
  }
  return out;
};

export const damageBytes = (bytes: Uint8Array, d: ByteDamage): Uint8Array => {
  switch (d.kind) {
    case "badBytes": {
      const where = at(bytes.length, d.at);
      return concatBytes(bytes.subarray(0, where), Uint8Array.from(d.bytes), bytes.subarray(where));
    }
    case "cutBytes":
      return bytes.subarray(0, at(bytes.length, d.at));
    case "toggleMark":
      return bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
        ? bytes.subarray(3)
        : concatBytes(Uint8Array.from([0xef, 0xbb, 0xbf]), bytes);
    case "declare": {
      // The declaration is ASCII: it is rewritten through a byte-for-byte text round trip, so that every byte the
      // declaration does not hold is kept as it is. A decoder that remapped 0x80-0x9F would break that; the test "the
      // declared-encoding damage rewrites the declaration only" (`xml-differential.property.test.ts`) pins it.
      const text = Buffer.from(bytes).toString("latin1");
      const marked = text.startsWith("\u00EF\u00BB\u00BF");
      const rewritten = declaring(marked ? text.slice(3) : text, d.label);
      return new Uint8Array(Buffer.from((marked ? "\u00EF\u00BB\u00BF" : "") + rewritten, "latin1"));
    }
    case "utf16": {
      let text = utf8.decode(bytes).replace(/^\uFEFF/, "");
      if (d.declare) text = declaring(text, "UTF-16");
      return utf16Bytes(d.mark ? `\uFEFF${text}` : text, d.bigEndian);
    }
  }
};

/** The damaged bytes: the text damages first, on the text, then the byte damages, on its UTF-8 bytes. */
export const damaged = (seed: string, ds: readonly Damage[]): Uint8Array => {
  const text = ds.filter((d): d is TextDamage => !isByteDamage(d)).reduce(damageText, seed);
  return ds.filter(isByteDamage).reduce(damageBytes, new TextEncoder().encode(text));
};
