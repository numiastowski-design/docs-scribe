/**
 * The bytes of a part become text the way XML 1.0 §4.3.3 and Appendix F say, so that a VALID part encoded in UTF-16
 * (ECMA-376 Part 2 allows UTF-8 and UTF-16) is read. No versioned fixture and no corpus document holds a UTF-16 part:
 * the cases below transcode REAL fixture parts into UTF-16, both byte orders, with a BOM, and require the very events
 * our reader gives on the original UTF-8 part.
 */
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { decodeXmlPart } from "./decode";
import { tokenize } from "./tokenizer";
import { fixtureXmlParts, publicOfficeXmlParts } from "./xml-fixtures.test.support";
import { eventsOf, eventsOfBytes, readerAccepts } from "./xml-shadow.test.support";

const utf16le = (text: string): Uint8Array => new Uint8Array(Buffer.from(text, "utf16le"));
const utf16be = (text: string): Uint8Array => {
  const bytes = utf16le(text);
  for (let i = 0; i + 1 < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1]!, bytes[i]!];
  return bytes;
};
const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

const refusalOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(ReaderError);
    return (error as ReaderError).message;
  }
  throw new Error("expected a refusal");
};

const tokenizeBytes = (bytes: Uint8Array): void => {
  const { text, encoding } = decodeXmlPart(bytes);
  tokenize(text, { open() {}, close() {}, text() {} }, encoding);
};

describe("xml decode — the encoding of a part's bytes (XML 1.0 §4.3.3, Appendix F)", () => {
  it("detects UTF-16 by its byte-order mark, either order, and keeps the mark as U+FEFF for the tokenizer to skip", () => {
    expect(decodeXmlPart(utf16le("\uFEFF<a/>"))).toEqual({ text: "\uFEFF<a/>", encoding: "UTF-16LE" });
    expect(decodeXmlPart(utf16be("\uFEFF<a/>"))).toEqual({ text: "\uFEFF<a/>", encoding: "UTF-16BE" });
  });

  // XML 1.0 §4.3.3: « Entities encoded in UTF-16 MUST […] begin with the Byte Order Mark » (read at the source,
  // 25/10/2026, at integration). No
  // valid Office part can be one: refused as soon as Appendix F.1 recognises `<?` in either 16-bit order.
  it("refuses UTF-16 without a mark (`<?` in either order, Appendix F.1), whatever it declares", () => {
    for (const declared of ["UTF-16", "UTF-16LE", "UTF-16BE"]) {
      const text = `<?xml version="1.0" encoding="${declared}"?><a/>`;
      expect(
        refusalOf(() => decodeXmlPart(utf16le(text))),
        declared,
      ).toBe("xml: UTF-16 without a byte-order mark");
      expect(
        refusalOf(() => decodeXmlPart(utf16be(text))),
        declared,
      ).toBe("xml: UTF-16 without a byte-order mark");
    }
  });

  it("everything else is UTF-8", () => {
    expect(decodeXmlPart(utf8("\uFEFF<a>é</a>"))).toEqual({ text: "\uFEFF<a>é</a>", encoding: "UTF-8" });
    expect(decodeXmlPart(utf8("<a>é</a>"))).toEqual({ text: "<a>é</a>", encoding: "UTF-8" });
    expect(decodeXmlPart(new Uint8Array())).toEqual({ text: "", encoding: "UTF-8" });
  });

  it("refuses bytes that are not legal in the encoding (§4.3.3: a fatal error), never a replacement character", () => {
    expect(refusalOf(() => decodeXmlPart(new Uint8Array([0x3c, 0x61, 0x3e, 0xff, 0x3c, 0x2f, 0x61, 0x3e])))).toBe(
      "xml: bytes not legal in their encoding",
    );
    expect(refusalOf(() => decodeXmlPart(new Uint8Array([0xff, 0xfe, 0x3c, 0x00, 0x61])))).toBe(
      "xml: bytes not legal in their encoding",
    );
    expect(refusalOf(() => decodeXmlPart(new Uint8Array([0xff, 0xfe, 0x3c, 0x00, 0x00, 0xd8, 0x3e, 0x00])))).toBe(
      "xml: bytes not legal in their encoding",
    );
  });
});

describe("xml decode — the declaration must agree with the bytes (§4.3.3)", () => {
  it("reads a UTF-16 part with a mark, declaring UTF-16 (any case) or nothing", () => {
    for (const bytes of [
      utf16le(`\uFEFF<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\r\n<a x="é">t</a>`),
      utf16be(`\uFEFF<?xml version="1.0" encoding="utf-16"?><a x="é">t</a>`),
      utf16le(`\uFEFF<a x="é">t</a>`),
    ]) {
      expect(eventsOfBytes(bytes)).toEqual([
        ["open", "a", [["x", "é"]]],
        ["text", "t"],
        ["close", "a"],
      ]);
    }
  });

  // `UTF-16LE` and `UTF-16BE` are other charset names than `UTF-16` (RFC 2781 section 3.3: no mark is used with them), and
  // Beside the mark a UTF-16 part must have, a name that fixes the byte order is refused (§4.3.3: with these names the
  // mark is not a mark).
  it("refuses a declaration that names another encoding than the bytes', or a byte order beside the mark", () => {
    expect(refusalOf(() => tokenizeBytes(utf16le(`\uFEFF<?xml version="1.0" encoding="UTF-8"?><a/>`)))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => tokenizeBytes(utf16be(`\uFEFF<?xml version="1.0" encoding="UTF-16BE"?><a/>`)))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => tokenizeBytes(utf16le(`\uFEFF<?xml version="1.0" encoding="UTF-16LE"?><a/>`)))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => tokenizeBytes(utf8(`<?xml version="1.0" encoding="UTF-16"?><a/>`)))).toBe(
      "xml: declared encoding does not match the bytes",
    );
  });
});

// Option B (25/10/2026). ECMA-376 read at the source the same day (5th editions, ecma-international.org): the
// « UTF-8 or UTF-16 only » rule (Part 2 §6.2.5 a, 4th edition [M1.17]) covers the PACKAGE's own XML parts only (content
// types, relationships, core properties, signatures); Part 1 sets no encoding rule for the format's parts, and calls a
// Custom XML Data part « any XML ». XML 1.0 §4.3.3 lets a processor read other encodings. So a declared encoding we can
// decode is read AS DECLARED (no gibberish, no refusal of a valid file); one we cannot is refused, our own literal.
describe("xml decode — a declared encoding other than UTF-8 or UTF-16 is read as declared", () => {
  const bytes = (...pieces: (string | number[])[]): Uint8Array =>
    new Uint8Array(pieces.flatMap((piece) => (typeof piece === "string" ? [...Buffer.from(piece, "latin1")] : piece)));
  const declaring = (label: string): string => `<?xml version="1.0" encoding="${label}"?><a>`;

  it("Windows-1252, ISO-8859-1 and us-ascii: the bytes above 0x7F are the characters the declaration names", () => {
    expect(eventsOfBytes(bytes(declaring("windows-1252"), [0xe9, 0x80], "</a>"))).toEqual([
      ["open", "a", []],
      ["text", "é€"],
      ["close", "a"],
    ]);
    expect(eventsOfBytes(bytes(declaring("ISO-8859-1"), [0xe9], "</a>"))).toEqual([
      ["open", "a", []],
      ["text", "é"],
      ["close", "a"],
    ]);
    expect(eventsOfBytes(bytes(declaring("us-ascii"), "x</a>"))).toEqual([
      ["open", "a", []],
      ["text", "x"],
      ["close", "a"],
    ]);
  });

  // Found by an agent on 25/10/2026: Node 22.13 to 22.22 decode `windows-1252` as Latin-1, so 0x80-0x9F come out
  // as C1 controls (€ → U+0080), text lost in silence (nodejs/node#60888, fixed by #60893). Forge runs `nodejs24.x` since 26/10/2026,
  // minor version not ours: the 0x80-0x9F row is written here, character by character (WHATWG index-windows-1252;
  // 0x81, 0x8D, 0x8F, 0x90, 0x9D are unassigned and map to the C1 control of the same value), never left to the runtime.
  // On a runtime with the bug this test is red without our own table; on Node 24 it only pins the table.
  it("Windows-1252 0x80-0x9F: every byte is the character of the WHATWG index, whatever the runtime's decoder", () => {
    const row = Array.from({ length: 32 }, (_, i) => 0x80 + i);
    const expected = "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ";
    expect(eventsOfBytes(bytes(declaring("windows-1252"), row, "</a>"))).toEqual([
      ["open", "a", []],
      ["text", expected],
      ["close", "a"],
    ]);
  });

  // Review of 25/10/2026: a per-byte replace callback cost 36.6 s and 2.26 GB on 64 MB of the 0x80-0x9F
  // row (well-formed content). Bounded WORK: one table, one loop. 64 MB, the per-part ceiling: the old code is superlinear
  // (16 MB 1.1 s, 64 MB 9-11 s in a fresh process; counter-review), the table about 0.12 s.
  it("Windows-1252 is decoded in one table pass: 64 MB of the 0x80-0x9F row in well under the time a per-byte callback takes", () => {
    const body = new Uint8Array(64 * 1024 * 1024).map((_, i) => 0x80 + (i % 32));
    const part = bytes(declaring("windows-1252"), [...body.subarray(0, 0)], "");
    const whole = new Uint8Array(part.length + body.length);
    whole.set(part);
    whole.set(body, part.length);
    const started = performance.now();
    const { text } = decodeXmlPart(whole);
    expect(performance.now() - started).toBeLessThan(3_000);
    expect(text.length).toBe(whole.length);
    expect(text.at(-1)).toBe("Ÿ");
  }, 60_000);

  it("a multi-byte legacy encoding (Shift_JIS)", () => {
    expect(eventsOfBytes(bytes(declaring("Shift_JIS"), [0x93, 0xfa, 0x96, 0x7b], "</a>"))).toEqual([
      ["open", "a", []],
      ["text", "日本"],
      ["close", "a"],
    ]);
  });

  it("bytes illegal in the declared encoding are refused, never replaced", () => {
    expect(refusalOf(() => eventsOfBytes(bytes(declaring("Shift_JIS"), [0x81, 0x20], "</a>")))).toBe(
      "xml: bytes not legal in their encoding",
    );
  });

  it("an encoding we cannot decode is refused with our own literal: unknown, or decoded only into private-use characters", () => {
    for (const label of ["x-no-such-encoding", "iso-2022-kr", "x-user-defined"]) {
      expect(
        refusalOf(() => eventsOfBytes(bytes(declaring(label), "x</a>"))),
        label,
      ).toBe("xml: unknown encoding declared");
    }
  });

  // Counter-review of option B (25/10/2026): a declaration made long by white space (well-formed) was read past a
  // 256-byte window, decoded as UTF-8, then refused as a contradiction with the bytes. The declaration is now
  // read to its end; past our 4 KB ceiling the refusal is ours, its own literal.
  it("a long declaration is read to its end; past our ceiling, its own refusal", () => {
    const spaced = `<?xml version="1.0"${" ".repeat(260)}encoding="windows-1252"?><a>`;
    expect(eventsOfBytes(bytes(spaced, [0xe9], "</a>"))).toEqual([
      ["open", "a", []],
      ["text", "é"],
      ["close", "a"],
    ]);
    const tooLong = bytes(`<?xml version="1.0"${" ".repeat(5_000)}encoding="windows-1252"?><a>x</a>`);
    expect(refusalOf(() => eventsOfBytes(tooLong))).toBe("xml: XML declaration too long");
  });

  it("a declaration the bytes contradict: a UTF-8 mark naming another encoding, ASCII-compatible bytes naming UTF-16", () => {
    // A mark over Windows-1252 is weighed against the bytes (below): only bytes that are not legal UTF-8 contradict it.
    expect(refusalOf(() => eventsOfBytes(bytes([0xef, 0xbb, 0xbf], declaring("windows-1252"), [0xe9], "</a>")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    expect(refusalOf(() => eventsOfBytes(bytes([0xef, 0xbb, 0xbf], declaring("shift_jis"), "x</a>")))).toBe(
      "xml: declared encoding does not match the bytes",
    );
    for (const label of ["UTF-16", "UTF-16LE", "utf-16be"]) {
      expect(
        refusalOf(() => eventsOfBytes(bytes(declaring(label), "x</a>"))),
        label,
      ).toBe("xml: declared encoding does not match the bytes");
    }
  });

  // Decided by the product owner on 26/10/2026 (« préférer l'UTF-8 », after seeing what a refusal costs): when both readings
  // are legal, the bytes are read as UTF-8, not as declared. Read as declared, Windows-1252 named over UTF-8 accented
  // bytes would publish « Ã© » for « é ». A real
  // Windows-1252 text is almost never legal UTF-8 with an accent in it (a lead byte 0xC2-0xF4 followed by 0x80-0xBF), so
  // such a part is UTF-8 whose declaration lies. Bias, written: the rare Windows-1252 text that does spell such a pair comes
  // out as the UTF-8 it also is (« Ã© » typed on purpose publishes « é »).
  // Limit: only Windows-1252 (and its labels ISO-8859-1, us-ascii); a multi-byte encoding (Shift_JIS) reads legal UTF-8 by chance far more often.
  // The text is UTF-8 but the DECLARED name is returned, which the tokenizer checks against the declaration: returning "UTF-8"
  // made it refuse the part (review, 26/10/2026; the end-to-end case is in `docx-reader.test.ts`).
  it("a declared Windows-1252 over bytes that are legal UTF-8 with an accent is read as UTF-8, whatever name the declaration uses", () => {
    for (const label of ["windows-1252", "ISO-8859-1", "us-ascii"]) {
      const part = bytes(declaring(label), [...utf8("café €")], "</a>");
      const read = decodeXmlPart(part);
      expect(read.text, label).toContain("café €");
      expect(read.encoding, label).toBe("windows-1252");
      expect(eventsOfBytes(part), label).toEqual([
        ["open", "a", []],
        ["text", "café €"],
        ["close", "a"],
      ]);
    }
  });

  it("a UTF-8 mark over such bytes reads them as UTF-8 too: the mark is stronger evidence, not weaker", () => {
    const part = bytes([0xef, 0xbb, 0xbf], declaring("windows-1252"), [...utf8("café")], "</a>");
    expect(decodeXmlPart(part).text).toContain("café");
    expect(eventsOfBytes(part)).toEqual([
      ["open", "a", []],
      ["text", "café"],
      ["close", "a"],
    ]);
  });

  it("the same declaration over bytes that are NOT legal UTF-8, or over ASCII alone, is still read as declared", () => {
    expect(decodeXmlPart(bytes(declaring("windows-1252"), [0x63, 0x61, 0x66, 0xe9], "</a>")).text).toContain("café");
    expect(decodeXmlPart(bytes(declaring("windows-1252"), "cafe</a>")).text).toContain("cafe");
  });
});

describe("xml decode — real parts of the versioned fixtures, transcoded to UTF-16", () => {
  /** The declaration rewritten to name UTF-16 (or `name`), or one added. */
  const declaring = (text: string, name: string): string => {
    const body = text.startsWith("\uFEFF") ? text.slice(1) : text;
    const declaration = /^<\?xml[^?]*\?>/.exec(body);
    if (declaration === null) return `<?xml version="1.0" encoding="${name}"?>${body}`;
    const rewritten = /encoding=["'][^"']*["']/.test(declaration[0])
      ? declaration[0].replace(/encoding=["'][^"']*["']/, `encoding="${name}"`)
      : declaration[0].replace(/^<\?xml\s+version=(["'])[^"']*\1/, (v) => `${v} encoding="${name}"`);
    return rewritten + body.slice(declaration[0].length);
  };

  it("the same events as on the UTF-8 original, in both byte orders", async () => {
    // Small parts only (a transcoded copy of each is made twice); the parts our reader reads.
    const parts = await fixtureXmlParts((part) => part.text.length < 64 * 1024 && readerAccepts(part.text));
    // Floor re-measured on the public documents (25/11/2026): 2 982 parts.
    expect(parts.length).toBeGreaterThan(2_981);
    for (const part of parts) {
      const expected = eventsOf(part.text);
      const le = declaring(part.text, "UTF-16");
      expect(eventsOfBytes(utf16le(`\uFEFF${le}`)), part.id).toEqual(expected);
      expect(eventsOfBytes(utf16be(`\uFEFF${le}`)), part.id).toEqual(expected);
    }
  }, 300_000);

  it("on every versioned part, the bytes decode to the very text the product decodes (no part is UTF-16, none is refused)", async () => {
    const product = new TextDecoder("utf-8", { ignoreBOM: true });
    // Streamed: the whole fixtures never sit in memory.
    const differing: string[] = [];
    let count = 0;
    await publicOfficeXmlParts((part) => {
      count++;
      if (part.bytes === undefined || decodeXmlPart(part.bytes).text !== product.decode(part.bytes))
        differing.push(part.id);
    });
    // Floor re-measured on the public documents (25/11/2026): 3 116 parts.
    expect(count).toBeGreaterThan(3_115);
    expect(differing).toEqual([]);
  }, 300_000);

  it("a valid UTF-16 part is read", () => {
    const bytes = utf16le(`\uFEFF<?xml version="1.0" encoding="UTF-16"?><a>t</a>`);
    expect(eventsOfBytes(bytes)).toEqual([
      ["open", "a", []],
      ["text", "t"],
      ["close", "a"],
    ]);
  });
});
