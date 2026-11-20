/** A part's bytes are decoded ONCE by `parseXml` (review of 26/10/2026): the DTD guard reads the text `decodeXmlPart` gives,
 *  and the reader receives that text with its encoding, so nothing is decoded a second time (up to 2 x 64 MB of transient
 *  text at the part ceiling). `decodeXmlPart` is the only door from bytes to text, so counting its calls counts the decodes. */

import { beforeEach, describe, expect, it, vi } from "vitest";

const decodeCalls = vi.hoisted(() => ({ count: 0 }));

vi.mock("../xml/decode", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../xml/decode")>();
  return {
    ...actual,
    decodeXmlPart: (bytes: Uint8Array) => {
      decodeCalls.count++;
      return actual.decodeXmlPart(bytes);
    },
  };
});

const { parseXml } = await import("./xml");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const part = (declaration: string, text: string): string =>
  `${declaration}<w:document ${W}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
const textOfFirstRun = (root: ReturnType<typeof parseXml>): unknown =>
  JSON.stringify(root).match(/"#text":"([^"]*)"/)?.[1];

describe("parseXml decodes a part's bytes once", () => {
  beforeEach(() => {
    decodeCalls.count = 0;
  });

  it.each([
    ["UTF-8", new TextEncoder().encode(part("", "café"))],
    [
      "UTF-8 declared Windows-1252 over accented UTF-8",
      new TextEncoder().encode(part('<?xml version="1.0" encoding="windows-1252"?>', "café")),
    ],
    [
      "Windows-1252 as declared",
      new Uint8Array(Buffer.from(part('<?xml version="1.0" encoding="windows-1252"?>', "café"), "latin1")),
    ],
    ["UTF-16LE with its mark", new Uint8Array(Buffer.from(`\uFEFF${part("", "café")}`, "utf16le"))],
  ])("%s: one call, and the text is the part's own", (_name, bytes) => {
    expect(textOfFirstRun(parseXml(bytes))).toBe("café");
    expect(decodeCalls.count).toBe(1);
  });

  it("a DTD is still refused on the one decode", () => {
    expect(() => parseXml(new TextEncoder().encode(part("<!DOCTYPE w:document>", "x")))).toThrow("a DTD declaration");
    expect(decodeCalls.count).toBe(1);
  });

  // `DECLARES_DTD` tests the decoded text, comments included: refused, ugly and visible, never a document that expands in silence.
  it("a `<!DOCTYPE` inside a comment is refused too, with the exact message", () => {
    const refusal = /^docx-reader: a DTD declaration \(forbidden by OPC\)$/;
    expect(() => parseXml(new TextEncoder().encode(part("<!--<!DOCTYPE w:document>-->", "x")))).toThrow(refusal);
    expect(() => parseXml(new TextEncoder().encode(part("", "x<!-- <!doctype a> -->y")))).toThrow(refusal);
    expect(() => parseXml(new TextEncoder().encode(part("<!-- DOCTYPE a -->", "x")))).not.toThrow();
  });
});
