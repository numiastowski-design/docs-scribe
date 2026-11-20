import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { BitReader, buildHuffmanTable } from "./huffman";

/** T.81 annex K, table K.3: luminance DC differences (counts per length, values). */
const K3 = {
  counts: new Uint8Array([0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0]),
  values: new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]),
};

const spec = (counts: number[], values: number[]) => ({
  counts: new Uint8Array([...counts, ...new Array<number>(16 - counts.length).fill(0)]),
  values: new Uint8Array(values),
});

/** Bytes from a string of bits, padded with 1-bits (T.81 F.1.2.3). */
const bits = (s: string): Uint8Array => {
  const padded = s + "1".repeat((8 - (s.length % 8)) % 8);
  return new Uint8Array(padded.match(/.{8}/g)!.map((b) => parseInt(b, 2)));
};

const refusal = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(ReaderError);
    expect((e as ReaderError).code).toBe("EXTRACT_FAILED");
    return (e as ReaderError).message;
  }
  throw new Error("expected a refusal");
};

describe("buildHuffmanTable (T.81 C.2, F.2.2.3)", () => {
  it("decodes the codes of table K.3: 00, 010, 011, 100, 101, 110, 1110 ... 111111110", () => {
    const table = buildHuffmanTable(K3);
    const codes = [
      "00",
      "010",
      "011",
      "100",
      "101",
      "110",
      "1110",
      "11110",
      "111110",
      "1111110",
      "11111110",
      "111111110",
    ];
    const reader = new BitReader(bits(codes.join("")));
    expect(codes.map(() => reader.decode(table))).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]);
  });

  it("refuses an over-subscribed table (3 codes of length 1)", () => {
    expect(refusal(() => buildHuffmanTable(spec([3], [0, 1, 2])))).toBe("jpeg: bad huffman table");
  });

  it("refuses a table that uses the all-1 code word (T.81 C.2: reserved)", () => {
    expect(refusal(() => buildHuffmanTable(spec([2], [0, 1])))).toBe("jpeg: bad huffman table");
    expect(refusal(() => buildHuffmanTable(spec([1, 2], [0, 1, 2])))).toBe("jpeg: bad huffman table");
    expect(() => buildHuffmanTable(spec([1, 1], [0, 1]))).not.toThrow();
  });

  it("refuses an empty table", () => {
    expect(refusal(() => buildHuffmanTable(spec([], [])))).toBe("jpeg: bad huffman table");
  });

  it("refuses a bit sequence that matches no code within 16 bits", () => {
    const table = buildHuffmanTable(spec([1], [7]));
    expect(refusal(() => new BitReader(new Uint8Array([0xff, 0x00, 0xff, 0x00])).decode(table))).toBe(
      "jpeg: bad huffman code",
    );
  });
});

describe("BitReader", () => {
  it("reads bits most significant first and extends them (T.81 F.2.2.1, EXTEND)", () => {
    const reader = new BitReader(bits("101" + "011" + "1"));
    expect(reader.receiveExtend(3)).toBe(5);
    expect(reader.receiveExtend(3)).toBe(-4);
    expect(reader.receiveExtend(1)).toBe(1);
    expect(reader.receiveExtend(0)).toBe(0);
  });

  it("skips the stuffed 0x00 after a 0xFF data byte (T.81 B.1.1.5)", () => {
    const reader = new BitReader(new Uint8Array([0xff, 0x00, 0x80]));
    expect(reader.receive(8)).toBe(0xff);
    expect(reader.receive(1)).toBe(1);
  });

  it("refuses to read past the end of the segment: never an endless run of zeros", () => {
    const reader = new BitReader(new Uint8Array([0xaa]));
    expect(reader.receive(8)).toBe(0xaa);
    expect(refusal(() => reader.receive(1))).toBe("jpeg: truncated scan data");
  });

  it("refuses a marker met inside the data", () => {
    const reader = new BitReader(new Uint8Array([0xff, 0xd0, 0x00]));
    expect(refusal(() => reader.receive(1))).toBe("jpeg: unexpected marker in scan data");
  });

  it("restarts at the expected RSTn, after fill bytes, and refuses another one", () => {
    const ok = new BitReader(new Uint8Array([0x7f, 0xff, 0xff, 0xd3, 0x80]));
    expect(ok.receive(1)).toBe(0);
    ok.restart(3);
    expect(ok.receive(1)).toBe(1);
    const wrong = new BitReader(new Uint8Array([0x7f, 0xff, 0xd4, 0x80]));
    wrong.receive(1);
    expect(refusal(() => wrong.restart(3))).toBe("jpeg: bad restart marker");
    const missing = new BitReader(new Uint8Array([0x7f, 0x80]));
    missing.receive(1);
    expect(refusal(() => missing.restart(0))).toBe("jpeg: bad restart marker");
  });
});
