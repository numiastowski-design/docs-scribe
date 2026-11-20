import { crc32, deflateSync } from "node:zlib";

import { describe, expect, it } from "vitest";

import { decodePng } from "./png-decode";

// PNGs built here byte by byte, rows filtered with the formulas of the PNG
// specification (ISO/IEC 15948 §9), independently of the decoder: our own
// encoder (`png.ts`) only writes filter 0, which would leave filters 1 to 4
// untested.

const u32 = (n: number): number[] => [(n >>> 24) & 0xff, (n >>> 16) & 0xff, (n >>> 8) & 0xff, n & 0xff];

const chunk = (type: string, data: Uint8Array): number[] => {
  const body = new Uint8Array([...Buffer.from(type, "latin1"), ...data]);
  return [...u32(data.length), ...body, ...u32(crc32(body))];
};

const pngOf = (
  width: number,
  height: number,
  ihdrTail: number[],
  raw: Uint8Array,
  extra: number[] = [],
  idat: Uint8Array = deflateSync(raw),
): Uint8Array =>
  new Uint8Array([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk("IHDR", new Uint8Array([...u32(width), ...u32(height), ...ihdrTail])),
    ...extra,
    ...chunk("IDAT", idat),
    ...chunk("IEND", new Uint8Array(0)),
  ]);

/** Row `y` filtered with `type`, `bpp` bytes per pixel, `prev` the unfiltered row above. */
const filterRow = (type: number, row: number[], prev: number[], bpp: number): number[] =>
  row.map((v, x) => {
    const a = x >= bpp ? row[x - bpp]! : 0;
    const b = prev[x] ?? 0;
    const c = x >= bpp ? (prev[x - bpp] ?? 0) : 0;
    const p = a + b - c;
    const pred =
      type === 0
        ? 0
        : type === 1
          ? a
          : type === 2
            ? b
            : type === 3
              ? (a + b) >> 1
              : Math.abs(p - a) <= Math.abs(p - b) && Math.abs(p - a) <= Math.abs(p - c)
                ? a
                : Math.abs(p - b) <= Math.abs(p - c)
                  ? b
                  : c;
    return (v - pred) & 0xff;
  });

describe("decodePng", () => {
  // 3×5 RGBA, one row per filter type 0…4, values that make every predictor differ.
  const rows = [0, 1, 2, 3, 4].map((y) =>
    [0, 1, 2].flatMap((x) => [(37 * y + 11 * x) & 0xff, (200 - 13 * x * y) & 0xff, 7 * x + y, 255 - 40 * y]),
  );

  it("unfilters the five filter types (8-bit RGBA)", () => {
    const raw = rows.flatMap((row, y) => [y, ...filterRow(y, row, rows[y - 1] ?? [], 4)]);
    const png = decodePng(pngOf(3, 5, [8, 6, 0, 0, 0], new Uint8Array(raw)), 100);
    // Failure value: filtered bytes left as they are (rows 1 to 4 differ).
    expect(png && [png.width, png.height, png.channels, [...png.pixels]]).toEqual([3, 5, 4, rows.flat()]);
  });

  it("joins an image whose data is split over several IDAT chunks", () => {
    const raw = new Uint8Array(rows.flatMap((row) => [0, ...row]));
    const z = deflateSync(raw);
    const cuts = [5, 17];
    const extra = [...chunk("IDAT", z.subarray(0, cuts[0])), ...chunk("IDAT", z.subarray(cuts[0], cuts[1]))];
    const png = decodePng(pngOf(3, 5, [8, 6, 0, 0, 0], raw, extra, z.subarray(cuts[1])), 100);
    expect(png && [...png.pixels]).toEqual(rows.flat());
  });

  it("refuses what it does not decode exactly: 16-bit, palette, interlaced, tRNS, too many pixels, a stream longer than announced", () => {
    const raw = new Uint8Array(rows.flatMap((row) => [0, ...row]));
    const ok = pngOf(3, 5, [8, 6, 0, 0, 0], raw);
    expect(decodePng(ok, 100)).toBeDefined();
    expect(decodePng(pngOf(3, 5, [16, 6, 0, 0, 0], raw), 100)).toBeUndefined();
    expect(decodePng(pngOf(3, 5, [8, 3, 0, 0, 0], raw), 100)).toBeUndefined();
    expect(decodePng(pngOf(3, 5, [8, 6, 0, 0, 1], raw), 100)).toBeUndefined();
    expect(decodePng(pngOf(3, 5, [8, 2, 0, 0, 0], raw, chunk("tRNS", new Uint8Array(6))), 100)).toBeUndefined();
    expect(decodePng(ok, 14)).toBeUndefined();
    // Announced 3×5, the stream inflates to 100 times more.
    expect(
      decodePng(pngOf(3, 5, [8, 6, 0, 0, 0], raw, [], deflateSync(new Uint8Array(raw.length * 100))), 100),
    ).toBeUndefined();
  });
});
