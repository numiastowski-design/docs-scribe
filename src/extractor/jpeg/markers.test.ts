import { decode as decodeJpegJs } from "jpeg-js";
import { describe, expect, it } from "vitest";

import { ReaderError } from "doc-ir";
import { jpegHeader } from "doc-ir/internal";
import { fixtureJpegs } from "./jpeg-fixtures.test.support";
import { parseJpeg, type JpegStructure } from "./markers";

const GUARD = { maxPixels: 12_000_000 };

// --- Synthetic streams -------------------------------------------------------

const u16 = (n: number): number[] => [(n >> 8) & 0xff, n & 0xff];
const seg = (marker: number, payload: number[]): number[] => [0xff, marker, ...u16(payload.length + 2), ...payload];
const SOI = [0xff, 0xd8];
const EOI = [0xff, 0xd9];
const dqt = (tq = 0, pq = 0): number[] => seg(0xdb, [(pq << 4) | tq, ...new Array<number>(pq ? 128 : 64).fill(1)]);
/** A DHT with one code of length 1 (value 0): structurally the smallest table. */
const dht = (tc: number, th: number, counts: number[] = [1], values: number[] = [0]): number[] => {
  const c = [...counts, ...new Array<number>(16 - counts.length).fill(0)];
  return seg(0xc4, [(tc << 4) | th, ...c, ...values]);
};
const sof = (
  marker: number,
  w: number,
  h: number,
  comps: [id: number, hv: number, tq: number][],
  precision = 8,
): number[] => seg(marker, [precision, ...u16(h), ...u16(w), comps.length, ...comps.flat()]);
const sos = (comps: [id: number, tables: number][], ss = 0, se = 63, ahal = 0): number[] =>
  seg(0xda, [comps.length, ...comps.flat(), ss, se, ahal]);
const ADOBE = (transform: number): number[] =>
  seg(0xee, [0x41, 0x64, 0x6f, 0x62, 0x65, 0x00, 0x64, 0, 0, 0, 0, transform]);
const JFIF = seg(0xe0, [0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0]);
const GREY: [number, number, number][] = [[1, 0x11, 0]];
const YCC: [number, number, number][] = [
  [1, 0x22, 0],
  [2, 0x11, 0],
  [3, 0x11, 0],
];
const ENTROPY = [0x12, 0x34];

const jpeg = (...parts: number[][]): Uint8Array => new Uint8Array(parts.flat());
const greyBaseline = (w = 8, h = 8, extra: number[][] = []): Uint8Array =>
  jpeg(SOI, dqt(), sof(0xc0, w, h, GREY), dht(0, 0), dht(1, 0), ...extra, sos([[1, 0x00]]), ENTROPY, EOI);

const refusal = (bytes: Uint8Array, guard = GUARD): ReaderError => {
  try {
    parseJpeg(bytes, guard);
  } catch (e) {
    expect(e).toBeInstanceOf(ReaderError);
    expect((e as ReaderError).code).toBe("EXTRACT_FAILED");
    expect((e as ReaderError).message).toMatch(/^jpeg: [a-z0-9 ,()/-]+$/);
    return e as ReaderError;
  }
  throw new Error("expected a refusal");
};

// --- Fixtures ----------------------------------------------------------------

describe("parseJpeg on every JPEG of the fixtures", () => {
  it("finds the 31 distinct JPEGs, and reads the same size as jpeg-js and as jpegHeader", async () => {
    const all = await fixtureJpegs();
    // Independent reading (Python survey of the same documents, 05/10/2026; the 73rd, from fixtures/docx-reel/11-dgemc-programme-terminale.docx,
    // read by Pillow 12.3 and a separate marker count on 06/10: progressive, 4:4:4, Adobe transform 1, 7 scans, 208×91).
    // Count re-measured on the public documents (25/11/2026).
    expect(all.length).toBe(31);
    for (const { id, bytes } of all) {
      const s = parseJpeg(bytes, GUARD);
      const ref = decodeJpegJs(bytes, {
        useTArray: true,
        formatAsRGBA: true,
        tolerantDecoding: false,
        maxResolutionInMP: 12,
        maxMemoryUsageInMB: 256,
      });
      expect([s.frame.width, s.frame.height], id).toEqual([ref.width, ref.height]);
      const header = jpegHeader(bytes)!;
      expect([s.frame.width, s.frame.height], id).toEqual([header.width, header.height]);
    }
  }, 120_000);

  it("reads the processes, samplings, Adobe transforms and scan counts of the independent survey", async () => {
    const all = await fixtureJpegs();
    const kinds = new Map<string, number>();
    const scans = new Map<number, number>();
    for (const { bytes } of all) {
      const s: JpegStructure = parseJpeg(bytes, GUARD);
      const key = `${s.frame.process} ${s.frame.components.map((c) => `${c.h}x${c.v}`).join(",")} adobe=${s.adobe?.transform ?? "-"}`;
      kinds.set(key, (kinds.get(key) ?? 0) + 1);
      scans.set(s.scans.length, (scans.get(s.scans.length) ?? 0) + 1);
      for (const scan of s.scans) expect(scan.data.length).toBeGreaterThan(0);
    }
    // Both tallies re-measured on the public documents (25/11/2026).
    expect(Object.fromEntries(kinds)).toEqual({
      "baseline 1x1,1x1,1x1 adobe=-": 3,
      "baseline 2x2,1x1,1x1 adobe=-": 23,
      "baseline 2x1,1x1,1x1 adobe=-": 1,
      "progressive 2x2,1x1,1x1 adobe=-": 3,
      "progressive 1x1,1x1,1x1 adobe=1": 1,
    });
    expect(Object.fromEntries(scans)).toEqual({ 1: 27, 10: 1, 8: 2, 7: 1 });
  }, 120_000);
});

// --- Structure ---------------------------------------------------------------

describe("parseJpeg on synthetic streams", () => {
  it("reads a minimal baseline grey stream: frame, MCU grid, tables, one scan and its entropy bytes", () => {
    const s = parseJpeg(greyBaseline(17, 9, [JFIF]), GUARD);
    expect(s.frame).toMatchObject({
      process: "baseline",
      width: 17,
      height: 9,
      maxH: 1,
      maxV: 1,
      mcusPerLine: 3,
      mcusPerColumn: 2,
    });
    expect(s.frame.components).toEqual([{ id: 1, h: 1, v: 1, tq: 0 }]);
    expect(s.jfif).toBe(true);
    expect(s.adobe).toBeUndefined();
    expect(s.quantTables[0]).toEqual(new Uint16Array(64).fill(1));
    expect(s.scans).toHaveLength(1);
    expect(s.scans[0]!.components).toEqual([{ componentIndex: 0, td: 0, ta: 0 }]);
    expect(s.scans[0]!.dcTables[0]!.counts[0]).toBe(1);
    expect([...s.scans[0]!.data]).toEqual(ENTROPY);
    expect(s.scans[0]!.restartInterval).toBe(0);
    expect(s.scans[0]!.dcTables).toHaveLength(4);
    expect(s.scans[0]!.acTables).toHaveLength(4);
  });

  it("reads the MCU grid of a 4:2:0 frame (T.81 A.2.4: 16 x 16 MCUs)", () => {
    const s = parseJpeg(
      jpeg(
        SOI,
        dqt(),
        sof(0xc0, 33, 17, YCC),
        dht(0, 0),
        dht(1, 0),
        sos([
          [1, 0],
          [2, 0],
          [3, 0],
        ]),
        ENTROPY,
        EOI,
      ),
      GUARD,
    );
    expect(s.frame).toMatchObject({ maxH: 2, maxV: 2, mcusPerLine: 3, mcusPerColumn: 2 });
  });

  it("keeps stuffed bytes, restart markers and nothing of the fill bytes in a scan's entropy segment (T.81 B.1.1.5)", () => {
    const data = [0x01, 0xff, 0x00, 0x02, 0xff, 0xd0, 0x03, 0xff, 0xd7, 0x04];
    const s = parseJpeg(
      jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), data, [0xff, 0xff, 0xff], EOI),
      GUARD,
    );
    expect([...s.scans[0]!.data]).toEqual(data);
  });

  it("snapshots the restart interval and the Huffman tables in effect at each scan", () => {
    const bytes = jpeg(
      SOI,
      dqt(),
      sof(0xc2, 8, 8, GREY),
      dht(0, 0),
      seg(0xdd, u16(4)),
      sos([[1, 0]], 0, 0),
      ENTROPY,
      dht(1, 0, [0, 1], [5]),
      seg(0xdd, u16(0)),
      sos([[1, 0]], 1, 63),
      ENTROPY,
      EOI,
    );
    const s = parseJpeg(bytes, GUARD);
    expect(s.frame.process).toBe("progressive");
    expect(s.scans.map((x) => x.restartInterval)).toEqual([4, 0]);
    expect(s.scans[0]!.acTables[0]).toBeUndefined();
    expect([...s.scans[1]!.acTables[0]!.values]).toEqual([5]);
    expect([s.scans[1]!.ss, s.scans[1]!.se, s.scans[1]!.ah, s.scans[1]!.al]).toEqual([1, 63, 0, 0]);
  });

  it("reads 16-bit quantisation tables and the extended process (SOF1)", () => {
    const s = parseJpeg(
      jpeg(SOI, dqt(2, 1), sof(0xc1, 8, 8, [[1, 0x11, 2]]), dht(0, 2), dht(1, 3), sos([[1, 0x23]]), ENTROPY, EOI),
      GUARD,
    );
    expect(s.frame.process).toBe("extended");
    expect(s.quantTables[2]).toEqual(new Uint16Array(64).fill(0x0101));
  });

  it("reads the Adobe APP14 transform, and accepts 4 components only with it (as jpeg-js)", () => {
    const cmyk: [number, number, number][] = [
      [1, 0x11, 0],
      [2, 0x11, 0],
      [3, 0x11, 0],
      [4, 0x11, 0],
    ];
    const body = (extra: number[][]): Uint8Array =>
      jpeg(
        SOI,
        ...extra,
        dqt(),
        sof(0xc0, 8, 8, cmyk),
        dht(0, 0),
        dht(1, 0),
        sos([
          [1, 0],
          [2, 0],
          [3, 0],
          [4, 0],
        ]),
        ENTROPY,
        EOI,
      );
    expect(parseJpeg(body([ADOBE(2)]), GUARD).adobe).toEqual({ transform: 2 });
    expect(refusal(body([])).message).toBe("jpeg: four components without an adobe segment");
  });

  it("skips APPn, COM and DNL segments, and ignores an APP14 too short to carry a transform", () => {
    const s = parseJpeg(
      greyBaseline(8, 8, [seg(0xe1, [1, 2, 3]), seg(0xfe, [0x41]), seg(0xee, [0x41, 0x64, 0x6f, 0x62, 0x65])]),
      GUARD,
    );
    expect(s.adobe).toBeUndefined();
    const dnl = jpeg(
      SOI,
      dqt(),
      sof(0xc0, 8, 8, GREY),
      dht(0, 0),
      dht(1, 0),
      sos([[1, 0]]),
      ENTROPY,
      seg(0xdc, u16(8)),
      EOI,
    );
    expect(parseJpeg(dnl, GUARD).scans).toHaveLength(1);
  });
});

// --- Guards ------------------------------------------------------------------

describe("parseJpeg refuses, with our own message", () => {
  // Review of 06/10/2026: each kept scan costs some 670 bytes the memory budget never sees, and the cap was
  // checked only after the whole parse (500 000 empty scans, 5 MB, grew the heap by 335 MB). Refused while reading.
  // Counter-review of 06/10/2026: a scan with Ss > Se counted NEGATIVELY in the work sum of `decode.ts`, so 125 of
  // them hid 131 heavy scans (4.3 s on 12 MP). Progressive parameters are now checked here, before any sum.
  it("🔴 a progressive scan with Ss > Se, or breaking the Ah/Al rules, while reading", () => {
    const progressive = (...scans: number[][]): Uint8Array =>
      jpeg(
        SOI,
        dqt(),
        sof(0xc2, 8, 8, GREY),
        dht(0, 0),
        dht(1, 0),
        sos([[1, 0x00]], 0, 0, 0x00),
        ENTROPY,
        ...scans,
        EOI,
      );
    expect(refusal(progressive([...sos([[1, 0x00]], 63, 0, 0x00), ...ENTROPY])).message).toBe(
      "jpeg: bad progressive scan parameters",
    );
    expect(refusal(progressive([...sos([[1, 0x00]], 0, 5, 0x00), ...ENTROPY])).message).toBe(
      "jpeg: bad progressive scan parameters",
    );
    expect(refusal(progressive([...sos([[1, 0x00]], 1, 63, 0x20), ...ENTROPY])).message).toBe(
      "jpeg: bad progressive scan parameters",
    );
    expect(
      parseJpeg(
        progressive([...sos([[1, 0x00]], 1, 63, 0x01), ...ENTROPY], [...sos([[1, 0x00]], 1, 63, 0x10), ...ENTROPY]),
        GUARD,
      ).scans,
    ).toHaveLength(3);
  });

  it("🔴 a 257th scan, while reading: no unbounded list of scans outside the memory budget", () => {
    const scans = Array.from({ length: 257 }, () => [...sos([[1, 0x00]]), ...ENTROPY]);
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), ...scans, EOI)).message).toBe(
      "jpeg: too many scans",
    );
    const fine = Array.from({ length: 256 }, () => [...sos([[1, 0x00]]), ...ENTROPY]);
    expect(
      parseJpeg(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), ...fine, EOI), GUARD).scans,
    ).toHaveLength(256);
  });

  it("a stream that does not start with SOI, or a second SOI", () => {
    expect(refusal(new Uint8Array([0x89, 0x50, 0x4e, 0x47])).message).toBe("jpeg: not a jpeg stream");
    expect(refusal(jpeg(SOI, SOI, EOI)).message).toBe("jpeg: unexpected marker");
  });

  it("a segment whose declared length runs past the end, or is under 2", () => {
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x10, 0x00, 1, 2, 3])).message).toBe("jpeg: truncated segment");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x01], EOI)).message).toBe("jpeg: truncated segment");
  });

  it("a stream that ends without EOI, without a frame or without a scan", () => {
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY)).message).toBe(
      "jpeg: truncated scan",
    );
    expect(refusal(jpeg(SOI, dqt())).message).toBe("jpeg: truncated stream");
    expect(refusal(jpeg(SOI, dqt(), EOI)).message).toBe("jpeg: no frame");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), EOI)).message).toBe("jpeg: no scan");
  });

  it("bytes between two segments that are not a marker", () => {
    expect(refusal(jpeg(SOI, [0x00], dqt(), EOI)).message).toBe("jpeg: marker expected");
  });

  it("an unsupported process: lossless, hierarchical, arithmetic, or an unknown marker", () => {
    for (const marker of [0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]) {
      expect(refusal(jpeg(SOI, dqt(), sof(marker, 8, 8, GREY), EOI)).message).toBe("jpeg: unsupported process");
    }
    expect(refusal(jpeg(SOI, seg(0xcc, [0, 0]), EOI)).message).toBe("jpeg: unsupported process");
    expect(refusal(jpeg(SOI, seg(0xde, [0, 0]), EOI)).message).toBe("jpeg: unsupported process");
    expect(refusal(jpeg(SOI, seg(0x02, [0, 0]), EOI)).message).toBe("jpeg: unexpected marker");
  });

  it("two frames, or a scan before the frame", () => {
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), sof(0xc0, 8, 8, GREY), EOI)).message).toBe(
      "jpeg: more than one frame",
    );
    expect(refusal(jpeg(SOI, dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY, EOI)).message).toBe(
      "jpeg: scan before frame",
    );
  });

  it("a frame over the pixel guard BEFORE anything is allocated, and accepts it exactly at the guard", () => {
    expect(parseJpeg(greyBaseline(4000, 3000), GUARD).frame.width).toBe(4000);
    expect(refusal(greyBaseline(4000, 3001)).message).toBe("jpeg: image over the pixel guard");
    expect(refusal(greyBaseline(65535, 65535)).message).toBe("jpeg: image over the pixel guard");
  });

  it("a zero width, a zero height (DNL-deferred: not supported), a precision other than 8", () => {
    expect(refusal(greyBaseline(0, 8)).message).toBe("jpeg: zero width");
    expect(refusal(greyBaseline(8, 0)).message).toBe("jpeg: height deferred to dnl not supported");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc1, 8, 8, GREY, 12), EOI)).message).toBe("jpeg: unsupported precision");
  });

  it("a frame header whose length disagrees with its component count", () => {
    const bad = seg(0xc0, [8, 0, 8, 0, 8, 3, 1, 0x11, 0]);
    expect(refusal(jpeg(SOI, dqt(), bad, EOI)).message).toBe("jpeg: bad frame header");
  });

  it("sampling factors outside 1..4 (T.81 B.2.2), and a duplicated component id", () => {
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, [[1, 0x01, 0]]), EOI)).message).toBe("jpeg: bad sampling factor");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, [[1, 0x51, 0]]), EOI)).message).toBe("jpeg: bad sampling factor");
    expect(
      refusal(
        jpeg(
          SOI,
          dqt(),
          sof(0xc0, 8, 8, [
            [1, 0x11, 0],
            [1, 0x11, 0],
            [3, 0x11, 0],
          ]),
          EOI,
        ),
      ).message,
    ).toBe("jpeg: duplicate component");
  });

  it("2 or 5 components (jpeg-js publishes neither), and a quantisation table id over 3", () => {
    expect(
      refusal(
        jpeg(
          SOI,
          dqt(),
          sof(0xc0, 8, 8, [
            [1, 0x11, 0],
            [2, 0x11, 0],
          ]),
          EOI,
        ),
      ).message,
    ).toBe("jpeg: unsupported component count");
    const five: [number, number, number][] = [1, 2, 3, 4, 5].map((id) => [id, 0x11, 0]);
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, five), EOI)).message).toBe("jpeg: unsupported component count");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, [[1, 0x11, 4]]), EOI)).message).toBe("jpeg: bad table id");
  });

  it("more than 10 blocks in an interleaved MCU (T.81 B.2.3), and accepts exactly 10", () => {
    const ten: [number, number, number][] = [
      [1, 0x42, 0],
      [2, 0x11, 0],
      [3, 0x11, 0],
    ];
    const eleven: [number, number, number][] = [
      [1, 0x33, 0],
      [2, 0x11, 0],
      [3, 0x11, 0],
    ];
    const scan = sos([
      [1, 0],
      [2, 0],
      [3, 0],
    ]);
    expect(
      parseJpeg(jpeg(SOI, dqt(), sof(0xc0, 8, 8, ten), dht(0, 0), dht(1, 0), scan, ENTROPY, EOI), GUARD).scans,
    ).toHaveLength(1);
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, eleven), dht(0, 0), dht(1, 0), scan, ENTROPY, EOI)).message).toBe(
      "jpeg: too many blocks per mcu",
    );
  });

  it("a scan naming an unknown or repeated component, or a Huffman table never defined", () => {
    const head = [SOI, dqt(), sof(0xc0, 8, 8, YCC), dht(0, 0), dht(1, 0)];
    expect(refusal(jpeg(...head, sos([[9, 0]]), ENTROPY, EOI)).message).toBe("jpeg: scan names an unknown component");
    expect(
      refusal(
        jpeg(
          ...head,
          sos([
            [1, 0],
            [1, 0],
          ]),
          ENTROPY,
          EOI,
        ),
      ).message,
    ).toBe("jpeg: scan names an unknown component");
    expect(refusal(jpeg(...head, sos([[1, 0x10]]), ENTROPY, EOI)).message).toBe("jpeg: missing huffman table");
    expect(refusal(jpeg(...head, sos([[1, 0x01]]), ENTROPY, EOI)).message).toBe("jpeg: missing huffman table");
  });

  it("a scan header whose length disagrees with its component count, or with 0 or 5 components", () => {
    const head = [SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0)];
    expect(refusal(jpeg(...head, seg(0xda, [2, 1, 0, 0, 63, 0]), ENTROPY, EOI)).message).toBe("jpeg: bad scan header");
    expect(refusal(jpeg(...head, seg(0xda, [0, 0, 63, 0]), ENTROPY, EOI)).message).toBe("jpeg: bad scan header");
  });

  it("a sequential scan whose spectral selection or approximation is not 0, 63, 0, 0 (T.81 B.2.3)", () => {
    const head = [SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0)];
    expect(refusal(jpeg(...head, sos([[1, 0]], 0, 0), ENTROPY, EOI)).message).toBe(
      "jpeg: bad sequential scan parameters",
    );
    expect(refusal(jpeg(...head, sos([[1, 0]], 0, 63, 0x01), ENTROPY, EOI)).message).toBe(
      "jpeg: bad sequential scan parameters",
    );
  });

  it("a frame component whose quantisation table is not defined when its first scan starts", () => {
    expect(refusal(jpeg(SOI, sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY, EOI)).message).toBe(
      "jpeg: missing quantization table",
    );
  });

  it("a quantisation table in use redefined with other values after a scan began (two readings disagree)", () => {
    const head = [SOI, dqt(), sof(0xc2, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]], 0, 0), ENTROPY];
    const other = seg(0xdb, [0, ...new Array<number>(64).fill(2)]);
    expect(refusal(jpeg(...head, other, sos([[1, 0]], 1, 63), ENTROPY, EOI)).message).toBe(
      "jpeg: quantization table changed after use",
    );
    // The same values again change nothing.
    expect(parseJpeg(jpeg(...head, dqt(), sos([[1, 0]], 1, 63), ENTROPY, EOI), GUARD).scans).toHaveLength(2);
  });

  it("bad DQT and DHT segments: precision 2, class 2, table id 4, too many values, counts past the segment", () => {
    expect(refusal(jpeg(SOI, seg(0xdb, [0x20, ...new Array<number>(64).fill(1)]), EOI)).message).toBe(
      "jpeg: bad quantization table",
    );
    expect(refusal(jpeg(SOI, seg(0xdb, [0x04, ...new Array<number>(64).fill(1)]), EOI)).message).toBe(
      "jpeg: bad table id",
    );
    expect(refusal(jpeg(SOI, seg(0xdb, [0x00, 1, 2, 3]), EOI)).message).toBe("jpeg: bad quantization table");
    expect(refusal(jpeg(SOI, seg(0xdb, [0x00, ...new Array<number>(64).fill(0)]), EOI)).message).toBe(
      "jpeg: bad quantization table",
    );
    expect(refusal(jpeg(SOI, dht(2, 0), EOI)).message).toBe("jpeg: bad huffman table");
    expect(refusal(jpeg(SOI, dht(0, 4), EOI)).message).toBe("jpeg: bad table id");
    expect(
      refusal(
        jpeg(SOI, dht(0, 0, [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200, 100], new Array<number>(300).fill(0)), EOI),
      ).message,
    ).toBe("jpeg: bad huffman table");
    expect(refusal(jpeg(SOI, seg(0xc4, [0x00, 4, ...new Array<number>(15).fill(0), 1, 2]), EOI)).message).toBe(
      "jpeg: bad huffman table",
    );
  });

  it("a baseline frame that uses Huffman table 2 or 3 (T.81 B.2.4.2: baseline has 0 and 1 only)", () => {
    expect(
      refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 2), dht(1, 2), sos([[1, 0x22]]), ENTROPY, EOI)).message,
    ).toBe("jpeg: bad table id");
  });

  it("a DRI segment of the wrong length", () => {
    expect(refusal(greyBaseline(8, 8, [seg(0xdd, [0, 1, 2])])).message).toBe("jpeg: bad restart interval");
  });

  it("never echoes the stream: a comment's text does not reach the message", () => {
    const secret = [...new TextEncoder().encode("confidential-client-text")];
    const e = refusal(jpeg(SOI, seg(0xfe, secret), [0xff, 0xe1, 0x7f, 0x00, ...secret]));
    expect(e.message).not.toContain("confidential");
  });

  it("a caller's invalid guard is a bug, not the stream's: plain Error", () => {
    expect(() => parseJpeg(greyBaseline(), { maxPixels: Number.NaN })).toThrow(/^parseJpeg: /);
    expect(() => parseJpeg(greyBaseline(), { maxPixels: -1 })).toThrow(/^parseJpeg: /);
  });
});

// --- Characterisation of the segment reader (tranche 10C4: parseJpeg cut into one reader per marker) -------

describe("parseJpeg: what each marker reader keeps, and the order in which the checks fire", () => {
  const JFIF_TAG = [0x4a, 0x46, 0x49, 0x46, 0];
  const ADOBE_TAG = [0x41, 0x64, 0x6f, 0x62, 0x65];
  const ADOBE_PAYLOAD = (transform: number): number[] => [...ADOBE_TAG, 0x00, 0x64, 0, 0, 0, 0, transform];

  it("JFIF: only an APP0 whose first five bytes are 'JFIF\\0' counts, five bytes being enough", () => {
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xe0, JFIF_TAG)]), GUARD).jfif).toBe(true);
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xe0, JFIF_TAG.slice(0, 4))]), GUARD).jfif).toBe(false);
    for (let k = 0; k < 5; k++) {
      const wrong = JFIF_TAG.map((b, i) => (i === k ? b ^ 1 : b));
      expect(parseJpeg(greyBaseline(8, 8, [seg(0xe0, wrong)]), GUARD).jfif, `byte ${k}`).toBe(false);
    }
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xe1, JFIF_TAG)]), GUARD).jfif).toBe(false);
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xfe, JFIF_TAG)]), GUARD).jfif).toBe(false);
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xef, JFIF_TAG)]), GUARD).jfif).toBe(false);
  });

  it("Adobe: only an APP14 with its tag and twelve bytes counts; the last one wins; one after the scan still counts", () => {
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xee, ADOBE_PAYLOAD(0))]), GUARD).adobe).toEqual({ transform: 0 });
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xee, ADOBE_PAYLOAD(0))]), GUARD).jfif).toBe(false);
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xee, ADOBE_PAYLOAD(2).slice(0, 11))]), GUARD).adobe).toBeUndefined();
    for (let k = 0; k < 5; k++) {
      const wrong = ADOBE_PAYLOAD(1).map((b, i) => (i === k ? b ^ 1 : b));
      expect(parseJpeg(greyBaseline(8, 8, [seg(0xee, wrong)]), GUARD).adobe, `byte ${k}`).toBeUndefined();
    }
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xed, ADOBE_PAYLOAD(1))]), GUARD).adobe).toBeUndefined();
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xe0, ADOBE_PAYLOAD(1))]), GUARD).adobe).toBeUndefined();
    expect(
      parseJpeg(greyBaseline(8, 8, [seg(0xee, ADOBE_PAYLOAD(1)), seg(0xee, ADOBE_PAYLOAD(3))]), GUARD).adobe,
    ).toEqual({ transform: 3 });
    const late = jpeg(
      SOI,
      dqt(),
      sof(0xc0, 8, 8, GREY),
      dht(0, 0),
      dht(1, 0),
      sos([[1, 0]]),
      ENTROPY,
      seg(0xee, ADOBE_PAYLOAD(1)),
      EOI,
    );
    expect(parseJpeg(late, GUARD).adobe).toEqual({ transform: 1 });
  });

  it("a result without Adobe has the keys frame, quantTables, scans, jfif; with it, adobe comes last", () => {
    expect(Object.keys(parseJpeg(greyBaseline(), GUARD))).toEqual(["frame", "quantTables", "scans", "jfif"]);
    expect(Object.keys(parseJpeg(greyBaseline(8, 8, [seg(0xee, ADOBE_PAYLOAD(1))]), GUARD))).toEqual([
      "frame",
      "quantTables",
      "scans",
      "jfif",
      "adobe",
    ]);
    expect(Object.keys(parseJpeg(greyBaseline(), GUARD).scans[0]!)).toEqual([
      "components",
      "ss",
      "se",
      "ah",
      "al",
      "restartInterval",
      "dcTables",
      "acTables",
      "data",
    ]);
  });

  it("DQT: several tables in one segment, 8 and 16 bits, each value at its own index", () => {
    const eight = Array.from({ length: 64 }, (_, k) => k + 1);
    const sixteen = Array.from({ length: 64 }, (_, k) => 300 + 257 * k);
    const payload = [0x00, ...eight, 0x13, ...sixteen.flatMap(u16)];
    const s = parseJpeg(
      jpeg(SOI, seg(0xdb, payload), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY, EOI),
      GUARD,
    );
    expect([...s.quantTables[0]!]).toEqual(eight);
    expect([...s.quantTables[3]!]).toEqual(sixteen);
    expect(s.quantTables[1]).toBeUndefined();
    expect(s.quantTables[2]).toBeUndefined();
    expect(s.quantTables).toHaveLength(4);
  });

  it("DQT: a second table cut short, or a zero at the last index (8 or 16 bits), refuses the whole segment", () => {
    const ok = [0x00, ...new Array<number>(64).fill(1)];
    const refused = (payload: number[]): string => refusal(jpeg(SOI, seg(0xdb, payload), EOI)).message;
    expect(refused([...ok, 0x11, ...new Array<number>(127).fill(1)])).toBe("jpeg: bad quantization table");
    expect(refused([...ok, 0x11, ...new Array<number>(128).fill(1)])).toBe("jpeg: no frame");
    expect(refused([0x00, ...new Array<number>(63).fill(1), 0])).toBe("jpeg: bad quantization table");
    expect(refused([0x10, ...new Array<number>(126).fill(1), 0, 0])).toBe("jpeg: bad quantization table");
    expect(refused([0x10, ...new Array<number>(126).fill(1), 1, 0])).toBe("jpeg: no frame");
    expect(refused([...ok, 0x04, ...new Array<number>(64).fill(1)])).toBe("jpeg: bad table id");
    expect(refused([...ok, 0x00])).toBe("jpeg: bad quantization table");
    expect(refused([])).toBe("jpeg: no frame");
  });

  it("DQT and DHT: a precision or class of 2, and a table id of 8 or more (not folded onto 0..7), are refused", () => {
    expect(refusal(jpeg(SOI, seg(0xdb, [0x20, ...new Array<number>(128).fill(1)]), EOI)).message).toBe(
      "jpeg: bad quantization table",
    );
    for (const id of [0x08, 0x0c, 0x0f]) {
      expect(refusal(jpeg(SOI, seg(0xdb, [id, ...new Array<number>(64).fill(1)]), EOI)).message).toBe(
        "jpeg: bad table id",
      );
      expect(refusal(jpeg(SOI, seg(0xc4, [id, 1, ...new Array<number>(15).fill(0), 0]), EOI)).message).toBe(
        "jpeg: bad table id",
      );
    }
  });

  it("DQT: a table is redefined freely before it is used, and later only when it is not one a started scan uses", () => {
    const two = (v: number): number[] => seg(0xdb, [0x00, ...new Array<number>(64).fill(v)]);
    const before = jpeg(SOI, two(1), two(2), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY, EOI);
    expect(parseJpeg(before, GUARD).quantTables[0]).toEqual(new Uint16Array(64).fill(2));
    // Table 1 is defined, never named by a scan: redefining it after the scan is allowed.
    const unused = jpeg(
      SOI,
      dqt(0),
      dqt(1),
      sof(0xc0, 8, 8, GREY),
      dht(0, 0),
      dht(1, 0),
      sos([[1, 0]]),
      ENTROPY,
      seg(0xdb, [0x01, ...new Array<number>(64).fill(9)]),
      EOI,
    );
    expect(parseJpeg(unused, GUARD).quantTables[1]).toEqual(new Uint16Array(64).fill(9));
    // Used by the second component of a later scan: refused.
    const ycc: [number, number, number][] = [
      [1, 0x11, 0],
      [2, 0x11, 1],
      [3, 0x11, 0],
    ];
    const head = [SOI, dqt(0), dqt(1), sof(0xc2, 8, 8, ycc), dht(0, 0), dht(1, 0), sos([[1, 0]], 0, 0), ENTROPY];
    const changed = seg(0xdb, [0x01, ...new Array<number>(64).fill(9)]);
    expect(parseJpeg(jpeg(...head, changed, EOI), { maxPixels: 100 }).quantTables[1]).toEqual(
      new Uint16Array(64).fill(9),
    );
    expect(
      refusal(jpeg(...head, sos([[2, 0]], 0, 0), ENTROPY, changed, sos([[2, 0]], 1, 63), ENTROPY, EOI)).message,
    ).toBe("jpeg: quantization table changed after use");
  });

  it("DQT: a table in use that differs in one value only, at the last index or in the high byte, is a change", () => {
    const head = [SOI, dqt(0), sof(0xc2, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]], 0, 0), ENTROPY];
    const last = [0x00, ...new Array<number>(63).fill(1), 2];
    expect(refusal(jpeg(...head, seg(0xdb, last), EOI)).message).toBe("jpeg: quantization table changed after use");
    const wide = [0x10, ...new Array<number>(64).fill(1).flatMap(u16)];
    wide[1] = 1;
    expect(refusal(jpeg(...head, seg(0xdb, wide), EOI)).message).toBe("jpeg: quantization table changed after use");
    expect(
      parseJpeg(jpeg(...head, seg(0xdb, [0x10, ...new Array<number>(64).fill(1).flatMap(u16)]), EOI), GUARD)
        .quantTables[0],
    ).toEqual(new Uint16Array(64).fill(1));
  });

  it("DHT: several tables in one segment, DC and AC apart, each at its id, copied (not viewed) from the segment", () => {
    const one = (tc: number, th: number, value: number): number[] => [
      (tc << 4) | th,
      1,
      ...new Array<number>(15).fill(0),
      value,
    ];
    const bytes = jpeg(
      SOI,
      dqt(),
      sof(0xc2, 8, 8, GREY),
      seg(0xc4, [...one(0, 1, 7), ...one(1, 3, 9), ...one(0, 0, 4)]),
      seg(0xc4, one(0, 1, 8)),
      sos([[1, 0x10]], 0, 0),
      ENTROPY,
      EOI,
    );
    const s = parseJpeg(bytes, GUARD);
    expect([...s.scans[0]!.dcTables[1]!.values]).toEqual([8]);
    expect([...s.scans[0]!.dcTables[0]!.values]).toEqual([4]);
    expect([...s.scans[0]!.acTables[3]!.values]).toEqual([9]);
    expect(s.scans[0]!.dcTables[2]).toBeUndefined();
    expect(s.scans[0]!.acTables[0]).toBeUndefined();
    expect(s.scans[0]!.dcTables).toHaveLength(4);
    expect(s.scans[0]!.acTables).toHaveLength(4);
    expect(s.scans[0]!.dcTables[0]!.counts.buffer).not.toBe(bytes.buffer);
    expect(s.scans[0]!.dcTables[0]!.values.buffer).not.toBe(bytes.buffer);
    expect(s.scans[0]!.dcTables[0]!.counts).toHaveLength(16);
  });

  it("DHT: 256 values are accepted, 257 are not, and a value missing at the end of the segment is a refusal", () => {
    const table = (total: number, present: number): Uint8Array =>
      jpeg(
        SOI,
        seg(0xc4, [
          0x00,
          total >> 8 ? 255 : total,
          ...new Array<number>(14).fill(0),
          total > 255 ? total - 255 : 0,
          ...new Array<number>(present).fill(3),
        ]),
        EOI,
      );
    expect(refusal(table(257, 257)).message).toBe("jpeg: bad huffman table");
    expect(refusal(table(256, 256)).message).toBe("jpeg: no frame");
    expect(refusal(table(256, 255)).message).toBe("jpeg: bad huffman table");
    expect(refusal(table(5, 5)).message).toBe("jpeg: no frame");
    expect(refusal(table(5, 4)).message).toBe("jpeg: bad huffman table");
    expect(refusal(jpeg(SOI, seg(0xc4, [0x00, ...new Array<number>(15).fill(0)]), EOI)).message).toBe(
      "jpeg: bad huffman table",
    );
    expect(refusal(jpeg(SOI, seg(0xc4, [0x00, ...new Array<number>(16).fill(0)]), EOI)).message).toBe("jpeg: no frame");
    expect(refusal(jpeg(SOI, seg(0xc4, [0x00, ...new Array<number>(16).fill(0), 0x10]), EOI)).message).toBe(
      "jpeg: bad huffman table",
    );
  });

  it("DHT: a table cut short after a good one refuses the segment, and a class above 1 or an id above 3 is named so", () => {
    const good = [0x00, 1, ...new Array<number>(15).fill(0), 0];
    expect(refusal(jpeg(SOI, seg(0xc4, [...good, 0x01, 1]), EOI)).message).toBe("jpeg: bad huffman table");
    expect(refusal(jpeg(SOI, seg(0xc4, [...good, 0x20, ...new Array<number>(16).fill(0)]), EOI)).message).toBe(
      "jpeg: bad huffman table",
    );
    expect(refusal(jpeg(SOI, seg(0xc4, [...good, 0x04, ...new Array<number>(16).fill(0)]), EOI)).message).toBe(
      "jpeg: bad table id",
    );
    expect(refusal(jpeg(SOI, seg(0xc4, [...good, 0x13, ...new Array<number>(16).fill(0)]), EOI)).message).toBe(
      "jpeg: no frame",
    );
  });

  it("a scan keeps the tables and the restart interval in effect when it started, and a later DHT or DRI does not reach it", () => {
    const bytes = jpeg(
      SOI,
      dqt(),
      sof(0xc2, 8, 8, GREY),
      dht(0, 0, [1], [6]),
      seg(0xdd, u16(0x1234)),
      sos([[1, 0]], 0, 0),
      ENTROPY,
      dht(0, 0, [1], [7]),
      seg(0xdd, u16(5)),
      sos([[1, 0]], 0, 0, 0x01),
      ENTROPY,
      seg(0xdd, u16(9)),
      EOI,
    );
    const s = parseJpeg(bytes, GUARD);
    expect(s.scans.map((x) => x.restartInterval)).toEqual([0x1234, 5]);
    expect([...s.scans[0]!.dcTables[0]!.values]).toEqual([6]);
    expect([...s.scans[1]!.dcTables[0]!.values]).toEqual([7]);
    expect(s.scans[0]!.dcTables).not.toBe(s.scans[1]!.dcTables);
  });

  it("DRI: two bytes exactly; none, one or three are refused, and an empty segment too", () => {
    for (const payload of [[], [1], [0, 1, 2]])
      expect(refusal(greyBaseline(8, 8, [seg(0xdd, payload)])).message).toBe("jpeg: bad restart interval");
    expect(parseJpeg(greyBaseline(8, 8, [seg(0xdd, [0xff, 0xff])]), GUARD).scans[0]!.restartInterval).toBe(65535);
  });

  it("fill bytes before a marker are skipped; a stream ending in fill bytes, or right after a marker, is refused as truncated", () => {
    const filled = jpeg(
      SOI,
      [0xff, 0xff, 0xff],
      dqt(),
      [0xff],
      sof(0xc0, 8, 8, GREY),
      dht(0, 0),
      dht(1, 0),
      sos([[1, 0]]),
      ENTROPY,
      [0xff, 0xff],
      EOI,
    );
    expect(parseJpeg(filled, GUARD).scans).toHaveLength(1);
    expect(refusal(jpeg(SOI, [0xff, 0xff])).message).toBe("jpeg: truncated stream");
    expect(refusal(jpeg(SOI, [0xff, 0xe1])).message).toBe("jpeg: truncated segment");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00])).message).toBe("jpeg: truncated segment");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x02])).message).toBe("jpeg: truncated stream");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x02], [0xff, 0xd9])).message).toBe("jpeg: no frame");
    expect(refusal(jpeg(SOI)).message).toBe("jpeg: truncated stream");
    expect(refusal(new Uint8Array([0xff])).message).toBe("jpeg: not a jpeg stream");
    expect(refusal(new Uint8Array([0xff, 0xd8])).message).toBe("jpeg: truncated stream");
  });

  it("a segment may end exactly at the end of the stream, and its length counts itself", () => {
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x05, 1, 2, 3])).message).toBe("jpeg: truncated stream");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x05, 1, 2])).message).toBe("jpeg: truncated segment");
    expect(refusal(jpeg(SOI, [0xff, 0xe1, 0x00, 0x02, 1], EOI)).message).toBe("jpeg: marker expected");
  });

  it("the marker kinds: APP0..APP15, COM, DNL known; RST, TEM, JPG and the reserved ones are not", () => {
    for (const m of [0xe0, 0xe5, 0xef, 0xfe, 0xdc])
      expect(parseJpeg(greyBaseline(8, 8, [seg(m, [1, 2])]), GUARD).scans).toHaveLength(1);
    for (const m of [0xd0, 0xd7, 0x01, 0x02, 0xbf, 0xf0, 0xfd, 0xd8]) {
      expect(refusal(jpeg(SOI, seg(m, [0, 0]), EOI)).message, m.toString(16)).toBe("jpeg: unexpected marker");
    }
    for (const m of [0xc8, 0xdf, 0xde])
      expect(refusal(jpeg(SOI, seg(m, [0, 0]), EOI)).message, m.toString(16)).toBe("jpeg: unsupported process");
  });

  it("the order of the checks: marker kind, then segment length, then the reader of that marker", () => {
    expect(refusal(jpeg(SOI, [0xff, 0xc3, 0x00])).message).toBe("jpeg: unsupported process");
    expect(refusal(jpeg(SOI, [0xff, 0x02, 0x00])).message).toBe("jpeg: unexpected marker");
    expect(refusal(jpeg(SOI, [0xff, 0xdc, 0xff, 0xff, 1])).message).toBe("jpeg: truncated segment");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), sof(0xc0, 0, 0, GREY), EOI)).message).toBe(
      "jpeg: more than one frame",
    );
    expect(refusal(jpeg(SOI, seg(0xda, []), EOI)).message).toBe("jpeg: scan before frame");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), seg(0xda, []), EOI)).message).toBe("jpeg: bad scan header");
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), ENTROPY)).message).toBe(
      "jpeg: truncated scan",
    );
    const scans = Array.from({ length: 256 }, () => [...sos([[1, 0x00]]), ...ENTROPY]);
    const base = [SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), ...scans];
    expect(refusal(jpeg(...base, seg(0xda, []), EOI)).message).toBe("jpeg: too many scans");
  });

  it("the end of the stream: no frame is reported before no scan, and no scan before four components without Adobe", () => {
    expect(refusal(jpeg(SOI, EOI)).message).toBe("jpeg: no frame");
    expect(refusal(jpeg(SOI, dqt(), EOI)).message).toBe("jpeg: no frame");
    const cmyk: [number, number, number][] = [
      [1, 0x11, 0],
      [2, 0x11, 0],
      [3, 0x11, 0],
      [4, 0x11, 0],
    ];
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, cmyk), EOI)).message).toBe("jpeg: no scan");
  });

  it("a scan's entropy data is a view on the caller's bytes, the next segment starts where it ends", () => {
    const bytes = jpeg(
      SOI,
      dqt(),
      sof(0xc2, 8, 8, GREY),
      dht(0, 0),
      dht(1, 0),
      sos([[1, 0]], 0, 0),
      [1, 2, 3],
      sos([[1, 0]], 1, 63),
      [4, 5],
      EOI,
    );
    const s = parseJpeg(bytes, GUARD);
    expect(s.scans.map((x) => [...x.data])).toEqual([
      [1, 2, 3],
      [4, 5],
    ]);
    expect(s.scans[0]!.data.buffer).toBe(bytes.buffer);
    expect(s.scans[1]!.data.byteOffset).toBe(bytes.length - 2 - 2);
  });

  it("a scan before its Huffman tables is refused as missing; a scan whose entropy segment is empty is kept, with no bytes", () => {
    expect(refusal(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), sos([[1, 0]]), EOI)).message).toBe(
      "jpeg: missing huffman table",
    );
    const empty = parseJpeg(jpeg(SOI, dqt(), sof(0xc0, 8, 8, GREY), dht(0, 0), dht(1, 0), sos([[1, 0]]), EOI), GUARD);
    expect(empty.scans).toHaveLength(1);
    expect(empty.scans[0]!.data).toHaveLength(0);
  });
});
