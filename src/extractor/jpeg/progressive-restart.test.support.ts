/** Test support: a hand-assembled progressive greyscale JPEG with a restart interval (DRI), 32 x 16 pixels (a grid of
 *  4 x 2 blocks), made of a DC-first scan and an AC-first scan. The coefficients are chosen, not transformed from pixels:
 *  the expected image is known from the choice (`BLOCK_DC`, `BLOCK_AC1`).
 *  NOT here: any decoding or assertion (`jpeg/progressive-restart.test.ts`), refinement scans.
 *  Called by: `jpeg/progressive-restart.test.ts`. */

const u16 = (n: number): number[] => [(n >> 8) & 0xff, n & 0xff];
const seg = (marker: number, payload: number[]): number[] => [0xff, marker, ...u16(payload.length + 2), ...payload];

/** Pads a bit string to a byte boundary with 1-bits (T.81 F.1.2.3) and stuffs 0xFF. */
const bytesOf = (bits: string): number[] => {
  const padded = bits + "1".repeat((8 - (bits.length % 8)) % 8);
  return (padded.match(/.{8}/g) ?? []).flatMap((b) => {
    const v = parseInt(b, 2);
    return v === 0xff ? [0xff, 0x00] : [v];
  });
};

/** Entropy-coded segment: one bit string per restart interval, joined by RST0..RST7. */
const intervals = (parts: readonly string[]): number[] =>
  parts.flatMap((bits, n) => (n === 0 ? [] : [0xff, 0xd0 + ((n - 1) % 8)]).concat(bytesOf(bits)));

export const WIDTH = 32;
export const HEIGHT = 16;
/** Blocks per restart interval (the DRI value). */
export const RESTART_BLOCKS = 2;
/** Quantiser of every coefficient: a DC step of 1 moves a block's mean by 64 / 8 = 8 levels. */
export const QUANT = 64;
/** DC coefficient of each block, raster order. */
export const BLOCK_DC = [0, 1, 2, 3, -1, -2, 0, 1] as const;
/** Coefficient 1 (zigzag, first horizontal frequency) of each block: +1 puts light on the left, -1 on the right. */
export const BLOCK_AC1 = [1, 0, -1, 0, 1, 0, 0, -1] as const;

/** One-table specs of 2-bit codes 00, 01, 10 for three symbols. */
const dht = (tableClass: 0 | 1, values: number[]): number[] =>
  seg(0xc4, [tableClass << 4, 0, 3, ...new Array<number>(14).fill(0), ...values]);

/**
 * DC table: categories [0, 1, 2] -> 00, 01, 10. The prediction restarts at each interval, so the differences are taken
 * from zero at blocks 0, 2, 4, 6.
 */
const dcIntervals = (): string[] => {
  const code = ["00", "01", "10"];
  const magnitude = (diff: number): string => {
    const category = Math.abs(diff) === 0 ? 0 : Math.abs(diff) === 1 ? 1 : 2;
    if (category === 0) return code[0]!;
    const width = category;
    const raw = diff > 0 ? diff : diff + (1 << category) - 1;
    return code[category]! + raw.toString(2).padStart(width, "0");
  };
  const out: string[] = [];
  for (let start = 0; start < BLOCK_DC.length; start += RESTART_BLOCKS) {
    let previous = 0;
    let bits = "";
    for (let b = start; b < start + RESTART_BLOCKS; b++) {
      bits += magnitude(BLOCK_DC[b]! - previous);
      previous = BLOCK_DC[b]!;
    }
    out.push(bits);
  }
  return out;
};

/**
 * AC table: [0x01 (run 0, size 1), 0x00 (EOB), 0x10 (EOB run 2^1+bit)] -> 00, 01, 10. With `overlongRun`, the last block
 * of the first interval ends in an EOB run that reaches into the next interval (the second block of the image is
 * followed by a pending run of one block): an invalid stream, produced here on purpose. The test shows that the decoder
 * resets the EOB run to 0 at the restart, dropping the pending run. Without it, that block is a plain EOB.
 */
const acIntervals = (overlongRun: boolean): string[] => {
  const out: string[] = [];
  for (let start = 0; start < BLOCK_AC1.length; start += RESTART_BLOCKS) {
    let bits = "";
    for (let b = start; b < start + RESTART_BLOCKS; b++) {
      const ac = BLOCK_AC1[b]!;
      if (ac !== 0) bits += "00" + (ac > 0 ? "1" : "0") + "01";
      else bits += overlongRun && b === 1 ? "10" + "0" : "01";
    }
    out.push(bits);
  }
  return out;
};

export const progressiveWithRestarts = (options: { overlongRun: boolean }): Uint8Array =>
  new Uint8Array([
    0xff,
    0xd8,
    ...seg(0xdb, [0, ...new Array<number>(64).fill(QUANT)]),
    ...seg(0xc2, [8, ...u16(HEIGHT), ...u16(WIDTH), 1, 1, 0x11, 0]),
    ...dht(0, [0, 1, 2]),
    ...dht(1, [0x01, 0x00, 0x10]),
    ...seg(0xdd, u16(RESTART_BLOCKS)),
    // DC first: Ss 0, Se 0, Ah 0, Al 0.
    ...seg(0xda, [1, 1, 0x00, 0, 0, 0x00]),
    ...intervals(dcIntervals()),
    // AC first: Ss 1, Se 5, Ah 0, Al 0.
    ...seg(0xda, [1, 1, 0x00, 1, 5, 0x00]),
    ...intervals(acIntervals(options.overlongRun)),
    0xff,
    0xd9,
  ]);
