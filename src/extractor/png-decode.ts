/** A PNG decoded to raw pixels, with no dependency (`node:zlib` inflates) —
 *  so a cropped picture can be cut to its visible part and re-encoded.
 *  Only the variants it can decode exactly: 8 bits per channel, grey, grey +
 *  alpha, RGB, RGBA, not interlaced, no `tRNS` transparent colour. Any other
 *  PNG gives `undefined`: the caller then publishes nothing rather than a
 *  wrong image.
 *  NOT here: encoding (`extractor/png.ts`), undoing the row filters (`extractor/png-unfilter.ts`), nor what to cut
 *  (`extractor/pptx-reader/picture.ts`).
 *  Called by: `extractor/pptx-reader/picture.ts`, `extractor/image-reduction.ts`. */

import { inflateSync } from "node:zlib";

import { concatBytes, u32be as u32 } from "doc-ir/internal";
import { isPng } from "doc-ir/internal";
import { channelsOfColorType } from "doc-ir/internal";
import { unfilterPngRows } from "doc-ir/internal";

export interface DecodedPng {
  width: number;
  height: number;
  /** 1 grey, 2 grey + alpha, 3 RGB, 4 RGBA. */
  channels: 1 | 2 | 3 | 4;
  /** `height` rows of `width × channels` bytes, no filter byte. */
  pixels: Uint8Array;
}

// Palette (type 3) is left out of `png-color.ts`: none of the cropped pictures of the USA deck is one
// (all RGB or RGBA, read 25/09/2026), and the encoder has no palette.

/**
 * `maxPixels` refuses an image before anything is inflated; the inflated
 * stream is also capped at the exact size the header announces, so a small
 * file cannot unpack into a large one.
 */
export const decodePng = (bytes: Uint8Array, maxPixels: number): DecodedPng | undefined => {
  if (bytes.length < 33 || !isPng(bytes)) return undefined;
  let width = 0;
  let height = 0;
  let channels: 1 | 2 | 3 | 4 | undefined;
  const idat: Uint8Array[] = [];
  let at = 8;
  while (at + 8 <= bytes.length) {
    const length = u32(bytes, at);
    const type = String.fromCharCode(bytes[at + 4] ?? 0, bytes[at + 5] ?? 0, bytes[at + 6] ?? 0, bytes[at + 7] ?? 0);
    const data = bytes.subarray(at + 8, at + 8 + length);
    if (data.length !== length) return undefined;
    if (type === "IHDR") {
      width = u32(data, 0);
      height = u32(data, 4);
      const [depth, colorType, compression, filter, interlace] = [data[8], data[9], data[10], data[11], data[12]];
      channels = colorType === undefined ? undefined : channelsOfColorType(colorType);
      if (depth !== 8 || channels === undefined || compression !== 0 || filter !== 0 || interlace !== 0)
        return undefined;
    } else if (type === "tRNS") {
      // A transparent colour on a grey or RGB image: dropping it would make
      // the transparent areas opaque.
      return undefined;
    } else if (type === "IDAT") {
      idat.push(data);
    } else if (type === "IEND") {
      break;
    }
    at += 12 + length; // length, type, data, CRC
  }
  if (channels === undefined || width === 0 || height === 0 || width * height > maxPixels || idat.length === 0)
    return undefined;

  const stride = width * channels;
  const expected = height * (stride + 1);
  let raw: Uint8Array;
  try {
    raw = inflateSync(concatBytes(idat), { maxOutputLength: expected });
  } catch {
    return undefined;
  }
  if (raw.length !== expected) return undefined;

  const pixels = unfilterPngRows(raw, width, height, channels);
  return pixels === undefined ? undefined : { width, height, channels, pixels };
};
