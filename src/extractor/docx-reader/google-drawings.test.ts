// D2 voie A (MIKSUDOC « Spécification fonctionnelle » §9.5): Google's own image of a
// drawing, taken from the zipped HTML export (`files.export`,
// application/zip — https://developers.google.com/workspace/drive/api/guides/ref-export-formats),
// replaces our SVG only when ONE image of that export matches the drawing by
// what both files declare.

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { type HtmlExportImage, matchDrawings, readHtmlExport } from "./google-drawings";

describe("readHtmlExport — the <img> of Google's zipped HTML export", () => {
  it("bytes that are not a zip: no export to read (abstention, never a throw)", async () => {
    expect(await readHtmlExport(new TextEncoder().encode("not a zip"))).toBeUndefined();
  });

  it("an archive without an .html part: no export to read", async () => {
    const zip = new JSZip();
    zip.file("images/image1.png", new Uint8Array(4));
    expect(await readHtmlExport(new Uint8Array(await zip.generateAsync({ type: "uint8array" })))).toBeUndefined();
  });
});

const image = (over: Partial<HtmlExportImage> = {}): HtmlExportImage => ({
  path: "images/a.png",
  widthPx: 560,
  heightPx: 94,
  previousText: "before",
  ...over,
});

describe("matchDrawings — exactly one image, or abstain", () => {
  it("two drawings of the same size after the same paragraph: neither is matched", () => {
    const probe = { widthPx: 560, heightPx: 94, previousText: "before" };
    expect(matchDrawings([probe, probe], [image({ path: "a.png" }), image({ path: "b.png" })])).toEqual([
      undefined,
      undefined,
    ]);
  });

  it("two HTML images that fit one drawing: several matches, abstain", () => {
    expect(
      matchDrawings(
        [{ widthPx: 560, heightPx: 94, previousText: "before" }],
        [image({ path: "a.png" }), image({ path: "b.png" })],
      ),
    ).toEqual([undefined]);
  });

  it("the same size but another paragraph before: no match", () => {
    expect(
      matchDrawings([{ widthPx: 560, heightPx: 94, previousText: "before" }], [image({ previousText: "elsewhere" })]),
    ).toEqual([undefined]);
  });

  it("the same paragraph but another size: no match", () => {
    expect(matchDrawings([{ widthPx: 560, heightPx: 94, previousText: "before" }], [image({ widthPx: 300 })])).toEqual([
      undefined,
    ]);
  });

  it("two same-size drawings after different paragraphs each find their own image", () => {
    const html = [image({ path: "a.png", previousText: "one" }), image({ path: "b.png", previousText: "two" })];
    const probe = (previousText: string) => ({ widthPx: 560, heightPx: 94, previousText });
    expect(matchDrawings([probe("one"), probe("two")], html).map((m) => m?.path)).toEqual(["a.png", "b.png"]);
  });
});
