import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { IRBlock, IRTab } from "doc-ir";
import { encodePng } from "doc-ir/internal";
import type { Clock } from "../lib/clock";
import type { SlideImage, SlideImagePort, SlidePage } from "../lib/slide-image-port";
import { extractPptx, type GoogleSlidesSource, pageIdsByListIndex } from "./pptx";

/** The in-memory Google Slides image side, this test's own: a deck's pages, the PNG of some of them, the time the deck
 *  was last modified, and every call made. The port is declared by the library and implemented by its caller. */
class FakeSlideImagePort implements SlideImagePort {
  private readonly pagesById = new Map<string, SlidePage[]>();
  private readonly imagesByPage = new Map<string, Uint8Array>();
  private readonly modifiedTimes = new Map<string, string>();
  /** Every call, in order: `pages:<presentationId>`, `image:<presentationId>/<pageObjectId>` or `modified:<fileId>`. */
  readonly calls: string[] = [];

  setPages(presentationId: string, pages: SlidePage[]): void {
    this.pagesById.set(presentationId, pages);
  }

  setImage(presentationId: string, pageObjectId: string, bytes: Uint8Array): void {
    this.imagesByPage.set(`${presentationId}/${pageObjectId}`, bytes);
  }

  setModifiedTime(fileId: string, modifiedTime: string): void {
    this.modifiedTimes.set(fileId, modifiedTime);
  }

  /** No pages set: `undefined`, as a caller answers a failure. */
  async slidePages(presentationId: string): Promise<SlidePage[] | undefined> {
    this.calls.push(`pages:${presentationId}`);
    return this.pagesById.get(presentationId);
  }

  /** None set: `false`, as a refused lookup. */
  async unchangedSince(fileId: string, modifiedTime: string): Promise<boolean> {
    this.calls.push(`modified:${fileId}`);
    return this.modifiedTimes.get(fileId) === modifiedTime;
  }

  /** No image set for the page: `undefined`, as a refused thumbnail. */
  async slideImage(presentationId: string, pageObjectId: string): Promise<SlideImage | undefined> {
    this.calls.push(`image:${presentationId}/${pageObjectId}`);
    const bytes = this.imagesByPage.get(`${presentationId}/${pageObjectId}`);
    return bytes === undefined ? undefined : { bytes };
  }
}

/** A clock that never moves: no deck of these tests spends its time budget unless a test gives its own clock. */
const STILL_CLOCK: Clock = { now: () => 0 };

// MIKSUDOC « Spécification fonctionnelle » §9.3.4: a DIAGRAM slide of a Google Slides deck gets
// Google's own image instead of our SVG; a stored .pptx never asks; any
// refusal keeps the SVG. Made-up decks: two boxes joined by an arrow make a
// diagram slide, one box alone a flow slide.

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (inner: string): string =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

const at = (x: number, y: number, cx: number, cy: number): string =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
const box = (id: number, x: number, words: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr>${at(x, 0, 2_000_000, 1_000_000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
const arrow =
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9" name="c9"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
  `<p:spPr>${at(2_000_000, 500_000, 3_000_000, 0)}<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>` +
  '<a:ln><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>';

type Made = "diagram" | "flow" | "hidden-diagram";
const treeOf = (kind: Made, n: number): string =>
  kind === "flow" ? box(2, 0, `Only ${n}`) : box(2, 0, `Alpha ${n}`) + arrow + box(3, 5_000_000, `Omega ${n}`);

const deckOf = async (slides: readonly Made[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file("_rels/.rels", RELS(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0"?><p:presentation ${P_NS}><p:sldIdLst>` +
      slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("") +
      '</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>',
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    RELS(
      slides
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join(""),
    ),
  );
  slides.forEach((kind, i) => {
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld ${P_NS}${kind === "hidden-diagram" ? ' show="0"' : ""}><p:cSld><p:spTree>${treeOf(kind, i + 1)}</p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "uint8array" });
};

const tabsOf = (blocks: readonly IRBlock[]): IRTab[] => {
  const tabs = blocks[0];
  if (tabs?.kind !== "tabs") throw new Error("expected one tabs block");
  return tabs.tabs;
};

/** The image block of a tab: its mime type and pixel width. */
const imageOf = (tab: IRTab): { mimeType: string; pixelWidth?: number } => {
  const image = tab.blocks.find((b) => b.kind === "image");
  if (image?.kind !== "image") throw new Error("expected an image block");
  return {
    mimeType: image.ref.mimeType,
    ...(image.ref.pixelWidth === undefined ? {} : { pixelWidth: image.ref.pixelWidth }),
  };
};

const PRESENTATION = "pres-1";
/** The Drive `modifiedTime` of the listing the export was made from. */
const EXPORTED_AT = "2026-09-25T10:00:00.000Z";

/** A fake port whose deck has not changed since the export. */
const fakePort = (): FakeSlideImagePort => {
  const port = new FakeSlideImagePort();
  port.setModifiedTime(PRESENTATION, EXPORTED_AT);
  return port;
};

const from = (port: SlideImagePort, clock: Clock = STILL_CLOCK): GoogleSlidesSource => ({
  presentationId: PRESENTATION,
  port,
  clock,
  modifiedTime: EXPORTED_AT,
});
/** Distinct PNGs, 800 px wide like a MEDIUM thumbnail, so each page's image is its own. */
const thumbnail = (shade: number): Uint8Array =>
  encodePng(
    800,
    1,
    "rgb-24",
    Array.from({ length: 800 * 3 }, () => shade),
  )!;

describe("extractPptx — Google's own image of a diagram slide (MIKSUDOC « Spécification fonctionnelle » §9.3.4)", () => {
  it("a diagram slide of a Google Slides deck is Google's PNG, followed by the same collapsed text; one pages call, one image call", async () => {
    const port = fakePort();
    port.setPages(PRESENTATION, [{ objectId: "p1", skipped: false }]);
    port.setImage(PRESENTATION, "p1", thumbnail(10));
    const google = await extractPptx(await deckOf(["diagram"]), from(port));
    const stored = await extractPptx(await deckOf(["diagram"]));

    // Failure value: the stored deck's SVG.
    expect(imageOf(tabsOf(stored.doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
    const tab = tabsOf(google.doc.blocks)[0]!;
    // The collapsed text lives AFTER the tabs since option A (27/09/2026) —
    // identical whichever image (Google's PNG or our SVG) the tab shows.
    expect(tab.blocks.map((b) => b.kind)).toEqual(["image"]);
    expect(imageOf(tab)).toEqual({ mimeType: "image/png", pixelWidth: 800 });
    expect(google.doc.blocks[1]).toEqual(stored.doc.blocks[1]);
    expect(google.doc.blocks[1]).toMatchObject({ kind: "diagramText", sections: [{ label: "Slide 1" }] });
    expect(google.images).toHaveLength(1);
    expect(google.images[0]!.mimeType).toBe("image/png");
    expect(google.images[0]!.selfDescribed).toBeUndefined();
    expect(port.calls).toEqual([`pages:${PRESENTATION}`, `image:${PRESENTATION}/p1`, `modified:${PRESENTATION}`]);
  });

  it("a deck without a diagram slide makes no call at all", async () => {
    const port = fakePort();
    port.setPages(PRESENTATION, [{ objectId: "p1", skipped: false }]);
    await extractPptx(await deckOf(["flow"]), from(port));
    expect(port.calls).toEqual([]);
  });

  it("a refused thumbnail, refused pages, or a port that throws keep our SVG for that slide; the document never fails", async () => {
    const refusedImage = fakePort();
    refusedImage.setPages(PRESENTATION, [
      { objectId: "p1", skipped: false },
      { objectId: "p2", skipped: false },
    ]);
    refusedImage.setImage(PRESENTATION, "p2", thumbnail(20));
    const one = await extractPptx(await deckOf(["diagram", "diagram"]), from(refusedImage));
    expect(tabsOf(one.doc.blocks).map((t) => imageOf(t).mimeType)).toEqual(["image/svg+xml", "image/png"]);

    const refusedPages = fakePort();
    const two = await extractPptx(await deckOf(["diagram"]), from(refusedPages));
    expect(imageOf(tabsOf(two.doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
    expect(refusedPages.calls).toEqual([`pages:${PRESENTATION}`]);

    const throwing: SlideImagePort = {
      slidePages: async (): Promise<SlidePage[]> => [{ objectId: "p1", skipped: false }],
      slideImage: async (): Promise<SlideImage> => {
        throw new Error("network");
      },
      unchangedSince: async (): Promise<boolean> => {
        return true;
      },
    };
    const three = await extractPptx(await deckOf(["diagram"]), from(throwing));
    expect(imageOf(tabsOf(three.doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
  });

  it("bytes that are not a PNG with a readable width keep our SVG", async () => {
    const port = fakePort();
    port.setPages(PRESENTATION, [{ objectId: "p1", skipped: false }]);
    port.setImage(PRESENTATION, "p1", new TextEncoder().encode("<html>not an image</html>"));
    const { doc } = await extractPptx(await deckOf(["diagram"]), from(port));
    expect(imageOf(tabsOf(doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
  });

  it("a hidden slide counts in the mapping: the third slide of the file gets the third page, never the second", async () => {
    const port = fakePort();
    port.setPages(PRESENTATION, [
      { objectId: "p1", skipped: false },
      { objectId: "p2", skipped: true },
      { objectId: "p3", skipped: false },
    ]);
    for (const [id, shade] of [
      ["p1", 1],
      ["p2", 2],
      ["p3", 3],
    ] as const)
      port.setImage(PRESENTATION, id, thumbnail(shade));
    const { doc } = await extractPptx(await deckOf(["diagram", "hidden-diagram", "diagram"]), from(port));
    expect(tabsOf(doc.blocks)).toHaveLength(2);
    // Failure value (visible rank as the page index): p1 then p2.
    expect(port.calls).toEqual([
      `pages:${PRESENTATION}`,
      `image:${PRESENTATION}/p1`,
      `image:${PRESENTATION}/p3`,
      `modified:${PRESENTATION}`,
    ]);
  });

  it("an export that leaves the skipped slide out maps to the pages not skipped; counts that agree neither way give no image at all", async () => {
    const pages: SlidePage[] = [
      { objectId: "p1", skipped: false },
      { objectId: "p2", skipped: true },
      { objectId: "p3", skipped: false },
    ];
    const port = fakePort();
    port.setPages(PRESENTATION, pages);
    await extractPptx(await deckOf(["diagram", "diagram"]), from(port));
    expect(port.calls).toEqual([`pages:${PRESENTATION}`, `image:${PRESENTATION}/p1`, `image:${PRESENTATION}/p3`]);

    const disagree = fakePort();
    disagree.setPages(PRESENTATION, pages);
    const { doc } = await extractPptx(await deckOf(["diagram"]), from(disagree));
    expect(disagree.calls).toEqual([`pages:${PRESENTATION}`]);
    expect(imageOf(tabsOf(doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
  });
});

describe("pageIdsByListIndex", () => {
  const pages: SlidePage[] = [
    { objectId: "a", skipped: false },
    { objectId: "b", skipped: true },
    { objectId: "c", skipped: false },
  ];
  it("same count: position for position, skipped pages included", () => {
    expect(pageIdsByListIndex(3, pages)).toEqual(["a", "b", "c"]);
  });
  it("as many as the pages not skipped: those, in order", () => {
    expect(pageIdsByListIndex(2, pages)).toEqual(["a", "c"]);
  });
  it("any other count: undefined", () => {
    expect(pageIdsByListIndex(1, pages)).toBeUndefined();
    expect(pageIdsByListIndex(4, pages)).toBeUndefined();
  });
});

describe("extractPptx — Google's images are bounded and checked (review fixes, 25/09/2026)", () => {
  const threeDiagrams = (): FakeSlideImagePort => {
    const port = fakePort();
    port.setPages(PRESENTATION, [
      { objectId: "p1", skipped: false },
      { objectId: "p2", skipped: false },
      { objectId: "p3", skipped: false },
    ]);
    for (const [id, shade] of [
      ["p1", 1],
      ["p2", 2],
      ["p3", 3],
    ] as const)
      port.setImage(PRESENTATION, id, thumbnail(shade));
    return port;
  };

  it("past the deck's 120 s budget, the remaining diagram slides keep our SVG without asking", async () => {
    const port = threeDiagrams();
    // Every image call takes 70 s of the clock: asked at 0 s and 70 s; at 140 s the budget is spent.
    let t = 0;
    const clock: Clock = { now: () => t };
    const slow: SlideImagePort = {
      slidePages: (id) => port.slidePages(id),
      slideImage: async (id, page) => {
        t += 70_000;
        return port.slideImage(id, page);
      },
      unchangedSince: (id, at) => port.unchangedSince(id, at),
    };
    const { doc } = await extractPptx(await deckOf(["diagram", "diagram", "diagram"]), from(slow, clock));
    // Failure value: three PNGs, three image calls.
    expect(tabsOf(doc.blocks).map((tab) => imageOf(tab).mimeType)).toEqual(["image/png", "image/png", "image/svg+xml"]);
    expect(port.calls.filter((c) => c.startsWith("image:"))).toEqual([
      `image:${PRESENTATION}/p1`,
      `image:${PRESENTATION}/p2`,
    ]);
  });

  it("a deck modified since the export (or whose modifiedTime cannot be read) drops EVERY Google image", async () => {
    const changed = threeDiagrams();
    changed.setModifiedTime(PRESENTATION, "2026-09-25T10:00:05.000Z");
    const one = await extractPptx(await deckOf(["diagram", "diagram", "diagram"]), from(changed));
    // Failure value: three PNGs (no check after the thumbnails).
    expect(tabsOf(one.doc.blocks).map((tab) => imageOf(tab).mimeType)).toEqual([
      "image/svg+xml",
      "image/svg+xml",
      "image/svg+xml",
    ]);
    expect(one.images.every((i) => i.mimeType === "image/svg+xml")).toBe(true);

    const unreadable = threeDiagrams();
    unreadable.setModifiedTime(PRESENTATION, "");
    const two = await extractPptx(await deckOf(["diagram"]), from(unreadable));
    expect(imageOf(tabsOf(two.doc.blocks)[0]!).mimeType).toBe("image/svg+xml");
  });

  it("Google's PNG bytes stay out of the content: the same page and revision give the same document, another revision another", async () => {
    const withShade = (shade: number): FakeSlideImagePort => {
      const port = fakePort();
      port.setPages(PRESENTATION, [{ objectId: "p1", skipped: false }]);
      port.setImage(PRESENTATION, "p1", thumbnail(shade));
      return port;
    };
    const a = await extractPptx(await deckOf(["diagram"]), from(withShade(10)));
    const b = await extractPptx(await deckOf(["diagram"]), from(withShade(11)));
    // Failure value: two sha256, one per rendering of the same revision.
    expect(b.doc).toEqual(a.doc);
    const laterPort = withShade(10);
    laterPort.setModifiedTime(PRESENTATION, "2026-09-26T08:00:00.000Z");
    const later = await extractPptx(await deckOf(["diagram"]), {
      ...from(laterPort),
      modifiedTime: "2026-09-26T08:00:00.000Z",
    });
    expect(later.doc).not.toEqual(a.doc);
    // The attachment still carries Google's bytes.
    expect(a.images[0]!.bytes).toEqual(thumbnail(10));
  });
});
