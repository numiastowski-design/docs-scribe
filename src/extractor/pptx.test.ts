import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { ReaderError } from "doc-ir";
import type { IRBlock } from "doc-ir";
import { encodePng } from "doc-ir/internal";
import { MAX_PRESENTATION_TABS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";
import { extractPptx, foldNotices, createTabCapNotice } from "./pptx";
import {
  arrowConnector as connector,
  deckOf,
  diagramDeck as deck,
  diagramsOf,
  graphicFrame as frame,
  type MadeSlide,
  OFFICE_REL as REL,
  ONE_PIXEL_PNG,
  placedBox as box,
  shapeAt as at,
  tabsOf,
  textShape as text,
  titleShape as title,
} from "./pptx-deck.test.support";

// Witnesses: `fixtures/pptx/03` (a chart, speaker notes, two pictures); made-up decks for what no real
// one shows (the own pptx reader lot, 25/09/2026). Two real third-party decks of the measurement corpus (a 27-slide
// support deck, a 16-slide deck of tables) were witnesses here too: their cases are held by the publisher's
// application, which carries them.
// The READER alone is held here: what the published page does with the same decks is held by the publisher's
// application, on the decks of `pptx-deck.test.support.ts`.
// The 30 s deadline dates from when this file extracted the 27-slide deck twice (about one second alone, past 5 s once
// under the full suite's parallel load); kept as it stood.
vi.setConfig({ testTimeout: 30_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));

/** Every block of a tab, lists, tables and cells walked into. */
const allBlocks = (blocks: readonly IRBlock[]): IRBlock[] =>
  blocks.flatMap((b) => [
    b,
    ...(b.kind === "list" ? allBlocks(b.items.flat()) : []),
    ...(b.kind === "table" ? allBlocks(b.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks))) : []),
  ]);

const textOf = (blocks: readonly IRBlock[]): string =>
  allBlocks(blocks)
    .flatMap((b) => (b.kind === "paragraph" || b.kind === "heading" ? [b.inlines.map((i) => i.text).join("")] : []))
    .join("\n");

describe("extractPptx — fixtures/pptx/03 (a chart, notes, two pictures)", () => {
  it("one tab per slide, 'Slide N' when no shape declares a title", async () => {
    const { doc } = await extractPptx(
      fixture("pptx/03-Presentation-Resultats-T1.pptx"),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    expect(tabsOf(doc.blocks).map((t) => t.title)).toEqual([
      "Slide 1",
      "Slide 2",
      "Slide 3",
      "Slide 4",
      "Slide 5",
      "Slide 6",
    ]);
  });

  it("slide 3's chart is its data table and its title, as before step A; no picture (none in the archive)", async () => {
    const { doc } = await extractPptx(
      fixture("pptx/03-Presentation-Resultats-T1.pptx"),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const chart = tabsOf(doc.blocks)[2]!.blocks.find((b) => b.kind === "chart");
    expect(chart).toEqual({
      kind: "chart",
      title: "Évolution trimestrielle du chiffre d'affaires",
      data: {
        headers: ["", "Chiffre d'affaires"],
        rows: [
          ["T2 2025", "1,840"],
          ["T3 2025", "1,955"],
          ["T4 2025", "2,100"],
          ["T1 2026", "2,276"],
        ],
      },
      drawing: { type: "bar", values: [[1840, 1955, 2100, 2276]] },
    });
  });

  it("a picture part whose stream is corrupt is left out of the page, and named on the dashboard (lot A, 29/09/2026 — no more marker on the slide)", async () => {
    // Pin of the notice's exact shape; the mechanism's red witness is the
    // reader's own damaged-part test (pptx-reader.test.ts, arbitrage 28/09).
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
    const name = new TextEncoder().encode("ppt/media/image-4-1.png");
    const view = new DataView(bytes.buffer, bytes.byteOffset);
    let nameAt = -1;
    for (let i = 30; i < bytes.length - name.length; i++) {
      if (!name.every((b, j) => bytes[i + j] === b)) continue;
      if (view.getUint32(i - 30, true) === 0x04034b50 && view.getUint16(i - 30 + 26, true) === name.length) {
        nameAt = i;
        break;
      }
    }
    expect(nameAt).toBeGreaterThan(0);
    const dataStart = nameAt + name.length + view.getUint16(nameAt - 30 + 28, true);
    for (let i = 0; i < 8; i++) bytes[dataStart + i] = (bytes[dataStart + i] ?? 0) ^ 0xff;
    const { doc, notices } = await extractPptx(bytes, undefined, undefined, PAGE_SIZED_LIMITS);
    const texts = tabsOf(doc.blocks).map((t) => textOf(t.blocks));
    expect(texts.join("\n")).not.toContain("Unsupported");
    expect(notices).toContainEqual({
      kind: "unsupported-element",
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment -- vitest types its asymmetric matcher `expect.any` as `any`.
      slide: expect.any(Number),
      what: "damaged-part",
      outcome: "published",
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment -- vitest types its asymmetric matcher `expect.any` as `any`.
      count: expect.any(Number),
    });
  });

  it("speaker notes are never read: slide 2's note is nowhere in the document", async () => {
    // Failure value: this sentence is slide 2's speaker note in the fixture,
    // read into the document until 25/09/2026.
    const note = "dynamique commerciale du secteur énergie";
    const { doc } = await extractPptx(
      fixture("pptx/03-Presentation-Resultats-T1.pptx"),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    expect(JSON.stringify(doc)).not.toContain(note);
  });

  it("slides 4 and 5 each hold one picture, with its width on the slide and no alt text", async () => {
    const { doc, images } = await extractPptx(
      fixture("pptx/03-Presentation-Resultats-T1.pptx"),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const tabs = tabsOf(doc.blocks);
    for (const tab of [tabs[3]!, tabs[4]!]) {
      const pictures = tab.blocks.filter((b) => b.kind === "image");
      expect(pictures).toHaveLength(1);
      const picture = pictures[0]!;
      if (picture.kind !== "image") throw new Error("image");
      expect(picture.alt).toBeUndefined();
      expect(picture.ref.widthPt).toBeGreaterThan(0);
    }
    expect(images.map((i) => i.mimeType)).toEqual(["image/png", "image/png"]);
  });

  it("EXTRACT_FAILED on bytes that aren't a valid zip", async () => {
    await expect(
      extractPptx(new Uint8Array([1, 2, 3, 4, 5]), undefined, undefined, PAGE_SIZED_LIMITS),
    ).rejects.toSatisfy((error) => error instanceof ReaderError && error.code === "EXTRACT_FAILED");
  });
});

describe("extractPptx — no slide cap for presentations (decision (3), 25/09/2026)", () => {
  it("a deck of 16 visible slides gives 16 tabs; hidden slides give none", async () => {
    const slides: MadeSlide[] = Array.from({ length: 18 }, (_, i) => ({
      tree: text(2, `slide ${i + 1}`),
      hidden: i === 3 || i === 7,
    }));
    const { doc } = await extractPptx(await deckOf(slides), undefined, undefined, PAGE_SIZED_LIMITS);
    const tabs = tabsOf(doc.blocks);
    expect(tabs).toHaveLength(16);
    expect(textOf(tabs[3]!.blocks)).toBe("slide 5");
    // N is the rank among VISIBLE slides: the fifth slide of the file is the fourth tab.
    expect(tabs[3]!.title).toBe("Slide 4");
  });
});

describe("extractPptx — titles, frames and pictures", () => {
  it("a declared title is the tab's heading, never its label; repeated from the previous slide, it loses its heading", async () => {
    const { doc } = await extractPptx(
      await deckOf([
        { tree: title(2, "Plan") + text(3, "one") },
        { tree: title(2, "Plan") + text(3, "two") },
        { tree: text(3, "three") },
      ]),
    );
    const tabs = tabsOf(doc.blocks);
    expect(tabs.map((t) => t.title)).toEqual(["Slide 1", "Slide 2", "Slide 3"]);
    expect(tabs.map((t) => t.blocks.map((b) => b.kind))).toEqual([
      ["heading", "paragraph"],
      ["paragraph"],
      ["paragraph"],
    ]);
  });

  it("a PowerPoint table is a table; SmartArt and an OLE object contribute nothing to the page, each a dashboard notice; a chart without its part is a chart-withheld notice (07/10/2026: a missing part is said, the same notice in every format) (lot A, 29/09/2026)", async () => {
    const { doc, notices } = await extractPptx(
      await deckOf([
        {
          tree:
            frame(
              2,
              "http://schemas.openxmlformats.org/drawingml/2006/table",
              "<a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl>",
            ) +
            frame(
              3,
              "http://schemas.openxmlformats.org/drawingml/2006/diagram",
              '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2"/>',
            ) +
            frame(
              4,
              "http://schemas.openxmlformats.org/drawingml/2006/chart",
              '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdMissing"/>',
            ) +
            frame(5, "http://schemas.openxmlformats.org/presentationml/2006/ole", "<p:oleObj/>"),
        },
      ]),
    );
    expect(tabsOf(doc.blocks)[0]!.blocks).toEqual([
      { kind: "table", rows: [{ cells: [{ blocks: [{ kind: "paragraph", inlines: [{ text: "cell" }] }] }] }] },
    ]);
    expect(notices).toHaveLength(3);
    expect(notices).toContainEqual({
      kind: "unsupported-element",
      slide: 1,
      what: "smartart",
      outcome: "converted",
      count: 1,
    });
    expect(notices).toContainEqual({
      kind: "unsupported-element",
      slide: 1,
      what: "ole",
      outcome: "converted",
      count: 1,
    });
    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
  });

  it("a picture that is not published contributes nothing to the page, a dashboard notice instead (lot A, 29/09/2026)", async () => {
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:link="rId9"/></p:blipFill><p:spPr/></p:pic>';
    const { doc, images, notices } = await extractPptx(
      await deckOf([
        {
          tree: pic,
          rels: `<Relationship Id="rId9" Type="${REL}/image" Target="https://example.com/a.png" TargetMode="External"/>`,
        },
      ]),
    );
    expect(tabsOf(doc.blocks)[0]!.blocks).toEqual([{ kind: "paragraph", inlines: [] }]);
    expect(notices).toEqual([
      { kind: "unsupported-element", slide: 1, what: "external-link", outcome: "published", count: 1 },
    ]);
    expect(images).toEqual([]);
  });

  it("picture-not-drawn notices fold the same way: one entry per (slide, reason) with a count (item E)", () => {
    const one = { kind: "picture-not-drawn", slide: 3, reason: "format", count: 1 } as const;
    expect(foldNotices([one, one, one, { ...one, reason: "over-cap" }, { ...one, slide: 4 }])).toEqual([
      { ...one, count: 3 },
      { ...one, reason: "over-cap" },
      { ...one, slide: 4 },
    ]);
  });

  it("notices FOLD: twelve identical unsupported frames on a slide are ONE entry with count 12, distinct kinds stay distinct (review 30/09/2026, item E)", async () => {
    const ole = (id: number) => frame(id, "http://schemas.openxmlformats.org/presentationml/2006/ole", "<p:oleObj/>");
    const chart = (id: number) =>
      frame(
        id,
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdMissing"/>',
      );
    const { notices } = await extractPptx(
      await deckOf([{ tree: Array.from({ length: 12 }, (_, i) => ole(i + 2)).join("") + chart(20) }, { tree: ole(2) }]),
    );
    expect(notices).toEqual([
      { kind: "unsupported-element", slide: 1, what: "ole", outcome: "converted", count: 12 },
      { kind: "unsupported-element", slide: 2, what: "ole", outcome: "converted", count: 1 },
      { kind: "chart-withheld", count: 1 }, // the chart frame has no part: said (07/10/2026)
    ]);
  });

  it("worst case, 21 slides x 500 unsupported frames: the notices stay under 10 KB serialized (was ~800 KB, past Forge's ~240 KiB value limit)", async () => {
    const ole = (id: number) => frame(id, "http://schemas.openxmlformats.org/presentationml/2006/ole", "<p:oleObj/>");
    const tree = Array.from({ length: 500 }, (_, i) => ole(i + 2)).join("");
    const { notices } = await extractPptx(
      await deckOf(Array.from({ length: 21 }, () => ({ tree }))),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const unsupported = notices.filter((n) => n.kind === "unsupported-element");
    expect(unsupported).toHaveLength(21);
    expect(unsupported.every((n) => n.kind === "unsupported-element" && n.count === 500)).toBe(true);
    // Failure value: 10 500 entries, ~800 KB.
    expect(JSON.stringify(notices).length).toBeLessThan(10_000);
  });

  it("an empty slide is a tab holding one empty paragraph", async () => {
    const { doc } = await extractPptx(await deckOf([{ tree: "" }]), undefined, undefined, PAGE_SIZED_LIMITS);
    expect(tabsOf(doc.blocks)).toEqual([
      { title: "Slide 1", slideNumber: 1, blocks: [{ kind: "paragraph", inlines: [] }] },
    ]);
  });

  it("bulleted paragraphs become nested lists, a numbered one keeps its startAt", async () => {
    const body =
      '<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>a</a:t></a:r></a:p>' +
      '<a:p><a:pPr lvl="1"><a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr><a:r><a:t>b</a:t></a:r></a:p>' +
      '<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>c</a:t></a:r></a:p>' +
      "<a:p><a:r><a:t>d</a:t></a:r></a:p>";
    const sp = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`;
    const { doc } = await extractPptx(await deckOf([{ tree: sp }]), undefined, undefined, PAGE_SIZED_LIMITS);
    const p = (t: string): IRBlock => ({ kind: "paragraph", inlines: [{ text: t }] });
    expect(tabsOf(doc.blocks)[0]!.blocks).toEqual([
      {
        kind: "list",
        ordered: false,
        items: [[p("a"), { kind: "list", ordered: true, start: 3, items: [[p("b")]] }], [p("c")]],
      },
      p("d"),
    ]);
  });
});

describe("extractPptx — a diagram slide: drawing, then its text in one collapsed paragraph (decision (5 ter), 25/09/2026)", () => {
  // The slide's shapes and the two-slide deck (`deck`): `pptx-deck.test.support.ts`.
  const png = ONE_PIXEL_PNG;

  it("the tab is the heading then the SVG drawing; ONE diagramText after the tabs (option A): text joined by « · », breaks as spaces, the link kept, no picture", async () => {
    const { doc, images } = await extractPptx(await deck(), undefined, undefined, PAGE_SIZED_LIMITS);
    const [one, two] = tabsOf(doc.blocks);
    // Failure value (until 25/09/2026): heading, image (SVG), paragraphs
    // "Alpha" and "R&D\nteam", "Omega", the table, then the PNG picture.
    // Since 02/10/2026 the declared table follows the drawing as a table block, no longer flattened into the diagram text.
    expect(one!.blocks.map((b) => b.kind)).toEqual(["heading", "image", "table"]);
    expect(diagramsOf(doc.blocks).get("Slide 1")).toEqual([
      { text: "Alpha" },
      { text: " " },
      { text: "R&D" },
      { text: " " },
      { text: "team", href: "https://example.com/x?a=1&b=2" },
      { text: " · " },
      { text: "Omega" },
    ]);
    // Slide 2 has no text: no section for it.
    expect([...diagramsOf(doc.blocks).keys()]).toEqual(["Slide 1"]);
    expect(two!.blocks.map((b) => b.kind)).toEqual(["image"]);
    // The PNG is inside the drawings only, never published as a picture of its own.
    expect(images.length).toBeGreaterThan(0);
    expect(images.every((i) => i.mimeType === "image/svg+xml")).toBe(true);
  });

  /** 29/09/2026: ONE block for the whole deck, after the tabs, one section per diagram slide in order. */
  it("two diagram slides: ONE diagramText after the tabs, a section per slide in slide order", async () => {
    const diagram = (word: string): { tree: string } => ({
      tree: box(3, at(0, 0, 2_000_000, 1_000_000), `<a:p><a:r><a:t>${word}</a:t></a:r></a:p>`) + connector(4),
    });
    const flow = { tree: box(3, at(0, 0, 2_000_000, 1_000_000), "<a:p><a:r><a:t>Plain</a:t></a:r></a:p>") };
    const { doc } = await extractPptx(
      await deckOf([diagram("Alpha"), flow, diagram("Omega")], { sldSz: '<p:sldSz cx="9144000" cy="5143500"/>' }),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    tabsOf(doc.blocks);
    expect(doc.blocks.slice(1)).toEqual([
      {
        kind: "diagramText",
        sections: [
          { label: "Slide 1", slideNumber: 1, inlines: [{ text: "Alpha" }] },
          { label: "Slide 3", slideNumber: 3, inlines: [{ text: "Omega" }] },
        ],
      },
    ]);
  });

  it("no diagram slide with text: no diagramText block at all", async () => {
    const { doc } = await extractPptx(
      await deckOf([{ tree: connector(4) }], { sldSz: '<p:sldSz cx="9144000" cy="5143500"/>' }),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    expect(doc.blocks.filter((b) => b.kind === "diagramText")).toEqual([]);
  });

  it("a picture the drawing leaves out is a dashboard-only notice now (lot A, 29/09/2026) — the tab holds the drawing alone", async () => {
    const pic = (id: number, blip: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill>${blip}<a:stretch/></p:blipFill><p:spPr>${at(7_000_000, 4_000_000, 500_000, 500_000)}</p:spPr></p:pic>`;
    // A group stretched twice as wide around a picture turned by 90°: left
    // out, not approximated (a group that only turns draws its picture).
    const turnedGroup =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>' +
      '<a:xfrm><a:off x="7000000" y="4000000"/><a:ext cx="1000000" cy="500000"/><a:chOff x="7000000" y="4000000"/><a:chExt cx="500000" cy="500000"/></a:xfrm>' +
      `</p:grpSpPr>${pic(21, '<a:blip r:embed="rId2"/>').replace("<a:xfrm>", '<a:xfrm rot="5400000">')}</p:grpSp>`;
    // An EMF (a format no browser draws inside an SVG), a linked picture, a
    // PNG that alone passes the 1 MB cap, then the turned group.
    const big = new Uint8Array(1_000_001);
    big.set(png);
    const { doc, images, notices } = await extractPptx(
      await deckOf(
        [
          {
            tree:
              box(3, at(0, 0, 2_000_000, 1_000_000), "<a:p><a:r><a:t>Alpha</a:t></a:r></a:p>") +
              connector(4) +
              pic(7, '<a:blip r:embed="rId3"/>') +
              pic(8, '<a:blip r:link="rId4"/>') +
              pic(9, '<a:blip r:embed="rId6"/>') +
              turnedGroup,
            rels:
              `<Relationship Id="rId2" Type="${REL}/image" Target="../media/image1.png"/>` +
              `<Relationship Id="rId3" Type="${REL}/image" Target="../media/image2.emf"/>` +
              `<Relationship Id="rId4" Type="${REL}/image" Target="https://example.com/a.png" TargetMode="External"/>` +
              `<Relationship Id="rId6" Type="${REL}/image" Target="../media/image3.png"/>`,
          },
        ],
        {
          sldSz: '<p:sldSz cx="9144000" cy="5143500"/>',
          files: {
            "ppt/media/image1.png": png,
            "ppt/media/image2.emf": new Uint8Array([1, 0, 0, 0, 0x45, 0x4d, 0x46]),
            "ppt/media/image3.png": big,
          },
        },
      ),
    );
    const blocks = tabsOf(doc.blocks)[0]!.blocks;
    // Failure value (until lot A, 29/09/2026): a visible line per picture
    // left out. Now: the drawing alone, four dashboard-only notices.
    expect(blocks).toMatchObject([{ kind: "image" }]);
    expect(blocks).toHaveLength(1); // the diagramText is hoisted after the tabs (option A)
    expect(notices).toEqual([
      { kind: "picture-not-drawn", slide: 1, reason: "format", count: 1 },
      { kind: "picture-not-drawn", slide: 1, reason: "external-link", count: 1 },
      { kind: "picture-not-drawn", slide: 1, reason: "over-cap", count: 1 },
      { kind: "picture-not-drawn", slide: 1, reason: "group-stretched", count: 1 },
    ]);
    expect(diagramsOf(doc.blocks).get("Slide 1")).toMatchObject([{ text: "Alpha" }]);
    // Only the drawing is published: none of the pictures left out.
    expect(images.map((i) => i.mimeType)).toEqual(["image/svg+xml"]);
  });
});

// Review of 25/09/2026, finding 2: the ZIP layer. The 25/09 net
// (package.ts#readXml) only catches the XML PARSER; a corrupted compressed
// stream or a bad CRC rejects from `entry.async(...)` with JSZip's own words,
// which left readPptx as-is. Since the 27/09 evening review (finding 4) the
// net lives AT the zip layer (package.ts#readText/readBytes), so readPptx
// itself answers with our message and a TypeError of OUR reader stays ours —
// and googleDiagramImages keeps its own failure class (network, quota).
/** Re-writes `source` with every entry DEFLATED (how every real producer
 *  writes a pptx; a STORED entry cannot fail without a CRC check JSZip does
 *  not do by default), then flips the first bytes of each named entry's
 *  stream: the zip directory and headers stay intact, only the decompression
 *  of those entries fails. */
const corruptStreams = async (source: Uint8Array, names: readonly string[]): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(source);
  const bytes = new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  const view = new DataView(bytes.buffer, bytes.byteOffset);
  for (const rawName of names) {
    const name = new TextEncoder().encode(rawName);
    // The occurrence preceded by a LOCAL header (PK\x03\x04, then the exact
    // name length at +26) — the name also appears in the central directory.
    let nameAt = -1;
    for (let i = 30; i < bytes.length - name.length; i++) {
      if (!name.every((b, j) => bytes[i + j] === b)) continue;
      const h = i - 30;
      if (view.getUint32(h, true) === 0x04034b50 && view.getUint16(h + 26, true) === name.length) {
        expect(view.getUint16(h + 8, true), "the re-generated entry must be deflated").toBe(8);
        nameAt = i;
        break;
      }
    }
    expect(nameAt, rawName).toBeGreaterThan(0);
    const extralen = view.getUint16(nameAt - 30 + 28, true);
    const dataStart = nameAt + name.length + extralen;
    for (let i = 0; i < 8; i++) bytes[dataStart + i] = (bytes[dataStart + i] ?? 0) ^ 0xff;
  }
  return bytes;
};

describe("a pptx whose zip stream is corrupted fails with our own message, never JSZip's", () => {
  const corruptStream = (): Promise<Uint8Array> =>
    corruptStreams(fixture("pptx/03-Presentation-Resultats-T1.pptx"), ["ppt/slides/slide1.xml"]);

  it("readPptx ITSELF rejects with our message — the net lives at the zip layer, not around our reader's own bugs (review 27/09 evening, finding 4)", async () => {
    const { readPptx } = await import("./pptx-reader");
    await expect(readPptx(await corruptStream())).rejects.toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid PowerPoint presentation",
    });
  });

  it("extractPptx rejects with ReaderError EXTRACT_FAILED and the PowerPoint message", async () => {
    // Failure value really reachable: without the net, this rejects with
    // JSZip/pako's own words, not an ReaderError.
    await expect(extractPptx(await corruptStream(), undefined, undefined, PAGE_SIZED_LIMITS)).rejects.toMatchObject({
      code: "EXTRACT_FAILED",
      message: "the file is damaged or is not a valid PowerPoint presentation",
    });
  });
});

// Tested through the pure function with a small cap; the wired cap is asserted apart.
describe("createTabCapNotice — the 21-tab display cap of a presentation page", () => {
  it("under or at the cap: no notice", () => {
    expect(createTabCapNotice(2, 3)).toBeUndefined();
    expect(createTabCapNotice(3, 3)).toBeUndefined();
  });

  it("over the cap: a notice naming the slides left out (dashboard-only since lot A, 29/09/2026)", () => {
    expect(createTabCapNotice(5, 3)).toEqual({ kind: "tab-cap", from: 4, to: 5 });
  });

  it("one slide over: the notice still reads naturally", () => {
    expect(createTabCapNotice(4, 3)).toEqual({ kind: "tab-cap", from: 4, to: 4 });
  });

  it("the wired cap is the decided 21 (27/09/2026)", () => {
    expect(MAX_PRESENTATION_TABS).toBe(21);
  });
});

// ---------------------------------------------------------------------------
// 🔴 Review of the PowerPoint path, 27/09 — the docx hardening's method
// (docx-reader.test.ts « review 27/09 night »): hostile decks built in memory.

describe("arbitrage 28/09 (7a, révisé) — a slide over MAX_SLIDE_ELEMENTS carries NO TAB AT ALL (A7 of lot A, 29/09/2026)", () => {
  it("502 elements: no tab, a dashboard notice instead; the OTHER slide keeps its real number, the document still publishes", async () => {
    // The unit is withheld WHOLE, the document never refused, and no tab at all, not
    // even one bearing the notice. Failure values: e0…e499 published (a truncating rule),
    // a tab holding a line, or no tab at all for the deck (a refusal).
    const tree = Array.from({ length: 502 }, (_, i) => text(i + 2, `e${i}`)).join("");
    const { doc, notices } = await extractPptx(
      await deckOf([{ tree }, { tree: text(2, "next slide") }]),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const tabs = tabsOf(doc.blocks);
    expect(tabs).toHaveLength(1);
    // The surviving slide keeps its REAL number — « Slide 2 », never
    // re-ranked to « Slide 1 » because the first slide vanished.
    expect(tabs[0]!.title).toBe("Slide 2");
    expect(textOf(tabs[0]!.blocks)).toBe("next slide");
    expect(notices).toContainEqual({ kind: "slide-withheld", slide: 1, elementCount: 502 });
  });
});

describe("🔴 review pptx 27/09, finding 8 — a slide past the display cap never pays its media", () => {
  const pic = (n: number): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="2" name="p${n}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill><p:spPr/></p:pic>';
  const brokenSlide = (n: number): MadeSlide => ({
    tree: pic(n),
    rels: `<Relationship Id="rId2" Type="${REL}/image" Target="../media/broken${n}.png"/>`,
  });
  const png = encodePng(1, 1, "rgb-24", [1, 2, 3])!;
  const files = (numbers: readonly number[]): Record<string, Uint8Array> =>
    Object.fromEntries(numbers.map((n) => [`ppt/media/broken${n}.png`, png]));

  it("a 25-slide deck whose slides 22 to 25 hold broken image streams is published: 21 tabs and the notice", async () => {
    const beyond = [22, 23, 24, 25];
    const slides: MadeSlide[] = [
      ...Array.from({ length: 21 }, (_, i) => ({ tree: text(2, `slide ${i + 1}`) })),
      ...beyond.map(brokenSlide),
    ];
    const bytes = await corruptStreams(
      await deckOf(slides, { files: files(beyond) }),
      beyond.map((n) => `ppt/media/broken${n}.png`),
    );
    // Failure value: EXTRACT_FAILED — the whole deck was read (media
    // inflated, charts analysed) BEFORE the cut at 21.
    const { doc, images, notices } = await extractPptx(bytes, undefined, undefined, PAGE_SIZED_LIMITS);
    expect(tabsOf(doc.blocks)).toHaveLength(21);
    expect(notices).toContainEqual({ kind: "tab-cap", from: 22, to: 25 });
    expect(images).toEqual([]);
  });

  it("control: the same broken stream on a PUBLISHED slide is still READ — its damage is a dashboard notice on the right slide", async () => {
    // Until the arbitrage of 28/09/2026 the proof was the deck's refusal;
    // a damaged auxiliary part now degrades instead, so the proof that a
    // shown slide's media stays read is its notice — a slide past the cap
    // leaves none (the test above: no image read at all).
    const slides: MadeSlide[] = [{ tree: text(2, "slide 1") }, brokenSlide(2), { tree: text(2, "slide 3") }];
    const bytes = await corruptStreams(await deckOf(slides, { files: files([2]) }), ["ppt/media/broken2.png"]);
    const { doc, images, notices } = await extractPptx(bytes, undefined, undefined, PAGE_SIZED_LIMITS);
    expect(textOf(tabsOf(doc.blocks)[1]!.blocks)).toBe("");
    expect(notices).toContainEqual({
      kind: "unsupported-element",
      slide: 2,
      what: "damaged-part",
      outcome: "published",
      count: 1,
    });
    expect(images).toEqual([]);
  });
});

describe("🔴 review pptx 27/09, finding 4 — the published SVG never contains Infinity", () => {
  // Since 09/10/2026 a TOP-LEVEL shape wholly outside the slide is no longer published (rule « Rien d'invisible », decision
  // (e)); members of a group are never judged by position (child coordinates), so the overflow guard is exercised there.
  it('a grouped shape at off x="1e308" with ext cx="1e308" is left out of the drawing, never drawn at Infinity', async () => {
    const huge =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="g5"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="4000000" y="2000000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="4000000" y="2000000"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="s3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="1e308" y="0"/><a:ext cx="1e308" cy="1000000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>' +
      "<p:txBody><a:bodyPr/><a:p><a:r><a:t>far</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>";
    const connector =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="c4"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="0"/></a:xfrm>' +
      '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>';
    const { doc, images } = await extractPptx(
      await deckOf([{ tree: huge + connector }], { sldSz: '<p:sldSz cx="9144000" cy="5143500"/>' }),
    );
    const drawing = images.find((i) => i.mimeType === "image/svg+xml");
    expect(drawing).toBeDefined();
    const svg = new TextDecoder().decode(drawing!.bytes);
    // Failure value: width="Infinity" on the shape's outline (1e308 + 1e308
    // overflows past Number.MAX_VALUE).
    expect(svg).not.toContain("Infinity");
    expect(svg).not.toContain("NaN");
    // The shape's words are not lost: the collapsed text keeps them.
    expect(JSON.stringify(doc.blocks)).toContain("far");
  });
});

// Exhaustive audit of 07/10/2026 (P1-F1.1) and the maintainer (07/10: same situation, same notice whatever the format): a chart the reader
// leaves out — damaged part or over a cap — is « chart-withheld », as in Word and Excel, not « unsupported-element / converted ».
describe("extractPptx — a chart left out says chart-withheld, like docx and xlsx", () => {
  it("🔴 a damaged chart part gives one chart-withheld notice", async () => {
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const path = "ppt/charts/chart1.xml";
    const xml = await zip.file(path)!.async("string");
    const offset = xml.indexOf("?>") + 1;
    zip.file(path, xml.slice(0, offset) + "<p:corruptedUnclosed" + xml.slice(offset));
    const { notices } = await extractPptx(
      await zip.generateAsync({ type: "uint8array" }),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );

    expect(notices).toContainEqual({ kind: "chart-withheld", count: 1 });
    expect(notices.filter((n) => n.kind === "unsupported-element" && n.what === "chart")).toEqual([]);
  });
});

describe("extractPptx — a chart with no series says nothing, like docx and xlsx", () => {
  it("🔴 no notice for the chart frame", async () => {
    const zip = await JSZip.loadAsync(fixture("pptx/03-Presentation-Resultats-T1.pptx"));
    const path = "ppt/charts/chart1.xml";
    const xml = await zip.file(path)!.async("string");
    zip.file(path, xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ""));
    const { notices } = await extractPptx(
      await zip.generateAsync({ type: "uint8array" }),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );

    expect(xml).toMatch(/<c:ser>/);
    expect(
      notices.filter((n) => n.kind === "chart-withheld" || (n.kind === "unsupported-element" && n.what === "chart")),
    ).toEqual([]);
  });
});
