/**
 * A SmartArt is published as a DRAWING (an SVG picture) traced from the `dsp:drawing` Office saved beside its data model
 * (19/10/2026). Small packages built here
 * (`smartart-drawing.test.support.ts`) exercise one property each; the real files are in `smartart-drawing.fixtures.test.ts`.
 * Decisions of the autonomous session: C1 one SVG per slide, the `dsp:sp` added WITHOUT scaling and OUTSIDE the slide's
 * element count; Word: the SVG is born in `docx-reader/drawing-objects.ts`; C5 zero offsets are neutral; C6 text turned by
 * any angle is drawn (§10: `txXfrm rot` and `vert` ADD to the shape's `rot`); C7 no drawing,
 * no picture. C2 (`alt`) is OUT of this lot.
 * Run: npx vitest run --project=main src/extractor/smartart-drawing.test.ts
 */
import { describe, expect, it } from "vitest";

import { extractDocx } from "./docx";
import { extractPptx } from "./pptx";
import { MAX_SVG_BYTES } from "./slide-svg";
import {
  docxWith,
  dspShape,
  EMU_PER_PT,
  fillWith,
  outlineRects,
  pptxWith,
  slideBox,
  slideConnector,
  svgCompact,
  svgsOf,
  ZERO_OFFSETS_FILL,
  type SmartArtSpec,
  type SlideSpec,
} from "./smartart-drawing.test.support";
import { MAX_SLIDE_ELEMENTS, PAGE_SIZED_LIMITS } from "./page-caps.test.support";

type Host = "docx" | "pptx";

/** The SVGs a package with these SmartArts publishes, whichever the host. The pptx puts each SmartArt on a slide of its own. */
const publish = async (host: Host, specs: SmartArtSpec[]): Promise<string[]> =>
  host === "docx"
    ? svgsOf(await extractDocx(await docxWith(specs), PAGE_SIZED_LIMITS))
    : svgsOf(
        await extractPptx(
          await pptxWith(specs.map((s): SlideSpec => ({ smartArts: [s] }))),
          undefined,
          undefined,
          PAGE_SIZED_LIMITS,
        ),
      );

const one = (
  shape: ReturnType<typeof dspShape>[],
  model: string[],
  extra: Partial<SmartArtSpec> = {},
): SmartArtSpec => ({ n: 1, model, shapes: shape, ...extra });

describe.each<Host>(["docx", "pptx"])("SmartArt drawing, %s host", (host) => {
  // [C5], then 20/10/2026: the zero offsets are neutral and the modifiers ECMA-376 documents (tint, shade, lumMod, lumOff, satMod) are
  // resolved (`pptx-reader/color.modifiers.test.ts`); alpha (an opacity) leaves the colour unknown (black on white).
  it("a zero-offset colour is EXACT (the box is filled), lumMod 75 % of FF0000 is BF0000, alpha makes it unknown (black on white)", async () => {
    const shape = (fill: string) => [dspShape({ geom: "rect", text: ["Tinted"], fill, textColor: "FFFFFF" })];
    const [exact] = await publish(host, [one(shape(ZERO_OFFSETS_FILL), ["Tinted"])]);
    const [alpha] = await publish(host, [one(shape(fillWith('<a:alpha val="90000"/>')), ["Tinted"])]);
    const [lum] = await publish(host, [one(shape(fillWith('<a:lumMod val="75000"/>')), ["Tinted"])]);
    // Failure value: no SVG at all (`undefined`) today.
    expect(exact, "a picture is published").toBeDefined();
    expect(exact).toMatch(/fill="#FF0000"/);
    expect(alpha).toBeDefined();
    expect(alpha).not.toMatch(/#FF0000/i);
    expect(alpha).not.toMatch(/fill="#FFFFFF"/); // the white text is not kept on a box that is not drawn filled
    // Failure value: an outline only, no `fill="#BF0000"` (lumMod unknown).
    expect(lum).toMatch(/fill="#BF0000"/);
    expect(lum).toMatch(/fill="#FFFFFF"/);
    // The words are drawn either way.
    expect(svgCompact(alpha!)).toBe("Tinted");
  });

  // [MS-ODRAWXML] dsp:txXfrm: the text box; Word's own PDF of smart-art.docx shows the hexagons' text horizontal.
  it("a hexagon turned 90 degrees whose txXfrm turns it back (-90): outline drawn, text horizontal", async () => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({
            geom: "hexagon",
            box: [2_414_159, 77_407, 1_186_160, 1_031_959],
            rot: 5_400_000,
            txRot: -5_400_000,
            text: ["Point 1"],
          }),
        ],
        ["Point 1"],
      ),
    ]);
    expect(svg, "a picture is published").toBeDefined();
    expect(svg).toMatch(/<(path|polygon)\b/); // the hexagon's outline: not a rect
    expect(svgCompact(svg!)).toBe("Point1");
    expect(svg).not.toMatch(/<text\b[^>]*transform=/); // the net angle is 0
  });

  // C6 (§9, from smart-art.pdf): « Bar 1 » is VERTICAL in Word's own export: the rotations ADD (270 + 0), they do not cancel.
  it("a rect turned 270 degrees with a txXfrm without rot: the text is drawn turned by 270 around the box centre", async () => {
    const box: [number, number, number, number] = [-1_082_958, 1_642_494, 2_496_312, 264_340]; // docx drawing3, « Bar 1 »
    const [svg] = await publish(host, [
      one([dspShape({ geom: "rect", box, rot: 16_200_000, txXfrm: true, text: ["Bar 1"] })], ["Bar 1"]),
    ]);
    expect(svg, "a picture is published").toBeDefined();
    expect(outlineRects(svg ?? "").length).toBeGreaterThan(0); // the outline is drawn
    const turn = /<text\b[^>]*transform="rotate\((-?[\d.]+) (-?[\d.]+) (-?[\d.]+)\)"[^>]*>Bar 1</.exec(svg ?? "");
    expect(turn).not.toBeNull();
    expect([270, -90]).toContain(Number(turn![1]));
    // The pivot is the shape's centre, in points. The docx canvas has its origin at the frame; the pptx frame sits at (1 000 000, 1 000 000) EMU.
    const origin = host === "docx" ? 0 : 1_000_000;
    expect(Number(turn![2])).toBeCloseTo((origin + box[0] + box[2] / 2) / EMU_PER_PT, 0);
    expect(Number(turn![3])).toBeCloseTo((origin + box[1] + box[3] / 2) / EMU_PER_PT, 0);
  });

  // [C8], 20/10/2026: ONE mechanism for every angle. Outline: the frame NOT turned + `rotate(θ cx cy)` (SVG 1.1 §7.6)
  // about the frame's centre; text: laid flat in the `txXfrm` box, `rotate(θ frame centre) rotate(τ box centre)` (τ = txXfrm rot + vert).
  // Pivots below are points on the page: the docx canvas has its origin at the frame; the pptx frame sits at (1 000 000, 1 000 000) EMU.
  const origin = host === "docx" ? 0 : 1_000_000;
  const pt = (emu: number): number => (origin + emu) / EMU_PER_PT;
  const round2 = (v: number): number => Math.round(v * 100) / 100;
  /** The `rotate(...)` arguments of the `transform` of the `<tag>` holding `word` (empty: no transform). */
  const rotates = (tag: string, svg: string, word: string): string[] => {
    const element = new RegExp(`<${tag}\\b[^>]*>${word}<`).exec(svg);
    const transform = /transform="([^"]*)"/.exec(element?.[0] ?? "");
    return transform === null ? [] : [...transform[1]!.matchAll(/rotate\(([^)]*)\)/g)].map((r) => r[1]!);
  };

  it("a rect turned 30 degrees (txXfrm without rot): the outline is the frame, not turned, under rotate(30 centre); the text carries the same rotate", async () => {
    const [svg] = await publish(host, [
      one([dspShape({ geom: "rect", text: ["Slanted"], rot: 1_800_000, txXfrm: true })], ["Slanted"]),
    ]);
    // Failure value: « » (a shape turned by another angle than a quarter is left out whole).
    expect(svgCompact(svg ?? "")).toBe("Slanted");
    const [x, y, cx, cy] = [0, 0, 1_500_000, 600_000];
    const centre = `${round2(pt(x + cx / 2))} ${round2(pt(y + cy / 2))}`;
    const rect = /<rect\b[^>]*transform="rotate\(30 ([^)]*)\)"/.exec(svg ?? "");
    expect(rect, "the rect carries rotate(30 ...)").not.toBeNull();
    expect(rect![1]).toBe(centre);
    const r = outlineRects(svg ?? "")[0]!;
    expect([r.x, r.y, r.width, r.height].map(round2)).toEqual(
      [pt(x), pt(y), cx / EMU_PER_PT, cy / EMU_PER_PT].map(round2),
    );
    expect(rotates("text", svg ?? "", "Slanted")).toEqual([`30 ${centre}`]);
  });

  it("a rect turned 30 degrees whose txXfrm turns back by -30: the outline is slanted, the text is upright", async () => {
    const [svg] = await publish(host, [
      one([dspShape({ geom: "rect", text: ["Upright"], rot: 1_800_000, txRot: -1_800_000 })], ["Upright"]),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Upright"); // failure value: no word
    expect(svg).toMatch(/<rect\b[^>]*transform="rotate\(30 /);
    // The two rotations cancel and the two centres coincide: no transform at all (the same bytes as the hexagons of smart-art.docx).
    expect(rotates("text", svg ?? "", "Upright")).toEqual([]);
  });

  it("a rotation about the FRAME's centre carries a text box that is not centred on it (90 degrees, box at 25 % / 50 %)", async () => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({
            geom: "rect",
            box: [0, 0, 1_000_000, 1_000_000],
            rot: 5_400_000,
            txBox: [150_000, 400_000, 200_000, 200_000],
            text: ["Offset"],
          }),
        ],
        ["Offset"],
      ),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Offset");
    // Frame centre (500 000, 500 000), box centre (250 000, 500 000). The text turns by 90 about the FRAME's centre. The previous rule
    // turned it about the box centre: failure value `rotate(90 <box centre>)`.
    const fr = round2(pt(500_000));
    expect(rotates("text", svg ?? "", "Offset")).toEqual([`90 ${fr} ${fr}`]);
    // The text is laid flat at the box: x is the box's left edge plus the inset (83 820 EMU).
    const x = Number(/<text\b[^>]*\bx="(-?[\d.]+)"/.exec(svg ?? "")![1]);
    expect(x).toBeCloseTo(pt(150_000 + 83_820), 1);
  });

  // Corpus 212-4416657.docx: `rect` turned 270 degrees, `a:bodyPr vert="vert"` (+90): horizontal in Word's own PDF.
  it('vert="vert" adds 90 degrees: a rect turned 270 degrees is drawn horizontal (no transform on the text)', async () => {
    const [svg] = await publish(host, [
      one([dspShape({ geom: "rect", rot: 16_200_000, txXfrm: true, vert: "vert", text: ["Flat"] })], ["Flat"]),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Flat"); // failure value: no word (every `vert` is a turned text)
    expect(rotates("text", svg ?? "", "Flat")).toEqual([]);
    expect(svg).toMatch(/<rect\b[^>]*transform="rotate\(270 /);
  });

  it('vert="vert270" adds 270 degrees', async () => {
    const [svg] = await publish(host, [
      one([dspShape({ geom: "rect", txXfrm: true, vert: "vert270", text: ["Up"] })], ["Up"]),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Up");
    const turns = rotates("text", svg ?? "", "Up");
    expect(turns).toHaveLength(1);
    expect(turns[0]).toMatch(/^270 /);
  });

  // The rule is unknown for these (0 occurrences in 355 documents): the outline is drawn, the words are not (counted `smartart-not-drawn`).
  it.each<[string, Partial<Parameters<typeof dspShape>[0]>]>([
    ["stacked glyphs (vert=eaVert)", { vert: "eaVert" }],
    ["upright=1", { upright: true }],
    ["a:bodyPr rot with txXfrm rot", { bodyRot: 900_000, txRot: 900_000 }],
    ["flipV", { flipV: true }],
  ])("%s: the outline is drawn, the words are left out", async (_name, extra) => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({ geom: "rect", text: ["Lost"], txXfrm: true, ...extra }),
          dspShape({ geom: "rect", text: ["Seen"], box: [0, 900_000, 1_500_000, 600_000] }),
        ],
        ["Lost", "Seen"],
      ),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Seen");
    expect(outlineRects(svg ?? "").length).toBe(2);
  });

  // Review of 35a2d1f3, defect 1: `vert` turns the glyphs but does not swap the written width and height of the text box (a `txXfrm rot` of
  // the hexagons does: Word swapped them already). The lines run along the box's HEIGHT: laid in the box with the sizes swapped, same centre.
  it('vert="vert" lays the lines in the text box with its width and height swapped (212-4416657: 55 x 28.8 pt, wrapped at 28.8)', async () => {
    const box: [number, number, number, number] = [1_000_000, 1_000_000, 699_952, 365_561];
    const [svg] = await publish(host, [
      one(
        [dspShape({ geom: "rect", box, rot: 16_200_000, txXfrm: true, vert: "vert", text: ["to be or not to be"] })],
        ["to be or not to be"],
      ),
    ]);
    const xs = [...(svg ?? "").matchAll(/<text\b[^>]*\bx="(-?[\d.]+)"/g)].map((m) => Number(m[1]));
    // Swapped box: 28.8 wide, 55 tall, centred on (27.56, 14.4): left edge 13.16 plus the 6.6 pt inset. Failure value: 6.6 + origin (the written width).
    expect(xs.length).toBeGreaterThan(1);
    for (const x of xs) expect(x).toBeCloseTo(pt(1_000_000 + 699_952 / 2 - 365_561 / 2 + 83_820), 1);
    expect(svgCompact(svg ?? "")).toBe("tobeornottobe");
  });

  // Review of 35a2d1f3, defect 2: the bounds of a shape turned by a non-quarter angle are no background for a text laid flat.
  it("a white text on a shape without fill is black when it lies in the bounds of a slanted filled shape but not on it", async () => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({ geom: "rect", box: [0, 0, 3_000_000, 600_000], rot: 1_800_000 }),
          dspShape({
            geom: "rect",
            box: [200_000, 900_000, 800_000, 300_000],
            fill: "<a:noFill/>",
            text: ["Ghost"],
            textColor: "FFFFFF",
          }),
        ],
        ["Ghost"],
      ),
    ]);
    // Failure value: fill="#FFFFFF" (white on a page that is white there).
    expect(svg).toMatch(/<text\b[^>]*>Ghost</);
    expect(svg).not.toMatch(/<text\b[^>]*fill="#FFFFFF"[^>]*>Ghost</);
  });

  it("a turned text is deterministic: two runs, the same bytes", async () => {
    const spec = one(
      [dspShape({ geom: "rect", text: ["Slanted"], rot: 1_800_000, txBox: [10_000, 10_000, 1_000_000, 500_000] })],
      ["Slanted"],
    );
    const first = await publish(host, [spec]);
    expect(first).toHaveLength(1);
    expect(await publish(host, [spec])).toEqual(first);
  });

  // §7 test 7: a geometry the drawer cannot trace. Without text: nothing is lost, the picture still stands.
  it("shapes without text on a geometry that is not traced (star5, custGeom) cost nothing: the rest is drawn", async () => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({ geom: "star5", box: [0, 0, 500_000, 500_000] }),
          dspShape({ geom: "custGeom", box: [600_000, 0, 500_000, 500_000] }),
          dspShape({ geom: "roundRect", text: ["Kept"], box: [0, 700_000, 1_500_000, 600_000] }),
        ],
        ["Kept"],
      ),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Kept");
  });

  it("a shape with text on a geometry that is not traced leaves its words out of the picture (they are counted, see doc-processor.smartart-drawing.test.ts)", async () => {
    const [svg] = await publish(host, [
      one(
        [
          dspShape({ geom: "star5", text: ["Lost"], box: [0, 0, 1_500_000, 600_000] }),
          dspShape({ geom: "rect", text: ["Seen"], box: [0, 900_000, 1_500_000, 600_000] }),
        ],
        ["Lost", "Seen"],
      ),
    ]);
    expect(svgCompact(svg ?? "")).toBe("Seen");
  });

  // [C1]/[C7]: only a drawing that is there and holds shapes gives a picture; nothing else is invented, nothing refuses the document.
  it("falls back to NO picture when the drawing is absent, unlinked, damaged or empty; the document is extracted all the same", async () => {
    const good = (n: number): SmartArtSpec => ({
      n,
      model: ["Drawn"],
      shapes: [dspShape({ geom: "rect", text: ["Drawn"] })],
    });
    const specs: SmartArtSpec[] = [
      good(1),
      { n: 2, model: ["Nopart"], shapes: [dspShape({ text: ["Nopart"] })], state: "part-absent" },
      { n: 3, model: ["Noext"], shapes: [dspShape({ text: ["Noext"] })], state: "no-extension" },
      { n: 4, model: ["Damaged"], shapes: [dspShape({ text: ["Damaged"] })], state: "damaged" },
      { n: 5, model: ["Emptydrawing"], shapes: [] },
      { n: 6, model: ["Nodrawing"] },
    ];
    const svgs = await publish(host, specs);
    // Failure value: 0 (nothing is drawn today).
    expect(svgs).toHaveLength(1);
    expect(svgCompact(svgs[0]!)).toBe("Drawn");
  });

  // New text of the drawing is drawn, never counted lost.
  it("a drawing shape whose text the data model does not hold is drawn", async () => {
    const [svg] = await publish(host, [
      one(
        [dspShape({ text: ["Modelled"] }), dspShape({ text: ["Newtext"], box: [0, 900_000, 1_500_000, 600_000] })],
        ["Modelled"],
      ),
    ]);
    expect(svgCompact(svg ?? "")).toBe("ModelledNewtext");
  });

  it("two runs give the same SVG, byte for byte", async () => {
    const spec = one(
      [
        dspShape({
          geom: "hexagon",
          rot: 5_400_000,
          txRot: -5_400_000,
          text: ["Same"],
          fill: ZERO_OFFSETS_FILL,
          textColor: "FFFFFF",
        }),
      ],
      ["Same"],
    );
    const first = await publish(host, [spec]);
    const second = await publish(host, [spec]);
    expect(first).toHaveLength(1);
    expect(second).toEqual(first);
  });
});

describe("SmartArt pairing: by the model's dsp:dataModelExt relId, never by part number", () => {
  // Frame A names data2.xml, whose extension names drawing2.xml; frame B names data1.xml and drawing1.xml.
  // A pairing by part number would give frame A drawing1.
  const swapped = (): SmartArtSpec[] => [
    { n: 2, model: ["Alpha"], shapes: [dspShape({ text: ["Alpha"] })] },
    { n: 1, model: ["Beta"], shapes: [dspShape({ text: ["Beta"] })] },
  ];
  it("docx: each frame gets ITS drawing", async () => {
    const svgs = svgsOf(await extractDocx(await docxWith(swapped()), PAGE_SIZED_LIMITS));
    expect(svgs.map(svgCompact)).toEqual(["Alpha", "Beta"]);
  });
  it("pptx: each frame gets ITS drawing", async () => {
    const svgs = svgsOf(
      await extractPptx(
        await pptxWith([
          {
            smartArts: [
              { ...swapped()[0]!, frame: [1_000_000, 500_000, 5_000_000, 2_500_000] },
              { ...swapped()[1]!, frame: [1_000_000, 3_500_000, 5_000_000, 2_500_000] },
            ],
          },
        ]),
        undefined,
        undefined,
        PAGE_SIZED_LIMITS,
      ),
    );
    expect(svgs.map(svgCompact)).toEqual(["Alpha", "Beta"]);
  });
});

describe("pptx: SmartArt on a slide", () => {
  // [C1] (§9): the dsp:sp join the slide's shapes BEFORE renderSlideSvg: one SVG per slide, translated to the frame's corner,
  // scale 1 (the drawing declares no extent), never ratioed to the frame.
  it("a diagram slide: ONE picture holds the slide's shapes and the SmartArt's, translated to the frame corner without scaling", async () => {
    const frame: [number, number, number, number] = [1_000_000, 500_000, 5_000_000, 3_000_000]; // deliberately not the drawing's natural extent
    const drawn = dspShape({ geom: "rect", box: [200_000, 300_000, 1_500_000, 600_000], text: ["Inside"] });
    const result = await extractPptx(
      await pptxWith([
        {
          before: slideBox(2, "OnSlide") + slideConnector(3),
          smartArts: [{ n: 1, model: ["Inside"], shapes: [drawn], frame }],
        },
      ]),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const svgs = svgsOf(result);
    expect(svgs).toHaveLength(1);
    expect(svgCompact(svgs[0]!)).toContain("OnSlide");
    expect(svgCompact(svgs[0]!)).toContain("Inside");
    const rect = outlineRects(svgs[0]!).find((r) => Math.abs(r.width - 1_500_000 / EMU_PER_PT) < 0.05);
    expect(rect).toBeDefined();
    expect(rect!.x).toBeCloseTo((frame[0] + 200_000) / EMU_PER_PT, 1);
    expect(rect!.y).toBeCloseTo((frame[1] + 300_000) / EMU_PER_PT, 1);
    expect(rect!.height).toBeCloseTo(600_000 / EMU_PER_PT, 1);
  });

  // §9: the dsp:sp do not count in `elementCount`: 480 shapes + a SmartArt of 30 must not make the slide exceed 500 and be withheld.
  it("a diagram slide of 480 shapes plus a SmartArt of 30 is published, the SmartArt drawn", async () => {
    const boxes = Array.from({ length: 479 }, (_, i) =>
      slideBox(10 + i, `box${i}`, [(i % 20) * 500_000, 1_900_000 + Math.floor(i / 20) * 20_000, 100_000, 10_000]),
    ).join("");
    const thirty = Array.from({ length: 30 }, (_, i) =>
      dspShape({ text: [`part${i}`], box: [(i % 6) * 800_000, Math.floor(i / 6) * 300_000, 700_000, 250_000] }),
    );
    const result = await extractPptx(
      await pptxWith([
        {
          before: slideConnector(3) + boxes,
          smartArts: [{ n: 1, model: thirty.map((_, i) => `part${i}`), shapes: thirty }],
        },
      ]),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    expect((result.notices ?? []).some((n) => n.kind === "slide-withheld")).toBe(false);
    const svgs = svgsOf(result);
    expect(svgs).toHaveLength(1);
    expect(svgCompact(svgs[0]!)).toContain("part29");
  });

  it("the same SmartArt at the same place on two slides is ONE attachment and two picture blocks", async () => {
    const spec = (n: number): SmartArtSpec => ({
      n,
      model: ["Twin"],
      shapes: [dspShape({ text: ["Twin"] })],
      frame: [1_000_000, 1_000_000, 5_000_000, 3_000_000],
    });
    const result = await extractPptx(
      await pptxWith([{ smartArts: [spec(1)] }, { smartArts: [spec(2)] }]),
      undefined,
      undefined,
      PAGE_SIZED_LIMITS,
    );
    const svgs = result.images.filter((i) => i.mimeType === "image/svg+xml");
    expect(svgsOf(result)).toHaveLength(2);
    expect(svgs).toHaveLength(1);
  });
});

describe("SmartArt caps", () => {
  const many = (count: number): SmartArtSpec => ({
    n: 1,
    model: ["Capped"],
    shapes: Array.from({ length: count }, (_, i) =>
      dspShape({
        text: i === 0 ? ["Capped"] : [`shape number ${i} with some more words to weigh`],
        box: [(i % 10) * 500_000, Math.floor(i / 10) * 100_000, 400_000, 90_000],
      }),
    ),
  });
  it("a drawing of exactly MAX_SLIDE_ELEMENTS shapes is drawn, and weighs far less than MAX_SVG_BYTES (measured on the real bytes)", async () => {
    const result = await extractDocx(await docxWith([many(MAX_SLIDE_ELEMENTS)]), PAGE_SIZED_LIMITS);
    const svgs = svgsOf(result);
    expect(svgs).toHaveLength(1);
    expect(new TextEncoder().encode(svgs[0]).length).toBeLessThan(MAX_SVG_BYTES);
  });
  it("the cap is a parameter (ReaderLimits.slideElements): a 5-shape drawing is not drawn under 4, drawn under 5 and by default", async () => {
    const bytes = await docxWith([many(5)]);
    // FAILS WITH (limits ignored): expected [ '<svg …' ] to have a length of +0 but got 1.
    expect(svgsOf(await extractDocx(bytes, { slideElements: 4 }))).toHaveLength(0);
    expect(svgsOf(await extractDocx(bytes, { slideElements: 5 }))).toHaveLength(1);
    expect(svgsOf(await extractDocx(bytes, PAGE_SIZED_LIMITS))).toHaveLength(1);
  });
  it("a drawing of one shape more is not drawn (its words are counted, doc-processor.smartart-drawing.test.ts)", async () => {
    const svgs = svgsOf(
      await extractDocx(
        await docxWith(
          [{ ...many(MAX_SLIDE_ELEMENTS), n: 1 }, many(MAX_SLIDE_ELEMENTS + 1)].map((s, i) => ({ ...s, n: i + 1 })),
        ),
        PAGE_SIZED_LIMITS,
      ),
    );
    // The first is drawn (cap reached exactly), the second is not: failure value 0 and 2.
    expect(svgs).toHaveLength(1);
  });
});
