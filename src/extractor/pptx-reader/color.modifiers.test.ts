/**
 * Colour modifiers of a SmartArt's saved drawing (`dsp:drawing`), resolved only when the context says so (`modifiers: true`, set by
 * `extractor/smartart-drawing.ts`); a slide's own colours keep decision 12 of 25/09/2026 (a modifier = unknown, `pptx-reader.test.ts`).
 * Sources: ECMA-376 Part 1 §20.1.2.3.34 `tint`, §20.1.2.3.31 `shade` (the share of the colour kept / of white added, applied on the
 * colour in LINEAR light), §20.1.2.3.20 `lumMod`, §20.1.2.3.21 `lumOff`, §20.1.2.3.30 `satMod` (HSL, `val` in 1/1000 of a percent),
 * §20.1.2.3.1 `alpha` (opacity: no single colour, stays unknown). Expected values computed on 20/10/2026 with an independent script
 * (Python `colorsys`, the sRGB transfer function), and the tint one checked against the pixel Office's own PDF paints for the circles of
 * a real three-slide SmartArt deck kept in the publisher's private repository (#CFD5EA at the centre of a `tint 40000` accent1 circle).
 * Run: npx vitest run --project=main src/extractor/pptx-reader/color.modifiers.test.ts
 */
import { describe, expect, it } from "vitest";

import type { XNode } from "../ooxml/xml";
import { parseXml } from "../ooxml/xml";
import { type ColorContext, exactColorOf } from "./color";

const colour = (inner: string): XNode =>
  parseXml(
    `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="accent1">${inner}</a:schemeClr>`,
  )!;

const context = (modifiers: boolean): ColorContext => ({
  clrMap: {},
  clrScheme: { accent1: "#4472C4" },
  fillStyleLst: [],
  ...(modifiers ? { modifiers: true } : {}),
});

describe("colour modifiers with `modifiers: true`", () => {
  it.each([
    ['<a:tint val="40000"/>', "#CFD5EA"],
    ['<a:shade val="50000"/>', "#2F528F"],
    ['<a:lumMod val="75000"/>', "#2F5597"],
    ['<a:lumMod val="40000"/><a:lumOff val="60000"/>', "#B4C7E7"],
    ['<a:satMod val="50000"/>', "#647BA4"],
    ['<a:tint val="40000"/><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/>', "#CFD5EA"],
  ])("%s gives %s", (inner, expected) => {
    expect(exactColorOf(colour(inner), context(true))).toBe(expected);
  });

  it("alpha (an opacity) and a modifier not read stay unknown", () => {
    expect(exactColorOf(colour('<a:alpha val="90000"/>'), context(true))).toBeUndefined();
    expect(exactColorOf(colour('<a:hueMod val="90000"/>'), context(true))).toBeUndefined();
    expect(exactColorOf(colour('<a:tint val="40000"/><a:alpha val="50000"/>'), context(true))).toBeUndefined();
  });

  it("without `modifiers` the same colours are unknown (decision 12 of 25/09/2026 holds for a slide)", () => {
    expect(exactColorOf(colour('<a:tint val="40000"/>'), context(false))).toBeUndefined();
    expect(exactColorOf(colour('<a:lumMod val="75000"/>'), context(false))).toBeUndefined();
  });
});
