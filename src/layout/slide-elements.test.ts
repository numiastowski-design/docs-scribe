import { describe, expect, it } from "vitest";

import { centreInside, coversSlide, slideItemFlags } from "./slide-elements";

const SLIDE = { cx: 1_000, cy: 500 };

describe("slide elements — background picture, posed on a picture", () => {
  it("a background covers the slide exactly or more; one unit short on any side is none", () => {
    expect(coversSlide({ x: 0, y: 0, cx: 1_000, cy: 500 }, SLIDE)).toBe(true);
    expect(coversSlide({ x: -10, y: -10, cx: 1_020, cy: 520 }, SLIDE)).toBe(true);
    // Failure values: a tolerance would accept these.
    expect(coversSlide({ x: 1, y: 0, cx: 999, cy: 500 }, SLIDE)).toBe(false);
    expect(coversSlide({ x: 0, y: 1, cx: 1_000, cy: 499 }, SLIDE)).toBe(false);
    expect(coversSlide({ x: 0, y: 0, cx: 999, cy: 500 }, SLIDE)).toBe(false);
    expect(coversSlide({ x: 0, y: 0, cx: 1_000, cy: 499 }, SLIDE)).toBe(false);
  });

  it("a centre lies in a box edges included, a half-unit centre compared exactly", () => {
    const picture = { x: 100, y: 100, cx: 100, cy: 100 };
    // Centre (100.5, 150): inside by half a unit. Centre (99.5, 150): outside
    // by half a unit. Failure value: a centre rounded half up (100) would
    // put the second on the edge, inside.
    expect(centreInside({ x: 100, y: 145, cx: 1, cy: 10 }, picture)).toBe(true);
    expect(centreInside({ x: 99, y: 145, cx: 0, cy: 10 }, picture)).toBe(false);
    expect(centreInside({ x: 99, y: 145, cx: 1, cy: 10 }, picture)).toBe(false);
    // On the right edge exactly: inside.
    expect(centreInside({ x: 190, y: 145, cx: 20, cy: 10 }, picture)).toBe(true);
    // Centre on the picture with 38 % of the box on it (USA 4's "JIRA" on its
    // logo: 42 %): inside. The centre decides, not the share of the box.
    expect(centreInside({ x: 20, y: 130, cx: 260, cy: 40 }, picture)).toBe(true);
  });

  it("flags: a background is not a picture anything is posed on; a picture is not posed on itself; no box, no flag", () => {
    const flags = slideItemFlags(
      [
        { box: { x: 0, y: 0, cx: 1_000, cy: 500 }, picture: true },
        { box: { x: 100, y: 100, cx: 50, cy: 50 }, picture: false },
        { box: { x: 600, y: 300, cx: 100, cy: 100 }, picture: true },
        { box: { x: 630, y: 330, cx: 40, cy: 40 }, picture: false },
        { box: { x: 660, y: 360, cx: 20, cy: 20 }, picture: true },
        { picture: false },
      ],
      SLIDE,
    );
    // Failure values: the text over the background (1) or the small picture
    // (2) flagged as posed; the picture on a picture (4) not flagged.
    expect(flags).toEqual([
      { background: true, posedOnPicture: false },
      { background: false, posedOnPicture: false },
      { background: false, posedOnPicture: false },
      { background: false, posedOnPicture: true },
      { background: false, posedOnPicture: true },
      { background: false, posedOnPicture: false },
    ]);
  });
});
