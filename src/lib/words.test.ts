/** `wordsOf`: the one word of the app. A script written without spaces (Han, Kana, Hangul) and Latin letters or digits never share a
 *  token: Word puts the automatic space between them in its PDF, and the Confluence search splits them too (CONCEPTION-MESURE-MOTS-PDF [C1]). */
import { describe, expect, it } from "vitest";
import { wordsOf } from "./words";

describe("wordsOf — cut at the boundary between an unspaced script and Latin or digits", () => {
  it("cuts « 国際会議SC23における » into three tokens", () => {
    expect(wordsOf("国際会議SC23における")).toEqual(["国際会議", "sc23", "における"]);
  });
  it("cuts Hangul and Latin, in both directions", () => {
    expect(wordsOf("한국어ABC한국어")).toEqual(["한국어", "abc", "한국어"]);
  });
  it("leaves a Latin-only or Han-only run whole", () => {
    expect(wordsOf("L'ÉTÉ 2026 国際会議")).toEqual(["l", "été", "2026", "国際会議"]);
  });
  it("never makes a word of the middle dot, which shares the script extensions of Han and Kana", () => {
    expect(wordsOf("a · 東 · b")).toEqual(["a", "東", "b"]);
  });
  it("keeps Hiragana, Katakana and Han together (a Japanese word is one run of those scripts)", () => {
    expect(wordsOf("東京タワーです")).toEqual(["東京タワーです"]);
  });
});
