/** Characterisation of the docx second reading's tree walk (`source-reading-docx.ts#readWalk`), element by element: what each
 *  kind of node contributes to the text, and with which cause (`MissingWordCause`) its words are tagged. Written before the walk
 *  was cut into one function per case; every expectation was seen green on the uncut walk.
 *  NOT here: the page-geometry rules of drawings (`source-reading-hidden.test.ts`), withheld tables (`source-reading-withheld-table*.test.ts`),
 *  the vertical-merge fold (`invisible-rule-vmerge.test.ts`), the real fixtures (`no-word-lost.test.ts`). */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { MAX_TABLE_COLS, PAGE_SIZED_LIMITS } from "../extractor/page-caps.test.support";
import { readNumberedParagraphs, readSource, wordsOf } from "./source-reading-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS = `xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`;
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

interface Parts {
  body: string;
  styles?: string;
  footnotes?: string;
  endnotes?: string;
}

const docxOf = async (parts: Parts): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document ${NS}><w:body>${parts.body}</w:body></w:document>`);
  if (parts.styles !== undefined)
    zip.file("word/styles.xml", `<?xml version="1.0"?><w:styles ${NS}>${parts.styles}</w:styles>`);
  if (parts.footnotes !== undefined)
    zip.file("word/footnotes.xml", `<?xml version="1.0"?><w:footnotes ${NS}>${parts.footnotes}</w:footnotes>`);
  if (parts.endnotes !== undefined)
    zip.file("word/endnotes.xml", `<?xml version="1.0"?><w:endnotes ${NS}>${parts.endnotes}</w:endnotes>`);
  return zip.generateAsync({ type: "uint8array" });
};

/** A run whose text is set off by blanks, so that the words of neighbouring runs stay apart. */
const run = (t: string): string => `<w:r><w:t xml:space="preserve"> ${t} </w:t></w:r>`;
/** A run with nothing around its text, for the tests of the exact text. */
const tight = (t: string): string => `<w:r><w:t>${t}</w:t></w:r>`;
const para = (inner: string): string => `<w:p>${inner}</w:p>`;
const fld = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (t: string): string => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
const hiddenRun = (inner: string): string => `<w:r><w:rPr><w:vanish/></w:rPr>${inner}</w:r>`;
const note = (kind: "footnote" | "endnote", id: number, inner: string): string =>
  `<w:${kind} w:id="${id}">${inner}</w:${kind}>`;

const read = async (parts: Parts) => readSource(await docxOf(parts), undefined, PAGE_SIZED_LIMITS);
/** "word:cause" for every word, in order. */
const pairs = async (parts: Parts): Promise<string[]> => {
  const { text, wordCauses } = await read(parts);
  return wordsOf(text).map((w, i) => `${w}:${wordCauses[i]}`);
};

describe("readSource — skipped elements and plain text", () => {
  // An inline picture separates two words: one `w:p` can hold a text, a `wp:inline` picture, then another text with
  // no space between them. What is not text between two texts is a word boundary.
  it("une image en ligne sépare les mots de part et d'autre, dans le même paragraphe", async () => {
    const picture = `<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/></wp:inline></w:drawing></w:r>`;
    const body = para(tight("Competition") + picture + tight("There"));
    // Failure value: ["competitionthere:text"].
    expect(await pairs({ body })).toEqual(["competition:text", "there:text"]);
  });

  // The bound: a word the document writes in two runs, which Word does all the time, stays one word.
  it("deux runs sans image restent un seul mot", async () => {
    expect(await pairs({ body: para(tight("Competitio") + tight("n")) })).toEqual(["competition:text"]);
  });

  it("drops deletions, moved-from text, deleted text and the fallback of an alternate content", async () => {
    const body = para(
      [
        run("one"),
        `<w:del><w:r><w:t>gone</w:t></w:r></w:del>`,
        `<w:moveFrom><w:r><w:t>moved</w:t></w:r></w:moveFrom>`,
        `<w:r><w:delText>deleted</w:delText></w:r>`,
        `<w:r><w:delInstrText>deletedcode</w:delInstrText></w:r>`,
        `<mc:AlternateContent><mc:Fallback>${run("fallback")}</mc:Fallback></mc:AlternateContent>`,
        run("two"),
      ].join(""),
    );
    expect(await pairs({ body })).toEqual(["one:text", "two:text"]);
  });

  it("reads through wrappers it does not know (hyperlink, smart tag) in place", async () => {
    const body = para(`<w:hyperlink>${run("link")}</w:hyperlink><w:smartTag>${run("tag")}</w:smartTag>`);
    expect(await pairs({ body })).toEqual(["link:text", "tag:text"]);
  });

  it("tab, break and carriage return are one space; a no-break hyphen is a hyphen; a symbol is its code point", async () => {
    const body = para(
      `<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:cr/><w:noBreakHyphen/><w:sym w:char="F0E0"/><w:sym w:char="0"/><w:sym w:char="zz"/><w:sym/></w:r>`,
    );
    expect((await read({ body })).text).toBe("\na b c -\n");
  });

  it("🔴 a w:sym in the Symbol font is read as the character it draws, like the page (26/10/2026)", async () => {
    // FAILS WITH: « \uF071 » — the page publishes θ, the second reading would count it invented.
    const body = para(
      `<w:r><w:t>a</w:t><w:sym w:font="Symbol" w:char="F071"/><w:sym w:font="Wingdings" w:char="F0B7"/></w:r>`,
    );
    expect((await read({ body })).text).toBe("\naθ\uF0B7\n");
  });

  it('🔴 a w:char beyond Unicode (w:char="110000") is skipped, never a thrown error, like the page reader (review of 26/10/2026)', async () => {
    // FAILS WITH: RangeError: Invalid code point 1114112 — the second reading fails on a hostile document the page reads.
    const body = para(`<w:r><w:t>a</w:t><w:sym w:font="Symbol" w:char="110000"/><w:t>b</w:t></w:r>`);
    expect((await read({ body })).text).toBe("\nab\n");
  });

  it("🔴 a w:char in the surrogate range (D800–DFFF) is skipped, never a lone surrogate (counter-review of 26/10/2026)", async () => {
    // FAILS WITH: « a\uD800b » — a lone surrogate read into the text.
    const body = para(`<w:r><w:t>a</w:t><w:sym w:font="Arial" w:char="D800"/><w:t>b</w:t></w:r>`);
    expect((await read({ body })).text).toBe("\nab\n");
  });

  it("a math run is set off by one space on each side, its text read", async () => {
    const body = para(`${tight("x")}<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>${tight("z")}`);
    expect((await read({ body })).text).toBe("\nx y z\n");
  });

  it("a paragraph is framed by two line breaks, in document order", async () => {
    expect((await read({ body: para(tight("a")) + para(tight("b")) })).text).toBe("\na\n\nb\n");
  });
});

describe("readSource — complex and simple fields", () => {
  it("the code of a complex field is not read, its cached result is, with the field-result cause", async () => {
    const body = para(
      [
        run("one"),
        fld("begin"),
        instr(" HYPERLINK x "),
        run("codeword"),
        fld("separate"),
        run("linked"),
        fld("end"),
        run("two"),
      ].join(""),
    );
    expect(await pairs({ body })).toEqual(["one:text", "linked:field-result", "two:text"]);
  });

  it("the cached result of a TOC field is not read, the text after the field is", async () => {
    const body = para(
      [fld("begin"), instr(" TOC \\o "), fld("separate"), run("tocword"), fld("end"), run("after")].join(""),
    );
    expect(await pairs({ body })).toEqual(["after:text"]);
  });

  it("a TOC code written in lower case, preceded by blanks, is still a TOC", async () => {
    const body = para(
      [fld("begin"), instr("  toc \\h "), fld("separate"), run("tocword"), fld("end"), run("after")].join(""),
    );
    expect(await pairs({ body })).toEqual(["after:text"]);
  });

  it("a code that only starts like TOC (TOCA) is an ordinary field", async () => {
    const body = para([fld("begin"), instr(" TOCA "), fld("separate"), run("kept"), fld("end")].join(""));
    expect(await pairs({ body })).toEqual(["kept:field-result"]);
  });

  it("an instruction text outside any field, a separator with no field open and an end with none open change nothing", async () => {
    const body = para([instr(" TOC "), fld("separate"), fld("end"), run("plain")].join(""));
    expect(await pairs({ body })).toEqual(["plain:text"]);
  });

  it("a nested field closes the inner one only: the outer result goes on after it", async () => {
    const body = para(
      [
        fld("begin"),
        instr(" A "),
        fld("separate"),
        run("outer1"),
        fld("begin"),
        instr(" B "),
        fld("separate"),
        run("inner"),
        fld("end"),
        run("outer2"),
        fld("end"),
        run("after"),
      ].join(""),
    );
    expect(await pairs({ body })).toEqual([
      "outer1:field-result",
      "inner:field-result",
      "outer2:field-result",
      "after:text",
    ]);
  });

  it("inside the code of a field, a text is hidden but a tab and a no-break hyphen still count, a symbol does not", async () => {
    const body = para(
      [
        fld("begin"),
        `<w:r><w:tab/><w:noBreakHyphen/><w:sym w:char="0041"/><w:t>codeword</w:t></w:r>`,
        fld("separate"),
        `<w:r><w:sym w:char="0042"/></w:r>`,
        fld("end"),
      ].join(""),
    );
    expect((await read({ body })).text).toBe("\n -B\n");
  });

  it("a simple field reads its content with the field-result cause; a simple TOC field hides it", async () => {
    const body = para(
      `<w:fldSimple w:instr=" PAGE ">${run("pagenum")}</w:fldSimple>${run("between")}<w:fldSimple w:instr=" TOC \\o ">${run("tocsimple")}</w:fldSimple>${run("after")}<w:fldSimple>${run("nocode")}</w:fldSimple>`,
    );
    expect(await pairs({ body })).toEqual([
      "pagenum:field-result",
      "between:text",
      "after:text",
      "nocode:field-result",
    ]);
  });

  it("the field state is restored after a simple field: the next text is plain again", async () => {
    const body = para(`<w:fldSimple w:instr=" TOC ">${run("hiddenword")}</w:fldSimple>${run("visible")}`);
    expect(await pairs({ body })).toEqual(["visible:text"]);
  });
});

describe("readSource — hidden runs and content controls", () => {
  it("a run hidden by w:vanish is not read", async () => {
    expect(await pairs({ body: para(run("shown") + hiddenRun("<w:t>secret</w:t>")) })).toEqual(["shown:text"]);
  });

  it("a hidden run still opens and closes the fields: only its field characters count", async () => {
    const body = para(
      [
        run("before"),
        hiddenRun(`<w:fldChar w:fldCharType="begin"/>`),
        run("incode"),
        hiddenRun(`<w:fldChar w:fldCharType="separate"/>`),
        run("result"),
        hiddenRun(`<w:fldChar w:fldCharType="end"/><w:t>secret</w:t>`),
        run("after"),
      ].join(""),
    );
    expect(await pairs({ body })).toEqual(["before:text", "result:field-result", "after:text"]);
  });

  it("a hidden run's instruction text feeds the field code (a TOC field opened by hidden runs hides its result)", async () => {
    const body = para(
      [
        hiddenRun(
          `<w:fldChar w:fldCharType="begin"/><w:instrText> TOC </w:instrText><w:fldChar w:fldCharType="separate"/>`,
        ),
        run("tocword"),
        hiddenRun(`<w:fldChar w:fldCharType="end"/>`),
        run("after"),
      ].join(""),
    );
    expect(await pairs({ body })).toEqual(["after:text"]);
  });

  it("a run hidden by its paragraph style, by its character style, with the docDefaults toggled off by a direct value", async () => {
    const styles =
      `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>` +
      `<w:style w:type="paragraph" w:styleId="Hid"><w:rPr><w:vanish/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Hid"/></w:style>` +
      `<w:style w:type="character" w:styleId="CHid"><w:rPr><w:vanish/></w:rPr></w:style>`;
    const body = [
      para(run("normal")),
      `<w:p><w:pPr><w:pStyle w:val="Hid"/></w:pPr>${run("styled")}<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>forcedback</w:t></w:r></w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>${run("inherited")}</w:p>`,
      `<w:p><w:r><w:rPr><w:rStyle w:val="CHid"/></w:rPr><w:t>charhidden</w:t></w:r>${run("charplain")}</w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="Hid"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CHid"/></w:rPr><w:t>bothtoggled</w:t></w:r></w:p>`,
    ].join("");
    expect(await pairs({ body, styles })).toEqual([
      "normal:text",
      "forcedback:text",
      "charplain:text",
      "bothtoggled:text",
    ]);
  });

  it("a content control gives its content the content-control cause, nested paragraphs included", async () => {
    const body = `<w:sdt><w:sdtContent>${para(run("ctrl"))}</w:sdtContent></w:sdt>${para(run("plain"))}`;
    expect(await pairs({ body })).toEqual(["ctrl:content-control", "plain:text"]);
  });

  const gallery = (inner: string): string =>
    `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

  it("a table-of-contents content control with no TOC field in it is skipped whole", async () => {
    expect(await pairs({ body: gallery(para(run("tocgallery"))) + para(run("after")) })).toEqual(["after:text"]);
  });

  it("a table-of-contents content control that holds a TOC field is read, the field's cache apart", async () => {
    const inside = para(`<w:fldSimple w:instr=" TOC ">${run("cache")}</w:fldSimple>${run("kept")}`);
    expect(await pairs({ body: gallery(inside) })).toEqual(["kept:content-control"]);
  });

  it("another gallery is read like any content control", async () => {
    const body = `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr><w:sdtContent>${para(run("cover"))}</w:sdtContent></w:sdt>`;
    expect(await pairs({ body })).toEqual(["cover:content-control"]);
  });
});

describe("readSource — ruby, drawings and text boxes", () => {
  it("the base and the guide of a ruby both carry the ruby cause", async () => {
    const body = para(
      `${run("before")}<w:r><w:ruby><w:rubyBase>${run("base")}</w:rubyBase><w:rt>${run("guide")}</w:rt></w:ruby></w:r>${run("after")}`,
    );
    expect(await pairs({ body })).toEqual(["before:text", "base:ruby", "guide:ruby", "after:text"]);
  });

  it("a text box is read after the whole body, framed by line breaks, with the text-box cause", async () => {
    const box = `<w:r><w:pict><w:txbxContent>${para(tight("boxed"))}</w:txbxContent></w:pict></w:r>`;
    const body = para(tight("first") + box + tight("second")) + para(tight("last"));
    const read1 = await read({ body });
    expect(read1.text).toBe("\nfirstsecond\n\nlast\n" + "\n" + "\nboxed\n" + "\n");
    expect(wordsOf(read1.text).map((w, i) => `${w}:${read1.wordCauses[i]}`)).toEqual([
      "firstsecond:text",
      "last:text",
      "boxed:text-box",
    ]);
  });

  it("a text box inside a text box: the inner one is deferred first, so it comes before the outer one", async () => {
    const inner = `<w:r><w:pict><w:txbxContent>${para(run("inner"))}</w:txbxContent></w:pict></w:r>`;
    const outer = `<w:r><w:pict><w:txbxContent>${para(run("outer") + inner)}</w:txbxContent></w:pict></w:r>`;
    expect(await pairs({ body: para(run("main") + outer) })).toEqual(["main:text", "inner:text-box", "outer:text-box"]);
  });

  it("a drawing hidden by its docPr is not read; a visible one is, its text box deferred", async () => {
    const drawing = (hidden: string, text: string): string =>
      `<w:r><w:drawing><wp:inline><wp:docPr id="1" name="d"${hidden}/><w:txbxContent>${para(run(text))}</w:txbxContent></wp:inline></w:drawing></w:r>`;
    const body = para(
      run("main") +
        drawing(' hidden="1"', "hiddenbox") +
        drawing(' hidden=" TRUE "', "alsohidden") +
        drawing("", "shownbox") +
        drawing(' hidden="0"', "zerobox"),
    );
    expect(await pairs({ body })).toEqual(["main:text", "shownbox:text-box", "zerobox:text-box"]);
  });

  it("a drawing wholly outside the largest page is not read; one inside is", async () => {
    const anchor = (x: number, text: string): string =>
      `<w:r><w:drawing><wp:anchor simplePos="1"><wp:simplePos x="${x}" y="0"/><wp:extent cx="100" cy="100"/><w:txbxContent>${para(run(text))}</w:txbxContent></wp:anchor></w:drawing></w:r>`;
    const body =
      para(anchor(99_999_999, "offpage") + anchor(10, "onpage")) +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    expect(await pairs({ body })).toEqual(["onpage:text-box"]);
  });
});

describe("readSource — called notes", () => {
  const fnText = (t: string): string => para(run(t));

  it("a note is read only when called; footnotes then endnotes in part order, a note called by a note after them", async () => {
    const body = para(
      `${run("body")}<w:r><w:endnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`,
    );
    const footnotes =
      note("footnote", 1, fnText("fnone") + para(`<w:r><w:endnoteReference w:id="2"/></w:r>`)) +
      note("footnote", 2, fnText("neverasked"));
    const endnotes = note("endnote", 1, fnText("enone")) + note("endnote", 2, fnText("entwo"));
    expect(await pairs({ body, footnotes, endnotes })).toEqual([
      "body:text",
      "fnone:footnote",
      "enone:endnote",
      "entwo:endnote",
    ]);
  });

  it("of several notes with one id, the last is the one; a call with no id and a call to an absent note read nothing", async () => {
    const body = para(
      `${run("body")}<w:r><w:footnoteReference/></w:r><w:r><w:footnoteReference w:id="9"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`,
    );
    const footnotes = note("footnote", 1, fnText("firstversion")) + note("footnote", 1, fnText("lastversion"));
    expect(await pairs({ body, footnotes })).toEqual(["body:text", "lastversion:footnote"]);
  });

  it("a call made inside the code of a field is not a call; one made in its result is", async () => {
    const body = para(
      [
        fld("begin"),
        `<w:r><w:footnoteReference w:id="1"/></w:r>`,
        fld("separate"),
        `<w:r><w:footnoteReference w:id="2"/></w:r>`,
        fld("end"),
      ].join(""),
    );
    const footnotes = note("footnote", 1, fnText("incode")) + note("footnote", 2, fnText("inresult"));
    expect(await pairs({ body, footnotes })).toEqual(["inresult:footnote"]);
  });

  it("a field left open in the body does not tag the words of the notes: the state is reset before each note", async () => {
    const body = para(
      [fld("begin"), fld("separate"), `<w:r><w:footnoteReference w:id="1"/></w:r>`, run("open")].join(""),
    );
    const footnotes = note("footnote", 1, fnText("notetext"));
    expect(await pairs({ body, footnotes })).toEqual(["open:field-result", "notetext:footnote"]);
  });

  it("the text box of a note goes after the note's own paragraphs, before the next note", async () => {
    const body = para(`<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r>`);
    const box = `<w:r><w:pict><w:txbxContent>${para(run("boxed"))}</w:txbxContent></w:pict></w:r>`;
    const footnotes =
      note("footnote", 1, para(run("main") + box) + fnText("tail")) + note("footnote", 2, fnText("next"));
    expect(await pairs({ body, footnotes })).toEqual([
      "main:footnote",
      "tail:footnote",
      "boxed:text-box",
      "next:footnote",
    ]);
  });

  it("notes that call each other terminate, each read once", async () => {
    const body = para(`<w:r><w:footnoteReference w:id="1"/></w:r>`);
    const footnotes =
      note("footnote", 1, para(`${run("a")}<w:r><w:footnoteReference w:id="2"/></w:r>`)) +
      note("footnote", 2, para(`${run("b")}<w:r><w:footnoteReference w:id="1"/></w:r>`));
    expect(await pairs({ body, footnotes })).toEqual(["a:footnote", "b:footnote"]);
  });
});

describe("readSource — what a context carries down to its children", () => {
  const HIDDEN_STYLE = `<w:style w:type="paragraph" w:styleId="Hid"><w:rPr><w:vanish/></w:rPr></w:style>`;
  const inHiddenStyle = (inner: string): string => `<w:p><w:pPr><w:pStyle w:val="Hid"/></w:pPr>${inner}</w:p>`;
  const call = (id: number): string => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
  const wideRow = (firstCell: string): string =>
    `<w:tbl><w:tr>${[firstCell, ...Array.from({ length: MAX_TABLE_COLS }, (_, c) => `<w:tc>${para(run(`w${c}`))}</w:tc>`)].join("")}</w:tr></w:tbl>`;

  it("the paragraph style hides the runs of a content control and of a ruby inside the paragraph", async () => {
    const body =
      inHiddenStyle(
        `<w:sdt><w:sdtContent>${run("controlled")}</w:sdtContent></w:sdt><w:ruby><w:rubyBase>${run("based")}</w:rubyBase></w:ruby>`,
      ) + para(run("shown"));
    expect(await pairs({ body, styles: HIDDEN_STYLE })).toEqual(["shown:text"]);
  });

  it("a text box standing directly in a hidden-style paragraph starts with no style of its own: its runs are read", async () => {
    const body = inHiddenStyle(`<w:txbxContent>${run("direct")}</w:txbxContent>`);
    expect(await pairs({ body, styles: HIDDEN_STYLE })).toEqual(["direct:text-box"]);
  });

  it("a withheld table expects none of its words nor its text boxes, but its note calls are still read", async () => {
    const box = `<w:r><w:pict><w:txbxContent>${para(run("wbox"))}</w:txbxContent></w:pict></w:r>`;
    const body = wideRow(`<w:tc>${para(run("cell") + box + call(1))}</w:tc>`) + para(run("after"));
    const footnotes = note("footnote", 1, para(run("noteone")));
    expect(await pairs({ body, footnotes })).toEqual(["after:text", "noteone:footnote"]);
  });

  it("a vMerge continuation cell is folded whether its table is withheld or shown: its words and its note calls are read in neither", async () => {
    const restart = `<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${para(run("top"))}</w:tc>`;
    const continuation = `<w:tc><w:tcPr><w:vMerge/></w:tcPr>${para(run("folded") + call(1))}</w:tc>`;
    const plain = (c: number): string => `<w:tc>${para(run(`w${c}`))}</w:tc>`;
    const others = Array.from({ length: MAX_TABLE_COLS }, (_, c) => plain(c));
    const footnotes = note("footnote", 1, para(run("noteone")));
    const shown = `<w:tbl><w:tr>${restart}${plain(0)}</w:tr><w:tr>${continuation}${plain(1)}</w:tr></w:tbl>`;
    const withheld = `<w:tbl><w:tr>${restart}${others.join("")}${plain(9)}</w:tr><w:tr>${continuation}${plain(1)}</w:tr></w:tbl>`;
    expect(await pairs({ body: shown, footnotes })).toEqual(["top:text", "w0:text", "w1:text"]);
    expect(await pairs({ body: withheld + para(run("after")), footnotes })).toEqual(["after:text"]);
  });

  it("a withheld table keeps the paragraph style it sits in: a call in a run that style hides is no call", async () => {
    const body = inHiddenStyle(wideRow(`<w:tc>${para(run("lead"))}${call(1)}</w:tc>`));
    const footnotes = note("footnote", 1, para(run("noteone")));
    expect(await pairs({ body, styles: HIDDEN_STYLE, footnotes })).toEqual([]);
  });

  it("a hidden run registers no note call", async () => {
    const body = para(run("body") + hiddenRun(`<w:footnoteReference w:id="1"/>`));
    expect(await pairs({ body, footnotes: note("footnote", 1, para(run("noteone"))) })).toEqual(["body:text"]);
  });

  it("a call with no id reads no note, not even one that has no id either", async () => {
    const body = para(`${run("body")}<w:r><w:footnoteReference/></w:r>`);
    const footnotes = `<w:footnote w:type="separator">${para(run("idless"))}</w:footnote>`;
    expect(await pairs({ body, footnotes })).toEqual(["body:text"]);
  });

  it("a note that calls a note the part does not hold is read, the missing one skipped", async () => {
    const body = para(call(1));
    const footnotes = note("footnote", 1, para(run("noteone") + call(9)));
    expect(await pairs({ body, footnotes })).toEqual(["noteone:footnote"]);
  });
});

describe("readSource — nesting", () => {
  it("elements nested to the parser's ceiling (512) are walked without exhausting the stack", async () => {
    const depth = 500;
    const body = para(
      `${'<w:fldSimple w:instr=" X ">'.repeat(depth)}${run("deep")}${"</w:fldSimple>".repeat(depth)}${"<w:smartTag>".repeat(5)}${run("wrapped")}${"</w:smartTag>".repeat(5)}`,
    );
    expect(await pairs({ body })).toEqual(["deep:field-result", "wrapped:text"]);
  });
});

describe("readNumberedParagraphs — which paragraphs carry a numbering", () => {
  it("a direct numId, a numbered style, and a direct numId 0 that cancels the style", async () => {
    const styles = `<w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Plain"/>`;
    const num = (id: string): string => `<w:pPr><w:numPr><w:numId w:val="${id}"/></w:numPr></w:pPr>`;
    const body = [
      `<w:p>${num("1")}${tight("direct")}</w:p>`,
      `<w:p>${num("0")}${tight("zero")}</w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="List"/></w:pPr>${tight("viastyle")}</w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="List"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>${tight("cancelled")}</w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="Plain"/></w:pPr>${tight("plain")}</w:p>`,
    ].join("");
    expect(await readNumberedParagraphs(await docxOf({ body, styles }))).toEqual(["direct", "viastyle"]);
  });
});

describe("readSource — tables", () => {
  it("a table, its rows and its cells are each framed by line breaks", async () => {
    const cell = (t: string): string => `<w:tc>${para(tight(t))}</w:tc>`;
    const body = `<w:tbl><w:tr>${cell("a")}${cell("b")}</w:tr></w:tbl>`;
    expect((await read({ body })).text).toBe("\n" + "\n" + "\n" + "\na\n" + "\n" + "\n" + "\nb\n" + "\n" + "\n" + "\n");
  });
});
