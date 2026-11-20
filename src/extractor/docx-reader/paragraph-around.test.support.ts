/** Test support: the span of the OUTER paragraph that holds a given position of a Word `document.xml`, for the tests that
 *  move or repeat the paragraph of a drawing. A drawing sits in a paragraph that may hold others (a text box's own), so
 *  « the next `</w:p>` » closes the wrong one: the splice was not well-formed XML, which fast-xml-parser, the reader the
 *  product used until 26/10/2026, read in silence and our own XML reader, which replaced it that day, refuses as damaged.
 *  NOT here: anything about what the paragraph means.
 *  Called by: `docx-reader.budget.test.ts`, `docx-reader.google-swap.test.ts`, `docx-reader.note-drawings.test.ts` (`withFourDrawings`: the first two only). */

/** A paragraph's start tag (not `<w:pPr`) or its end tag. */
const TAG = /<w:p(?=[ >/])[^>]*>|<\/w:p>/g;

/** `[from, to)`: the outermost `<w:p …>…</w:p>` around `at`, paragraphs matched by depth (a self-closed one opens nothing). */
export const outerParagraphAround = (xml: string, at: number): { from: number; to: number } => {
  let from = -1;
  let depth = 0;
  for (const tag of xml.matchAll(TAG)) {
    if (tag[0] === "</w:p>") {
      depth--;
      const to = tag.index + tag[0].length;
      if (depth === 0 && from <= at && at < to) return { from, to };
    } else if (!tag[0].endsWith("/>")) {
      if (depth === 0) from = tag.index;
      depth++;
    }
  }
  throw new Error("no paragraph around the position");
};

/** The `document.xml` of fixture `gdoc-docx/13-Essai-Dessin-Formes.docx` with the paragraph of its first drawing followed by
 *  three copies, as well-formed XML: the first copy identical, the next two each with a longer « Bonjour » in the drawn
 *  text box, so that four drawings show three different SVGs, the first two the same one (the allowance tests count SVGs
 *  and bytes). The splice used to cut the paragraph at its first `</w:p>`, a text box's, and repeat the fragment: the XML
 *  was unbalanced, the XML reader then in use read it in silence into other nestings, and the three SVGs came from that accident. */
export const withFourDrawings = (xml: string): string => {
  const { from, to } = outerParagraphAround(xml, xml.indexOf("<w:drawing>"));
  const paragraph = xml.slice(from, to);
  const longer = (extra: string): string => {
    const text = '<w:t xml:space="preserve">Bonjour</w:t>';
    if (!paragraph.includes(text)) throw new Error("the fixture's first drawing no longer holds « Bonjour »");
    return paragraph.replace(text, `<w:t xml:space="preserve">Bonjour${extra}</w:t>`);
  };
  return xml.slice(0, to) + paragraph + longer(" two") + longer(" three three") + xml.slice(to);
};
