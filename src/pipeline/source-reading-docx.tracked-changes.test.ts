/** Word shows text inserted under track changes (`w:ins`, ECMA-376 Part 1) as normal text and hides deleted text
 *  (`w:del`); the page publishes the insertion, so the control reading must expect it and must not expect the deletion. */
import JSZip from "jszip";
import { expect, it } from "vitest";

import { readSource, wordsOf } from "./source-reading-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

it("expects the words of a w:ins and none of a w:del", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>kept </w:t></w:r><w:ins w:id="1" w:author="a"><w:r><w:t>added</w:t></w:r></w:ins></w:p>` +
      `<w:p><w:r><w:t>stays </w:t></w:r><w:del w:id="2" w:author="a"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p></w:body></w:document>`,
  );
  const { text } = await readSource(await zip.generateAsync({ type: "uint8array" }));
  expect(wordsOf(text)).toEqual(["kept", "added", "stays"]);
});
