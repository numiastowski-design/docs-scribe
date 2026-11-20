/** The helper a sweep asks for its documents (`public-fixtures.test.support.ts`) gives the rows the list calls `public`, of the
 *  extensions asked, and nothing else. The list is read a second time here, line by line with a pattern, so that the two readings
 *  must agree. Run: npx vitest run --project=main packages/docs-scribe/src/extractor/public-fixtures.test.ts */
import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { publicFixturePath, publicFixtures, publicRowsOf } from "./public-fixtures.test.support";

const LIST = readFileSync(publicFixturePath("fixtures/PUBLIC.tsv"), "utf8");
// A line may end with a carriage return (a checkout on Windows): neither end-of-line character is part of a path.
const rowsSaying = (status: string, extension: string): string[] =>
  [...LIST.matchAll(new RegExp(String.raw`^([^#\t\r\n][^\t\r\n]*\.${extension})\t${status}\t`, "gmu"))].map(
    (match) => match[1] ?? "",
  );

describe("publicFixtures: the documents a sweep may read", () => {
  it("gives every public row of an extension, in the order of the list, each a file on disk", () => {
    const documents = publicFixtures(".docx");
    expect(documents).toEqual(rowsSaying("public", "docx"));
    expect(documents.length).toBeGreaterThan(30);
    expect(documents).toContain("fixtures/docx/07-Fiche-Contact.docx");
    expect(documents.filter((path) => !existsSync(publicFixturePath(path)))).toEqual([]);
  });

  it("takes several extensions at once, and gives nothing for an extension no document has", () => {
    expect(publicFixtures(".docx", ".xlsx").length).toBe(
      rowsSaying("public", "docx").length + rowsSaying("public", "xlsx").length,
    );
    expect(publicFixtures(".nothing")).toEqual([]);
  });

  it("leaves out the rows still in question, on a list that holds one (the real list may hold none)", () => {
    const list = [
      "# a comment",
      ["fixtures/docx/a.docx", "public", "CC0-1.0", "generated"].join("\t"),
      ["fixtures/docx/b.docx", "pending", "", "asked"].join("\t"),
      `${["fixtures/xlsx/c.xlsx", "public", "CC0-1.0", "generated"].join("\t")}\r`,
      ["fixtures/pdf/d.pdf", "public", "CC0-1.0", "generated"].join("\t"),
    ].join("\n");
    expect(publicRowsOf(list, [".docx", ".xlsx"])).toEqual(["fixtures/docx/a.docx", "fixtures/xlsx/c.xlsx"]);
    const pending = [...rowsSaying("pending", "docx"), ...rowsSaying("pending", "xlsx")];
    const given = new Set(publicFixtures(".docx", ".xlsx"));
    expect(pending.filter((path) => given.has(path))).toEqual([]);
  });
});
