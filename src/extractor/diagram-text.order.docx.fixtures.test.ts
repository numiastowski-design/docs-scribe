import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";

import type { IRBlock } from "doc-ir";
import { extractDocx } from "./docx";

// Milestone 1 of the diagram text (relation order), Word side: the text boxes of ONE paragraph's drawings are published
// in the order of the connectors the geometry declares (no `a:stCxn` in any Word file measured). `BEFORE` is the
// word count and the sha256 of the sorted words each document published before the order moved: only the order may
// move, never a word.
// Came from the publisher's application on 28/11/2026, where it read only this package.
vi.setConfig({ testTimeout: 60_000 });

const here = dirname(fileURLToPath(import.meta.url));
const fixture = (path: string): Uint8Array => new Uint8Array(readFileSync(join(here, "../../fixtures", path)));

/** Words count and the sha256 of the sorted words, per document, as published before the order moved. */
const BEFORE: Record<string, { words: number; sha256: string }> = {
  "diagrammes-generes/a-en-label-on-connector.docx": {
    words: 14,
    sha256: "b259181a815e9c077fd8a18430436292505094098712517e633407a5d73dc8a4",
  },
  "diagrammes-generes/b-fr-label-above-arrow.docx": {
    words: 13,
    sha256: "8c39c9aca1b2a924e517bac3ad4a1b3b5abc82589aee84950556b53e01a5e29c",
  },
  "diagrammes-generes/e-en-org-chart-elbow.docx": {
    words: 12,
    sha256: "bf81f9b6706506eaed599a4f948801b5686d0af78c8ca7e06079f7ceb762e9d9",
  },
  "diagrammes-reel/05-libya-qatar-org-structure.docx": {
    words: 3649,
    sha256: "55c866c76bee1c7b4b74d71cb938710a048d29e9c22ff622480be6a6116a0aec",
  },
};

const paragraphs = async (path: string): Promise<string[]> => {
  const { doc } = await extractDocx(fixture(path));
  return doc.blocks.flatMap((block: IRBlock) =>
    block.kind === "paragraph" ? [block.inlines.map((inline) => inline.text).join("")] : [],
  );
};

const collect = (node: unknown, out: string[]): void => {
  if (Array.isArray(node)) for (const child of node) collect(child, out);
  else if (node !== null && typeof node === "object") {
    for (const [key, value] of Object.entries(node)) {
      if (key === "text" && typeof value === "string") out.push(value);
      else collect(value, out);
    }
  }
};

describe("text boxes in the order of their connectors (docx, milestone 1)", () => {
  it("a: the label between source and target", async () => {
    const got = await paragraphs("diagrammes-generes/a-en-label-on-connector.docx");
    expect(got.slice(1)).toEqual(["Order", "is converted into", "Invoice"]);
  });

  it("b: the label above the arrow between source and target", async () => {
    const got = await paragraphs("diagrammes-generes/b-fr-label-above-arrow.docx");
    expect(got.slice(1)).toEqual(["Commande", "génère", "Facture"]);
  });

  it("e: the org chart, root then children (the same order as before)", async () => {
    const got = await paragraphs("diagrammes-generes/e-en-org-chart-elbow.docx");
    expect(got.slice(1)).toEqual(["Director", "Sales", "Engineering", "Support"]);
  });
});

describe("conservation over the docx diagram fixtures: only the order moves", () => {
  for (const [file, was] of Object.entries(BEFORE)) {
    it(`${file}: same words`, async () => {
      const { doc } = await extractDocx(fixture(file));
      const texts: string[] = [];
      collect(doc.blocks, texts);
      const words = texts
        .flatMap((text) => text.split(/\s+/))
        .filter((word) => word !== "" && word !== "·")
        .sort((left, right) => (left < right ? -1 : Number(left > right)));
      expect({ words: words.length, sha256: createHash("sha256").update(words.join("\n")).digest("hex") }).toEqual(was);
    });
  }
});
