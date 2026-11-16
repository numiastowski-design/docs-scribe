# docs-scribe

Reads Word, PowerPoint, Excel and Google exports into a structured representation, and checks that no
word was lost on the way.

The declared-formats reader of [Miksumia Publisher](https://miksumia.com), the Atlassian Marketplace
[app](https://marketplace.atlassian.com/apps/2647582831/miksumia-publisher-for-confluence) that turns
uploaded files into Confluence pages. PDF is the other half, published as
[skryba](https://www.npmjs.com/package/skryba).

Published on npm as [`docs-scribe`](https://www.npmjs.com/package/docs-scribe): `npm i docs-scribe`.

## Read a document

```js
import { readFileSync } from "node:fs";
import { extractDocx } from "docs-scribe";

const { doc, images, notices } = await extractDocx(new Uint8Array(readFileSync("report.docx")));

doc.blocks;   // headings, paragraphs, lists, tables, image references
notices;      // every part it could not publish, each with a named reason
```

One representation for all four formats: a heading is a heading whether it came from Word, a slide or a
spreadsheet; tables keep their merges and lists keep their levels.

```js
import { extractPptx, extractXlsx, extractGoogleDoc, extractGoogleSheet } from "docs-scribe";

await extractPptx(bytes);        // one tab per slide, speaker notes included
await extractXlsx(bytes, "en");  // one tab per visible sheet
```

Google Docs and Sheets come in as their API payloads; their `.docx` and `.xlsx` exports go through the
readers above.

## Your own ceilings

It withholds nothing by default: a ceiling shaped by where a document is going is yours to set. The
reader's own bounds — bytes, parts, XML depth — have values you can only lower.

```js
await extractXlsx(bytes, "en", { workbookTabs: 20, sheetRows: 2000, sheetCols: 100 });
```

## Refuse a hostile archive before reading it

These formats are archives of XML, so refusing a bomb is part of reading them, not an option beside it.
No DTD is ever read, so entity expansion cannot happen, and an archive is judged on its **declared**
sizes before a byte is inflated.

```js
import { assertZipDeclaredSizesSafe } from "docs-scribe";
import { ReaderError } from "docs-scribe/light";

assertZipDeclaredSizesSafe(bytes); // throws on a declared-size bomb or on overlapping entries

try {
  await extractDocx(bytes);
} catch (error) {
  if (error instanceof ReaderError) error.code; // "FILE_UNSAFE", "EXTRACT_FAILED"…
}
```

## Check what you published

`docs-scribe/light` loads no reader. Hand it the text you published, block by block, and it reports the
words the page lacks, grouped by cause. It never repairs: a word missing from declared text is a defect
of the reader, not a render accident to paper over.

```js
import { missingWordGroups } from "docs-scribe/light";
import { undrawnDiagramWords } from "docs-scribe";

const groups = missingWordGroups(source, { blocks: myPublishedTexts }, [], undrawnDiagramWords(doc.blocks));

groups; // [{ cause: "text", words: [...] }, { cause: "diagram", words: [...] }]
```

## Known limits

- **A SmartArt diagram is published as its picture**, so its text is in the document but not on the
  page: `undrawnDiagramWords` tells you which words those are.
- **An equation gives its characters, never its mathematical layout**, framed by a word boundary so the
  words around it stay findable.
- **An Excel formula gives its last cached value**, not a recomputation: a file saved by a tool that
  cached none has empty cells, counted and named by a `formula-cells-empty` notice.
- **Tracked changes are published in their accepted state**: an insertion is kept, a deletion is gone.
- **A chart that carries no copy of its values is read from its sheet** (0.4.0), within a ceiling on the
  cells it may read. Two defects are known in 0.4.0 and will be fixed in 0.4.1: a range shared by
  several series is counted once per series, so a large chart can be set aside although it fits; and a
  chart set aside by that ceiling comes out as its title and an empty table, with no notice saying so.

## The paper this reader is built on

Side Liu, Jiangpeng Liu, Jinwen Xin, Guojun Peng and Jiang Ming, **"Beyond the Editing Canvas: Evidence
Divergence in OOXML-to-LLM Ingestion"**, arXiv:2608.25880 — <https://arxiv.org/abs/2608.25880>.

They went through ECMA-376 looking for places where one valid file shows a person one thing and hands
another to whatever program reads it, and catalogued twenty-one of them. That is this reader's subject
seen from the other side: their twenty-one are text a human never saw reaching a machine, and our own
failure mode is the mirror image — text a human did see, lost on the way to the page.

We checked all twenty-one against this code, one by one, and wrote each verdict down before changing
anything. Hidden sheets, filtered rows, hidden columns, hidden runs and hidden shapes are never
published. Where we still diverge, it is named in that audit rather than discovered by someone else.
Their lesson is behind every decision above: **the extraction path, not the model, decides what a
language model sees of a document.**

## What other readers do, measured

Penghui Li, **"When File Ingestion Misleads Language Models: Measuring and Mitigating Divergences in
Office Documents"**, arXiv:2611.05806 — <https://arxiv.org/abs/2611.05806>.

He put ten widely used extractors on a bench and counted. Thirty-one channels diverge for at least one
of them: nineteen hand a model content no human saw, twelve drop content a human did see. **Six of the
seven spreadsheet extractors emit the stored value instead of the formatted one**, and **all six Word
extractors fuse split runs** — which is what makes his sharpest attack work: a hidden `un` in front of
a visible `authorized` comes out as `unauthorized`, an ordinary dictionary word no filter can catch.

This reader does neither. A number is written by our own formatter from the file's format code, and a
hidden run is dropped before anything is joined — both directions are held by tests, including the one
that proves our runs really are joined, so the immunity is the rule and not an accident.

His paper also names the failure that is ours rather than an attacker's: **omission**, content a human
sees and the reader loses. That is what the conservation check above exists to measure.

## Licence

AGPL-3.0-or-later: run a modified version as a network service and its source must be available to its
users. This repository is that source — a snapshot of the packages the npm bundle is built from, taken
from the development monorepo at version 0.4.0 (16 November 2026). The entry points are
`packages/doc-read/src/lib/docs.ts` and `packages/doc-read/src/lib/docs-light.ts`; the npm package ships
the bundle built from them. Questions: <contact@miksumia.com>.
