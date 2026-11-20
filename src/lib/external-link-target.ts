/** The ONE provenance filter for a link read from an Office file: which declared target may become an href.
 *  Decides: only the target of an External relationship, and never one of OUR internal forms.
 *  NOT here: the schemes a link may be published with (`lib/link-schemes.ts`, applied by the renderer), nor rendering a link.
 *  Called by: `extractor/docx-reader/body.ts`, `extractor/pptx-reader/text.ts`, `extractor/xlsx-cells.ts`,
 *  `extractor/google-sheet.ts`. */
import { INTERNAL_HREF_FORM } from "doc-ir/internal";

/** The target of an External relationship, unless it spells one of our internal forms (`lib/ir.ts#INTERNAL_HREF_FORM`);
 *  `undefined` otherwise. The scheme is not judged here. */
export const externalLinkTarget = (link: { external: boolean; target: string } | undefined): string | undefined =>
  link?.external && !INTERNAL_HREF_FORM.test(link.target) ? link.target : undefined;
