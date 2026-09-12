import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export const SAMPLE_DOCX = new Uint8Array(readFileSync(new URL("../file-sample.docx", import.meta.url)));

/** Return a copy of the docx with word/document.xml rewritten by `edit`. */
export function editDocumentXml(docx: Uint8Array, edit: (xml: string) => string): Uint8Array {
  const files = unzipSync(docx);
  const xml = strFromU8(files["word/document.xml"]);
  const edited = edit(xml);
  if (edited === xml) throw new Error("edit did not change document.xml");
  files["word/document.xml"] = strToU8(edited);
  return zipSync(files);
}
