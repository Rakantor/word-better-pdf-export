/** Debug helper: list the image XObjects in one or more PDFs.  Usage: npx tsx scripts/inspect-pdf.ts a.pdf b.pdf */
import { readFileSync } from "node:fs";
import { PDFDocument, PDFName, PDFRawStream } from "pdf-lib";
import { collectPdfImages } from "../src/core/pdf-images";
async function main() {
  for (const f of process.argv.slice(2)) {
    const doc = await PDFDocument.load(readFileSync(f), { updateMetadata: false });
    console.log(`== ${f}: ${doc.getPageCount()} pages, ${doc.context.enumerateIndirectObjects().length} objects`);
    for (const img of collectPdfImages(doc)) {
      console.log(`  ${img.ref} ${img.width}x${img.height} ${img.filters} bytes=${img.stream.contents.length} cs=${img.stream.dict.lookup(PDFName.of("ColorSpace"))} pages=${img.pages} slots=${img.slots.length}`);
    }
    const streams = doc.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFRawStream && (o as PDFRawStream).dict.lookup(PDFName.of("Subtype")) === PDFName.of("Image"));
    console.log(`  total image streams in file: ${streams.length}`);
  }
}
main();
