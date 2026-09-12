/**
 * Command-line version of what the add-in does, for testing against a PDF that
 * Word itself produced (File > Save As > PDF, or File > Export):
 *
 *   npm run fix-pdf -- <document.docx> <word-export.pdf> [output.pdf]
 */
import { readFileSync, writeFileSync } from "node:fs";
import { replaceImagesInPdf } from "../src/core";
import { nodePlatform } from "../src/node/platform-node";

async function main(): Promise<void> {
  const [docxPath, pdfPath, outPath] = process.argv.slice(2);
  if (!docxPath || !pdfPath) {
    console.error("usage: fix-pdf <document.docx> <word-export.pdf> [output.pdf]");
    process.exit(2);
  }
  const output = outPath ?? pdfPath.replace(/\.pdf$/i, "") + ".original-images.pdf";
  const docx = new Uint8Array(readFileSync(docxPath));
  const pdf = new Uint8Array(readFileSync(pdfPath));

  const { pdf: fixed, report } = await replaceImagesInPdf(pdf, docx, { platform: nodePlatform, log: (m) => console.log(m) });
  writeFileSync(output, fixed);
  console.log(
    `\n${report.replaced} replaced, ${report.skipped} skipped, ${report.unmatched} unmatched ` +
      `(${report.pdfImages} image(s) in PDF, ${report.docxImageParts} in docx)\n` +
      `${(report.inputBytes / 1024).toFixed(0)} KiB -> ${(report.outputBytes / 1024).toFixed(0)} KiB: ${output}`,
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
