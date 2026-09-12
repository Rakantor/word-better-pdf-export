/**
 * Usage: npm run make-test-pdf [-- <input.docx> <output.pdf>]
 * Creates a PDF that imitates Word's compressed export, for local testing.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { nodePlatform } from "../src/node/platform-node";
import { simulateWordExport } from "./simulate-word-export";

const [docxPath = "file-sample.docx", outPath = "file-sample.simulated-word.pdf"] = process.argv.slice(2);
const docx = new Uint8Array(readFileSync(docxPath));
simulateWordExport(docx, { platform: nodePlatform, addDecoy: true }).then((pdf) => {
  writeFileSync(outPath, pdf);
  console.log(`wrote ${outPath} (${pdf.length} bytes)`);
});
