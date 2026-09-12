/**
 * Debug helper: dump every image XObject of a PDF into a directory so before/after
 * exports can be compared at native resolution.
 *
 *   npx tsx scripts/extract-images.ts <file.pdf> <out-dir>
 *
 * JPEG streams are written as-is (.jpg); everything else is decoded to PNG.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { PDFDocument } from "pdf-lib";
import { decodePdfImage } from "../src/core/pdf-decode";
import { collectPdfImages } from "../src/core/pdf-images";
import { nodePlatform } from "../src/node/platform-node";

async function main(): Promise<void> {
  const [file, outDir = "extracted-images"] = process.argv.slice(2);
  if (!file) {
    console.error("usage: extract-images <file.pdf> [out-dir]");
    process.exit(2);
  }
  mkdirSync(outDir, { recursive: true });
  const doc = await PDFDocument.load(readFileSync(file), { ignoreEncryption: true, updateMetadata: false });
  const images = collectPdfImages(doc);
  console.log(`${images.length} image(s) in ${file}`);
  for (const img of images) {
    const base = `p${img.pages.map((p) => p + 1).join("-")}_${img.ref.toString().replace(/\s/g, "_")}_${img.width}x${img.height}`;
    if (img.filters.length === 1 && img.filters[0] === "DCTDecode") {
      const out = join(outDir, `${base}.jpg`);
      writeFileSync(out, img.stream.contents);
      console.log(`  ${out} (${img.stream.contents.length} bytes, JPEG passthrough)`);
      continue;
    }
    const raster = await decodePdfImage(doc, img, nodePlatform);
    if (!raster) {
      console.log(`  ${base}: cannot decode [${img.filters.join("+")}]`);
      continue;
    }
    const out = join(outDir, `${base}.png`);
    writeFileSync(out, await nodePlatform.encodeImage(raster, "image/png"));
    console.log(`  ${out} (decoded from ${img.filters.join("+") || "raw"})`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
