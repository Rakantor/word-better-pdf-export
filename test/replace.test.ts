import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { PDFDocument, PDFName } from "pdf-lib";
import { simulateWordExport } from "../scripts/simulate-word-export";
import { replaceImagesInPdf } from "../src/core";
import { collectPdfImages } from "../src/core/pdf-images";
import { croppedSize } from "../src/core/raster";
import { nodePlatform } from "../src/node/platform-node";
import { editDocumentXml, SAMPLE_DOCX } from "./helpers";

const platform = nodePlatform;
const ORIGINAL = { width: 2800, height: 1867, bytes: 555181 };

describe("replaceImagesInPdf", () => {
  it("replaces the downsampled picture with the original JPEG and leaves unrelated images alone", async () => {
    const wordPdf = await simulateWordExport(SAMPLE_DOCX, { platform, addDecoy: true });
    const { pdf, report } = await replaceImagesInPdf(wordPdf, SAMPLE_DOCX, { platform });

    assert.equal(report.pdfImages, 2);
    assert.equal(report.replaced, 1);
    assert.equal(report.unmatched, 1);
    const replaced = report.results.find((r) => r.status === "replaced")!;
    assert.equal(replaced.status, "replaced");
    assert.match(replaced.source, /image1\.jpeg/);
    assert.ok(replaced.similarity > 0.95, `similarity ${replaced.similarity}`);

    const out = await PDFDocument.load(pdf, { updateMetadata: false });
    assert.equal(out.getPageCount(), 2);
    const images = collectPdfImages(out);
    const big = images.find((i) => i.width === ORIGINAL.width)!;
    assert.ok(big, "original-resolution image present");
    assert.equal(big.height, ORIGINAL.height);
    assert.deepEqual(big.filters, ["DCTDecode"]);
    assert.equal(big.stream.contents.length, ORIGINAL.bytes, "original JPEG bytes embedded untouched");
    assert.deepEqual(big.pages, [0]);

    // The compressed original object was removed, nothing else was.
    const before = await PDFDocument.load(wordPdf, { updateMetadata: false });
    assert.equal(out.context.enumerateIndirectObjects().length, before.context.enumerateIndirectObjects().length);
    assert.equal(images.length, 2);
  });

  it("handles a cropped picture by embedding only the cropped region", async () => {
    const cropped = editDocumentXml(SAMPLE_DOCX, (xml) =>
      xml.replace('<a:blip r:embed="rId4"/>', '<a:blip r:embed="rId4"/><a:srcRect l="10000" t="20000" r="30000" b="5000"/>'),
    );
    const wordPdf = await simulateWordExport(cropped, { platform });
    const { pdf, report } = await replaceImagesInPdf(wordPdf, cropped, { platform });

    assert.equal(report.replaced, 1);
    const r = report.results[0];
    assert.equal(r.status, "replaced");
    assert.deepEqual(r.crop, { left: 0.1, top: 0.2, right: 0.3, bottom: 0.05 });
    assert.match(r.to, /cropped JPEG/);

    const out = await PDFDocument.load(pdf, { updateMetadata: false });
    const [img] = collectPdfImages(out);
    const expected = croppedSize(ORIGINAL.width, ORIGINAL.height, r.crop);
    assert.equal(img.width, expected.width);
    assert.equal(img.height, expected.height);
    assert.ok(img.width > 1177 && img.height > 982, "cropped region embedded at full source resolution");
  });

  it("skips pictures that carry effects Word bakes into the export", async () => {
    const withEffect = editDocumentXml(SAMPLE_DOCX, (xml) =>
      xml.replace('<a:blip r:embed="rId4"/>', '<a:blip r:embed="rId4"><a:grayscl/></a:blip>'),
    );
    const wordPdf = await simulateWordExport(withEffect, { platform });
    const { report } = await replaceImagesInPdf(wordPdf, withEffect, { platform });

    assert.equal(report.replaced, 0);
    assert.equal(report.skipped, 1);
    const r = report.results[0];
    assert.equal(r.status, "skipped");
    assert.match(r.reason, /grayscl/);
  });

  it("leaves an image alone when the PDF already contains the original bytes", async () => {
    const wordPdf = await simulateWordExport(SAMPLE_DOCX, { platform, passthrough: true });
    const { pdf, report } = await replaceImagesInPdf(wordPdf, SAMPLE_DOCX, { platform });
    assert.equal(report.replaced, 0);
    assert.equal(report.skipped, 1);
    assert.match((report.results[0] as { reason: string }).reason, /original bytes/);
    assert.ok(pdf.length > 0);
  });

  it("does not match when the picture is visually different (only aspect ratio agrees)", async () => {
    // Swap the docx picture for a synthetic one of the same size: nothing in the PDF should match it.
    const wordPdf = await simulateWordExport(SAMPLE_DOCX, { platform });
    const { unzipSync, zipSync } = await import("fflate");
    const { syntheticRaster } = await import("../scripts/simulate-word-export");
    const files = unzipSync(SAMPLE_DOCX);
    files["word/media/image1.jpeg"] = (await platform.encodeImage(syntheticRaster(2800, 1867), "image/jpeg", 0.9)) as Uint8Array<ArrayBuffer>;
    const swapped = zipSync(files);

    const { report } = await replaceImagesInPdf(wordPdf, swapped, { platform });
    assert.equal(report.replaced, 0);
    assert.equal(report.unmatched, 1);
  });

  it("keeps an ICC-based colour space from the exported image when components match", async () => {
    const wordPdf = await simulateWordExport(SAMPLE_DOCX, { platform });
    const doc = await PDFDocument.load(wordPdf, { updateMetadata: false });
    const [img] = collectPdfImages(doc);
    const icc = doc.context.flateStream(new Uint8Array(16), { N: 3 });
    const iccRef = doc.context.register(icc);
    img.stream.dict.set(PDFName.of("ColorSpace"), doc.context.obj([PDFName.of("ICCBased"), iccRef]));
    const iccPdf = await doc.save({ useObjectStreams: false });

    const { pdf, report } = await replaceImagesInPdf(iccPdf, SAMPLE_DOCX, { platform });
    assert.equal(report.replaced, 1);
    const out = await PDFDocument.load(pdf, { updateMetadata: false });
    const [replaced] = collectPdfImages(out);
    const cs = replaced.stream.dict.lookup(PDFName.of("ColorSpace"));
    assert.equal(String(cs).startsWith("[ /ICCBased"), true, String(cs));
  });
});
