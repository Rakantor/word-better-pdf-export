import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { zlibSync } from "fflate";
import { PDFDocument, PDFName } from "pdf-lib";
import { decodePdfImage } from "../src/core/pdf-decode";
import { collectPdfImages } from "../src/core/pdf-images";
import { nodePlatform } from "../src/node/platform-node";

/** Build a one-page PDF whose only image is a FlateDecode RGB stream, optionally PNG-predicted. */
async function pdfWithFlateImage(width: number, height: number, rgb: Uint8Array, predictor: boolean): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([200, 200]);
  let samples = rgb;
  if (predictor) {
    // PNG "Up" filter (type 2) on every row.
    const rowLen = width * 3;
    samples = new Uint8Array((rowLen + 1) * height);
    for (let y = 0; y < height; y++) {
      samples[y * (rowLen + 1)] = 2;
      for (let i = 0; i < rowLen; i++) {
        const cur = rgb[y * rowLen + i];
        const up = y ? rgb[(y - 1) * rowLen + i] : 0;
        samples[y * (rowLen + 1) + 1 + i] = (cur - up) & 0xff;
      }
    }
  }
  const dict: Parameters<typeof doc.context.stream>[1] = {
    Type: "XObject",
    Subtype: "Image",
    Width: width,
    Height: height,
    ColorSpace: "DeviceRGB",
    BitsPerComponent: 8,
    Filter: "FlateDecode",
  };
  if (predictor) dict!.DecodeParms = { Predictor: 15, Colors: 3, BitsPerComponent: 8, Columns: width };
  const stream = doc.context.stream(zlibSync(samples), dict);
  const ref = doc.context.register(stream);
  page.node.setXObject(PDFName.of("Im1"), ref);
  page.node.set(PDFName.of("Contents"), doc.context.register(doc.context.stream("q 100 0 0 100 50 50 cm /Im1 Do Q")));
  return doc.save({ useObjectStreams: false });
}

function gradient(width: number, height: number): Uint8Array {
  const rgb = new Uint8Array(width * height * 3);
  for (let y = 0; y < height; y++)
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 3;
      rgb[o] = (x * 255) / (width - 1);
      rgb[o + 1] = (y * 255) / (height - 1);
      rgb[o + 2] = 77;
    }
  return rgb;
}

describe("decodePdfImage", () => {
  for (const predictor of [false, true]) {
    it(`decodes FlateDecode RGB samples${predictor ? " with PNG predictors" : ""}`, async () => {
      const rgb = gradient(8, 4);
      const doc = await PDFDocument.load(await pdfWithFlateImage(8, 4, rgb, predictor));
      const [img] = collectPdfImages(doc);
      assert.deepEqual(img.filters, ["FlateDecode"]);
      const raster = await decodePdfImage(doc, img, nodePlatform);
      assert.ok(raster);
      assert.equal(raster.width, 8);
      assert.equal(raster.height, 4);
      for (let p = 0; p < 32; p++) {
        assert.equal(raster.data[p * 4], rgb[p * 3]);
        assert.equal(raster.data[p * 4 + 1], rgb[p * 3 + 1]);
        assert.equal(raster.data[p * 4 + 2], rgb[p * 3 + 2]);
        assert.equal(raster.data[p * 4 + 3], 255);
      }
    });
  }

  it("decodes DCTDecode images and images nested in form XObjects", async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([200, 200]);
    const jpg = await nodePlatform.encodeImage(
      { width: 4, height: 4, data: new Uint8ClampedArray(64).fill(200) },
      "image/jpeg",
      1,
    );
    const image = await doc.embedJpg(jpg.slice()); // pdf-lib ignores byteOffset
    await image.embed();
    // Wrap the image in a form XObject like Word does for text boxes.
    const form = doc.context.stream("q 4 0 0 4 0 0 cm /Im0 Do Q", {
      Type: "XObject",
      Subtype: "Form",
      BBox: [0, 0, 4, 4],
      Resources: { XObject: { Im0: image.ref } },
    });
    const formRef = doc.context.register(form);
    page.node.setXObject(PDFName.of("Fx1"), formRef);
    page.node.set(PDFName.of("Contents"), doc.context.register(doc.context.stream("q 100 0 0 100 50 50 cm /Fx1 Do Q")));
    const loaded = await PDFDocument.load(await doc.save({ useObjectStreams: false }));

    const images = collectPdfImages(loaded);
    assert.equal(images.length, 1);
    assert.deepEqual(images[0].filters, ["DCTDecode"]);
    assert.deepEqual(images[0].pages, [0]);
    const raster = await decodePdfImage(loaded, images[0], nodePlatform);
    assert.ok(raster);
    assert.equal(raster.width, 4);
    assert.ok(Math.abs(raster.data[0] - 200) < 4);
  });
});
