/**
 * Produces a PDF that behaves like Word's "Save as PDF" output for a .docx:
 * every picture is downsampled to `ppi` and recompressed as JPEG at `jpegQuality`
 * before being embedded. Used for tests and for developing without Word.
 */
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { extractDocxImages } from "../src/core/docx";
import type { Platform, Raster } from "../src/core/platform";
import { cropRaster, resampleRaster } from "../src/core/raster";

const EMU_PER_INCH = 914400;

export interface SimulateOptions {
  platform: Platform;
  ppi?: number;
  jpegQuality?: number;
  /** Add a synthetic picture with the same aspect ratio as the first real one (tests matching). */
  addDecoy?: boolean;
  /** Embed pictures untouched instead of downsampling (simulates "do not compress"). */
  passthrough?: boolean;
}

export async function simulateWordExport(docxBytes: Uint8Array, options: SimulateOptions): Promise<Uint8Array> {
  const { platform } = options;
  const ppi = options.ppi ?? 220;
  const quality = options.jpegQuality ?? 0.75;
  const docx = await extractDocxImages(docxBytes, platform);

  const pdf = await PDFDocument.create();
  pdf.setProducer("Simulated Microsoft Word exporter");
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  let page = pdf.addPage([612, 792]);
  let y = 740;
  page.drawText("Simulated Word PDF export", { x: 72, y, size: 18, font });
  y -= 30;

  let decoyDone = false;
  for (const usage of docx.usages) {
    const extent = usage.extentEmu ?? { cx: 4 * EMU_PER_INCH, cy: (4 * EMU_PER_INCH * usage.part.height) / usage.part.width };
    const widthIn = extent.cx / EMU_PER_INCH;
    const heightIn = extent.cy / EMU_PER_INCH;
    const drawW = widthIn * 72;
    const drawH = heightIn * 72;
    if (y - drawH < 72) {
      page = pdf.addPage([612, 792]);
      y = 740;
    }

    let image;
    if (options.passthrough && usage.part.mime === "image/jpeg" && !usage.crop) {
      image = await pdf.embedJpg(usage.part.bytes);
    } else {
      const full = await platform.decodeImage(usage.part.bytes, usage.part.mime);
      const cropped = cropRaster(full, usage.crop);
      // Word resamples the bitmap uniformly to the target ppi (never upsamples).
      const scale = Math.min(1, (widthIn * ppi) / cropped.width, (heightIn * ppi) / cropped.height);
      const small = resampleRaster(cropped, Math.max(1, Math.round(cropped.width * scale)), Math.max(1, Math.round(cropped.height * scale)));
      const jpg = await platform.encodeImage(small, "image/jpeg", quality);
      image = await pdf.embedJpg(jpg);
    }
    page.drawImage(image, { x: 72, y: y - drawH, width: drawW, height: drawH });
    y -= drawH + 12;
    page.drawText(`${usage.name ?? usage.part.path} (${usage.part.width}x${usage.part.height} -> ${image.width}x${image.height})`, {
      x: 72,
      y,
      size: 9,
      font,
      color: rgb(0.3, 0.3, 0.3),
    });
    y -= 24;

    if (options.addDecoy && !decoyDone) {
      decoyDone = true;
      const decoy = syntheticRaster(Math.round(widthIn * ppi), Math.round(heightIn * ppi));
      const jpg = await platform.encodeImage(decoy, "image/jpeg", quality);
      const decoyImage = await pdf.embedJpg(jpg);
      if (y - drawH < 72) {
        page = pdf.addPage([612, 792]);
        y = 740;
      }
      page.drawImage(decoyImage, { x: 72, y: y - drawH, width: drawW, height: drawH });
      y -= drawH + 12;
      page.drawText("decoy picture (same aspect ratio, not in the docx)", { x: 72, y, size: 9, font, color: rgb(0.3, 0.3, 0.3) });
      y -= 24;
    }
  }
  return pdf.save({ useObjectStreams: true });
}

/** Deterministic gradient + stripes; visually unlike a photo. */
export function syntheticRaster(width: number, height: number): Raster {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      const stripe = Math.floor(x / 40) % 2 === 0 ? 40 : 0;
      data[o] = (x / width) * 255;
      data[o + 1] = (y / height) * 255 - stripe;
      data[o + 2] = 128 + stripe;
      data[o + 3] = 255;
    }
  }
  return { width, height, data };
}
