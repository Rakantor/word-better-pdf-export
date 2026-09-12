import { PDFArray, PDFDocument, PDFImage, PDFName, PDFRawStream, PDFRef, PDFStream } from "pdf-lib";
import { extractDocxImages, type DocxImagePart, type DocxImageUsage } from "./docx";
import { decodePdfImage } from "./pdf-decode";
import { collectPdfImages, collectReferencedRefs, numberOf, type PdfImage } from "./pdf-images";
import type { Platform, Raster } from "./platform";
import { aspectMatches, cropRaster, croppedSize, grayThumbnail, isNoopCrop, similarity, type CropFractions } from "./raster";

export interface ReplaceOptions {
  platform: Platform;
  /** Minimum normalised cross-correlation (−1..1) to accept a visual match. Default 0.9. */
  minSimilarity?: number;
  /** Relative aspect-ratio tolerance for candidate pre-filtering. Default 0.03. */
  aspectTolerance?: number;
  /** JPEG quality (0..1) used when a cropped JPEG has to be re-encoded. Default 0.95. */
  jpegQuality?: number;
  /** Thumbnail edge length used for matching. Default 32. */
  thumbnailSize?: number;
  /** Remove the replaced (compressed) image objects from the file. Default true. */
  removeReplaced?: boolean;
  log?: (message: string) => void;
}

export type ImageOutcome =
  | { status: "replaced"; source: string; crop?: CropFractions; similarity: number; from: string; to: string; savedBytes: number }
  | { status: "skipped"; source?: string; reason: string; similarity?: number }
  | { status: "unmatched"; reason: string; bestSimilarity?: number };

export interface ImageIdentity {
  /** Object reference of the image in the input PDF, e.g. "12 0 R". */
  ref: string;
  pages: number[];
  width: number;
  height: number;
  filters: string[];
}

export type ImageResult = ImageIdentity & ImageOutcome;

export interface ReplaceReport {
  docxImageParts: number;
  docxImageUsages: number;
  pdfImages: number;
  replaced: number;
  skipped: number;
  unmatched: number;
  results: ImageResult[];
  inputBytes: number;
  outputBytes: number;
}

interface Candidate {
  part: DocxImagePart;
  crop?: CropFractions;
  width: number;
  height: number;
  usages: DocxImageUsage[];
  thumb?: Float32Array | null;
  raster?: Raster;
}

interface Embedded {
  image: PDFImage;
  bytes: number;
  description: string;
}

/**
 * Take the PDF produced by Word's exporter plus the .docx it came from and swap
 * every downsampled/recompressed picture in the PDF for the original image file.
 */
export async function replaceImagesInPdf(
  pdfBytes: Uint8Array,
  docxBytes: Uint8Array,
  options: ReplaceOptions,
): Promise<{ pdf: Uint8Array; report: ReplaceReport }> {
  const { platform } = options;
  const minSimilarity = options.minSimilarity ?? 0.9;
  const aspectTolerance = options.aspectTolerance ?? 0.03;
  const jpegQuality = options.jpegQuality ?? 0.95;
  const thumbSize = options.thumbnailSize ?? 32;
  const removeReplaced = options.removeReplaced ?? true;
  const log = options.log ?? (() => undefined);

  const docx = await extractDocxImages(docxBytes, platform);
  log(`docx: ${docx.parts.length} image part(s), ${docx.usages.length} usage(s)`);

  const doc = await PDFDocument.load(pdfBytes, { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
  const pdfImages = collectPdfImages(doc);
  log(`pdf: ${doc.getPageCount()} page(s), ${pdfImages.length} image XObject(s)`);

  const candidates = buildCandidates(docx.usages);
  const results: ImageResult[] = [];
  const embeddedCache = new Map<string, Embedded>();
  const replacedRefs: PdfImage[] = [];

  for (const img of pdfImages) {
    const base: ImageIdentity = { ref: img.ref.toString(), pages: img.pages, width: img.width, height: img.height, filters: img.filters };
    const outcome = await matchAndReplace(img);
    results.push({ ...base, ...outcome });
    log(`${base.ref} ${img.width}x${img.height} [${img.filters.join("+") || "raw"}] -> ${describe(outcome)}`);
  }

  if (removeReplaced && replacedRefs.length) {
    // Only drop objects nothing points at any more (an image may also be used by
    // annotations or other structures we did not rewrite).
    let referenced = collectReferencedRefs(doc);
    const deleted: PdfImage[] = [];
    for (const img of replacedRefs) {
      if (!referenced.has(img.ref.toString())) {
        doc.context.delete(img.ref);
        deleted.push(img);
      }
    }
    if (deleted.some((img) => img.smaskRef)) {
      referenced = collectReferencedRefs(doc);
      for (const img of deleted) {
        if (img.smaskRef && !referenced.has(img.smaskRef.toString())) doc.context.delete(img.smaskRef);
      }
    }
  }

  const pdf = await doc.save({ useObjectStreams: true, addDefaultPage: false, updateFieldAppearances: false });
  const report: ReplaceReport = {
    docxImageParts: docx.parts.length,
    docxImageUsages: docx.usages.length,
    pdfImages: pdfImages.length,
    replaced: results.filter((r) => r.status === "replaced").length,
    skipped: results.filter((r) => r.status === "skipped").length,
    unmatched: results.filter((r) => r.status === "unmatched").length,
    results,
    inputBytes: pdfBytes.length,
    outputBytes: pdf.length,
  };
  return { pdf, report };

  // -------------------------------------------------------------------------

  async function matchAndReplace(img: PdfImage): Promise<ImageOutcome> {
    if (!img.width || !img.height) return { status: "unmatched", reason: "image has no dimensions" };

    const byAspect = candidates.filter((c) => aspectMatches(img.width, img.height, c.width, c.height, aspectTolerance));
    if (!byAspect.length) return { status: "unmatched", reason: "no docx image with this aspect ratio" };

    const pdfRaster = await decodePdfImage(doc, img, platform);
    let best: Candidate | undefined;
    let bestScore = -Infinity;
    if (pdfRaster) {
      const pdfThumb = grayThumbnail(pdfRaster, thumbSize);
      for (const c of byAspect) {
        const thumb = await candidateThumb(c);
        if (!thumb) continue;
        const score = similarity(pdfThumb, thumb);
        if (score > bestScore) {
          bestScore = score;
          best = c;
        }
      }
      if (!best || bestScore < minSimilarity) {
        return { status: "unmatched", reason: `best visual similarity ${fmt(bestScore)} below ${minSimilarity}`, bestSimilarity: bestScore };
      }
    } else {
      // Could not decode the PDF image (unusual filter). Only accept an unambiguous aspect match.
      const distinct = new Set(byAspect.map((c) => candidateKey(c)));
      if (distinct.size !== 1) return { status: "unmatched", reason: "undecodable image with several aspect-ratio candidates" };
      best = byAspect[0];
      bestScore = NaN;
    }

    const source = describeCandidate(best);
    const blockers = new Set(best.usages.flatMap((u) => u.blockers));
    if (blockers.size) return { status: "skipped", source, reason: Array.from(blockers).join(", "), similarity: bestScore };

    if (isNoopCrop(best.crop) && bytesEqual(img.stream.contents, best.part.bytes)) {
      return { status: "skipped", source, reason: "already the original bytes", similarity: bestScore };
    }
    const losslessPdf = img.filters.every((f) => f === "FlateDecode" || f === "Fl");
    const losslessSource = best.part.mime === "image/png" || best.part.mime === "image/bmp" || best.part.mime === "image/gif";
    if (losslessPdf && losslessSource && img.width >= best.width && img.height >= best.height) {
      return { status: "skipped", source, reason: "already lossless at full resolution", similarity: bestScore };
    }

    let embedded: Embedded;
    try {
      embedded = await embedCandidate(best);
    } catch (err) {
      return { status: "skipped", source, reason: `could not embed original: ${(err as Error).message}`, similarity: bestScore };
    }
    await adoptColorSpace(img, embedded.image);
    for (const slot of img.slots) slot.dict.set(slot.key, embedded.image.ref);
    replacedRefs.push(img);
    return {
      status: "replaced",
      source,
      crop: best.crop,
      similarity: bestScore,
      from: `${img.width}x${img.height} ${img.filters.join("+") || "raw"}`,
      to: `${embedded.image.width}x${embedded.image.height} ${embedded.description}`,
      savedBytes: img.stream.contents.length - embedded.bytes,
    };
  }

  async function candidateThumb(c: Candidate): Promise<Float32Array | null> {
    if (c.thumb !== undefined) return c.thumb;
    try {
      const raster = await candidateRaster(c);
      c.thumb = grayThumbnail(raster, thumbSize);
    } catch (err) {
      log(`cannot decode ${describeCandidate(c)}: ${(err as Error).message}`);
      c.thumb = null;
    }
    return c.thumb;
  }

  async function candidateRaster(c: Candidate): Promise<Raster> {
    if (!c.raster) {
      const full = await platform.decodeImage(c.part.bytes, c.part.mime);
      c.raster = cropRaster(full, c.crop);
    }
    return c.raster;
  }

  async function embedCandidate(c: Candidate): Promise<Embedded> {
    const key = candidateKey(c);
    const cached = embeddedCache.get(key);
    if (cached) return cached;

    let result: Embedded;
    if (isNoopCrop(c.crop) && c.part.mime === "image/jpeg") {
      result = { image: await doc.embedJpg(standalone(c.part.bytes)), bytes: c.part.bytes.length, description: "original JPEG" };
    } else if (isNoopCrop(c.crop) && c.part.mime === "image/png") {
      result = { image: await doc.embedPng(standalone(c.part.bytes)), bytes: c.part.bytes.length, description: "original PNG (lossless)" };
    } else {
      // Cropped, or a format PDF cannot carry directly (GIF/BMP/TIFF/WebP): re-encode from pixels.
      const raster = await candidateRaster(c);
      if (c.part.mime === "image/jpeg") {
        const bytes = await platform.encodeImage(raster, "image/jpeg", jpegQuality);
        result = { image: await doc.embedJpg(standalone(bytes)), bytes: bytes.length, description: `cropped JPEG q${Math.round(jpegQuality * 100)}` };
      } else {
        const bytes = await platform.encodeImage(raster, "image/png");
        result = { image: await doc.embedPng(standalone(bytes)), bytes: bytes.length, description: isNoopCrop(c.crop) ? "PNG (lossless)" : "cropped PNG (lossless)" };
      }
    }
    await result.image.embed();
    embeddedCache.set(key, result);
    return result;
  }

  /**
   * Keep Word's ICC-based colour space when the component count matches, so the
   * restored picture renders with the same colour interpretation as before.
   */
  async function adoptColorSpace(oldImage: PdfImage, newImage: PDFImage): Promise<void> {
    const oldCs = oldImage.stream.dict.get(PDFName.of("ColorSpace"));
    const resolved = oldCs instanceof PDFRef ? doc.context.lookup(oldCs) : oldCs;
    if (!(resolved instanceof PDFArray) || resolved.lookup(0) !== PDFName.of("ICCBased")) return;
    const icc = resolved.lookup(1);
    const n = icc instanceof PDFStream ? numberOf(icc.dict.lookup(PDFName.of("N"))) : undefined;
    const newStream = doc.context.lookup(newImage.ref);
    if (!(newStream instanceof PDFRawStream)) return;
    const newCs = newStream.dict.lookup(PDFName.of("ColorSpace"));
    const newN = newCs === PDFName.of("DeviceRGB") ? 3 : newCs === PDFName.of("DeviceGray") ? 1 : newCs === PDFName.of("DeviceCMYK") ? 4 : 0;
    if (n && n === newN && oldCs) newStream.dict.set(PDFName.of("ColorSpace"), oldCs);
  }
}

// ---------------------------------------------------------------------------

/**
 * Group usages into distinct (part, crop) candidates. For cropped usages the
 * uncropped part is added as well: it is not documented whether Word crops the
 * bitmap or clips the full image, so both are offered and the match decides.
 */
function buildCandidates(usages: DocxImageUsage[]): Candidate[] {
  const map = new Map<string, Candidate>();
  const add = (part: DocxImagePart, crop: CropFractions | undefined, usage: DocxImageUsage | null): void => {
    const key = `${part.path}|${cropKey(crop)}`;
    let c = map.get(key);
    if (!c) {
      const size = croppedSize(part.width, part.height, crop);
      c = { part, crop: isNoopCrop(crop) ? undefined : crop, width: size.width, height: size.height, usages: [] };
      map.set(key, c);
    }
    if (usage) c.usages.push(usage);
  };
  for (const u of usages) {
    add(u.part, u.crop, u);
    if (!isNoopCrop(u.crop)) add(u.part, undefined, { ...u, crop: undefined, blockers: u.blockers.filter((b) => !b.includes("crop")) });
  }
  return Array.from(map.values());
}

function cropKey(crop: CropFractions | undefined): string {
  return isNoopCrop(crop) ? "full" : `${crop!.left},${crop!.top},${crop!.right},${crop!.bottom}`;
}

function candidateKey(c: Candidate): string {
  return `${c.part.path}|${cropKey(c.crop)}`;
}

function describeCandidate(c: Candidate): string {
  const name = c.usages.map((u) => u.name).find(Boolean);
  const crop = isNoopCrop(c.crop) ? "" : " (cropped)";
  return `${c.part.path}${crop}${name ? ` "${name}"` : ""}`;
}

function describe(o: ImageOutcome): string {
  switch (o.status) {
    case "replaced":
      return `replaced with ${o.source}: ${o.from} -> ${o.to} (similarity ${fmt(o.similarity)})`;
    case "skipped":
      return `skipped (${o.reason})${o.source ? ` [${o.source}]` : ""}`;
    case "unmatched":
      return `unmatched (${o.reason})`;
  }
}

function fmt(n: number): string {
  return Number.isFinite(n) ? n.toFixed(3) : "n/a";
}

/**
 * pdf-lib reads `new DataView(bytes.buffer)` without honouring `byteOffset`, so a
 * view into a larger buffer (e.g. a pooled Node Buffer) must be copied first.
 */
function standalone(bytes: Uint8Array): Uint8Array {
  return bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes : bytes.slice();
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}
