import type { Raster } from "./platform";

/** Crop rectangle expressed as fractions (0..1) of the source removed from each side. */
export interface CropFractions {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export function isNoopCrop(c: CropFractions | undefined): boolean {
  return !c || (c.left === 0 && c.top === 0 && c.right === 0 && c.bottom === 0);
}

/** Pixel dimensions of `size` after applying `crop` (mirrors how Word crops with a:srcRect). */
export function croppedSize(width: number, height: number, crop?: CropFractions): { width: number; height: number } {
  if (isNoopCrop(crop)) return { width, height };
  const c = crop!;
  const x0 = Math.round(c.left * width);
  const x1 = Math.round(width - c.right * width);
  const y0 = Math.round(c.top * height);
  const y1 = Math.round(height - c.bottom * height);
  return { width: Math.max(1, x1 - x0), height: Math.max(1, y1 - y0) };
}

export function cropRaster(src: Raster, crop?: CropFractions): Raster {
  if (isNoopCrop(crop)) return src;
  const c = crop!;
  const x0 = Math.min(src.width - 1, Math.max(0, Math.round(c.left * src.width)));
  const y0 = Math.min(src.height - 1, Math.max(0, Math.round(c.top * src.height)));
  const x1 = Math.max(x0 + 1, Math.round(src.width - c.right * src.width));
  const y1 = Math.max(y0 + 1, Math.round(src.height - c.bottom * src.height));
  const width = x1 - x0;
  const height = y1 - y0;
  const out = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y++) {
    const srcStart = ((y + y0) * src.width + x0) * 4;
    out.set(src.data.subarray(srcStart, srcStart + width * 4), y * width * 4);
  }
  return { width, height, data: out };
}

/**
 * Box-filter downsample to a `size`x`size` grayscale thumbnail (alpha composited
 * over white), then normalise to zero mean / unit variance so two thumbnails can
 * be compared with normalised cross-correlation.
 */
export function grayThumbnail(raster: Raster, size = 32): Float32Array {
  const { width, height, data } = raster;
  const sums = new Float64Array(size * size);
  const counts = new Uint32Array(size * size);
  for (let y = 0; y < height; y++) {
    const ty = Math.min(size - 1, Math.floor((y * size) / height));
    let i = y * width * 4;
    for (let x = 0; x < width; x++, i += 4) {
      const tx = Math.min(size - 1, Math.floor((x * size) / width));
      const a = data[i + 3] / 255;
      const r = data[i] * a + 255 * (1 - a);
      const g = data[i + 1] * a + 255 * (1 - a);
      const b = data[i + 2] * a + 255 * (1 - a);
      const idx = ty * size + tx;
      sums[idx] += 0.299 * r + 0.587 * g + 0.114 * b;
      counts[idx]++;
    }
  }
  const thumb = new Float32Array(size * size);
  let mean = 0;
  for (let i = 0; i < thumb.length; i++) {
    thumb[i] = counts[i] ? sums[i] / counts[i] : 255;
    mean += thumb[i];
  }
  mean /= thumb.length;
  let variance = 0;
  for (let i = 0; i < thumb.length; i++) {
    thumb[i] -= mean;
    variance += thumb[i] * thumb[i];
  }
  const std = Math.sqrt(variance / thumb.length);
  if (std > 1e-6) for (let i = 0; i < thumb.length; i++) thumb[i] /= std;
  else thumb.fill(0);
  return thumb;
}

/** Normalised cross-correlation of two thumbnails produced by `grayThumbnail` (-1..1). */
export function similarity(a: Float32Array, b: Float32Array): number {
  if (a.length !== b.length) throw new Error("thumbnail size mismatch");
  let dot = 0;
  for (let i = 0; i < a.length; i++) dot += a[i] * b[i];
  return dot / a.length;
}

/** True when two aspect ratios agree within `tolerance` (relative) or a couple of pixels. */
export function aspectMatches(
  w1: number,
  h1: number,
  w2: number,
  h2: number,
  tolerance = 0.03,
): boolean {
  const a1 = w1 / h1;
  const a2 = w2 / h2;
  if (Math.abs(a1 / a2 - 1) <= tolerance) return true;
  // Rounding slack for small images: predicted height of image 1 at image 2's width.
  return Math.abs(h1 - (w1 * h2) / w2) <= 2;
}

/** Box-filter resample to an arbitrary size (used to simulate Word's downsampling in tests). */
export function resampleRaster(src: Raster, width: number, height: number): Raster {
  const out = new Uint8ClampedArray(width * height * 4);
  const acc = new Float64Array(width * height * 4);
  const counts = new Uint32Array(width * height);
  for (let y = 0; y < src.height; y++) {
    const ty = Math.min(height - 1, Math.floor((y * height) / src.height));
    let i = y * src.width * 4;
    for (let x = 0; x < src.width; x++, i += 4) {
      const tx = Math.min(width - 1, Math.floor((x * width) / src.width));
      const o = (ty * width + tx) * 4;
      acc[o] += src.data[i];
      acc[o + 1] += src.data[i + 1];
      acc[o + 2] += src.data[i + 2];
      acc[o + 3] += src.data[i + 3];
      counts[ty * width + tx]++;
    }
  }
  for (let p = 0; p < width * height; p++) {
    const n = counts[p] || 1;
    out[p * 4] = acc[p * 4] / n;
    out[p * 4 + 1] = acc[p * 4 + 1] / n;
    out[p * 4 + 2] = acc[p * 4 + 2] / n;
    out[p * 4 + 3] = acc[p * 4 + 3] / n;
  }
  return { width, height, data: out };
}
