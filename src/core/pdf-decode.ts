import { inflateSync, unzlibSync } from "fflate";
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFObject, PDFRawStream, PDFStream, PDFString } from "pdf-lib";
import { filterNames, numberOf, type PdfImage } from "./pdf-images";
import type { Platform, Raster } from "./platform";

/**
 * Decode an image XObject to RGBA. Supports what Word's exporter emits
 * (DCTDecode JPEGs and FlateDecode samples in Gray/RGB/CMYK/ICC/Indexed, with
 * optional PNG predictors and an SMask). Returns null for anything else.
 */
export async function decodePdfImage(doc: PDFDocument, image: PdfImage, platform: Platform): Promise<Raster | null> {
  const raster = await decodeStreamToRaster(doc, image.stream, platform);
  if (!raster) return null;
  if (image.smaskRef) {
    const smask = doc.context.lookup(image.smaskRef);
    if (smask instanceof PDFRawStream) {
      const alpha = await decodeStreamToRaster(doc, smask, platform);
      if (alpha) applyAlpha(raster, alpha);
    }
  }
  return raster;
}

async function decodeStreamToRaster(doc: PDFDocument, stream: PDFRawStream, platform: Platform): Promise<Raster | null> {
  const dict = stream.dict;
  const width = numberOf(dict.lookup(PDFName.of("Width"))) ?? 0;
  const height = numberOf(dict.lookup(PDFName.of("Height"))) ?? 0;
  if (!width || !height) return null;

  const filters = filterNames(dict);
  const parms = decodeParms(dict, filters.length);
  let data: Uint8Array = stream.contents;
  for (let i = 0; i < filters.length; i++) {
    const filter = filters[i];
    if (filter === "DCTDecode" || filter === "DCT") {
      if (i !== filters.length - 1) return null;
      try {
        return await platform.decodeImage(data, "image/jpeg");
      } catch {
        return null;
      }
    }
    if (filter === "FlateDecode" || filter === "Fl") {
      data = inflate(data);
      data = applyPredictor(data, parms[i], width);
    } else {
      return null; // JPX, CCITT, LZW, ASCII filters: not something Word emits for pictures
    }
  }

  const bpc = numberOf(dict.lookup(PDFName.of("BitsPerComponent"))) ?? 8;
  const cs = describeColorSpace(doc, dict.lookup(PDFName.of("ColorSpace")));
  if (!cs) return null;
  const decodeArr = dict.lookup(PDFName.of("Decode"));
  const inverted = decodeArr instanceof PDFArray && numberOf(decodeArr.lookup(0)) === 1;
  return samplesToRgba(data, width, height, bpc, cs, inverted);
}

interface ColorSpaceInfo {
  kind: "gray" | "rgb" | "cmyk" | "indexed";
  components: number;
  /** For indexed: base colour space and lookup table. */
  base?: ColorSpaceInfo;
  lookup?: Uint8Array;
}

function describeColorSpace(doc: PDFDocument, cs: PDFObject | undefined): ColorSpaceInfo | null {
  if (cs instanceof PDFName) {
    switch (cs.decodeText()) {
      case "DeviceGray":
      case "CalGray":
      case "G":
        return { kind: "gray", components: 1 };
      case "DeviceRGB":
      case "CalRGB":
      case "RGB":
        return { kind: "rgb", components: 3 };
      case "DeviceCMYK":
      case "CMYK":
        return { kind: "cmyk", components: 4 };
      default:
        return null;
    }
  }
  if (cs instanceof PDFArray && cs.size() > 0) {
    const family = cs.lookup(0);
    const name = family instanceof PDFName ? family.decodeText() : "";
    if (name === "ICCBased") {
      const iccStream = cs.lookup(1);
      const n = iccStream instanceof PDFStream ? numberOf(iccStream.dict.lookup(PDFName.of("N"))) : undefined;
      if (n === 1) return { kind: "gray", components: 1 };
      if (n === 3) return { kind: "rgb", components: 3 };
      if (n === 4) return { kind: "cmyk", components: 4 };
      return null;
    }
    if (name === "CalRGB" || name === "Lab") return { kind: "rgb", components: 3 };
    if (name === "CalGray") return { kind: "gray", components: 1 };
    if (name === "Indexed" || name === "I") {
      const base = describeColorSpace(doc, cs.lookup(1));
      const lookupObj = cs.lookup(3);
      let lookup: Uint8Array | undefined;
      if (lookupObj instanceof PDFHexString || lookupObj instanceof PDFString) lookup = lookupObj.asBytes();
      else if (lookupObj instanceof PDFRawStream) {
        lookup = lookupObj.contents;
        const f = filterNames(lookupObj.dict);
        if (f.includes("FlateDecode")) lookup = inflate(lookup);
        else if (f.length) return null;
      }
      if (!base || !lookup || base.kind === "indexed") return null;
      return { kind: "indexed", components: 1, base, lookup };
    }
  }
  return null;
}

function decodeParms(dict: PDFDict, count: number): (PDFDict | undefined)[] {
  const parms = dict.lookup(PDFName.of("DecodeParms")) ?? dict.lookup(PDFName.of("DP"));
  const out: (PDFDict | undefined)[] = new Array(count).fill(undefined);
  if (parms instanceof PDFDict) out[0] = parms;
  else if (parms instanceof PDFArray) {
    for (let i = 0; i < Math.min(count, parms.size()); i++) {
      const p = parms.lookup(i);
      out[i] = p instanceof PDFDict ? p : undefined;
    }
  }
  return out;
}

function inflate(data: Uint8Array): Uint8Array {
  try {
    return unzlibSync(data);
  } catch {
    // Some producers write raw deflate or trailing garbage; try raw inflate on the payload.
    return inflateSync(data.subarray(2));
  }
}

/** Undo PNG (10-15) / TIFF (2) predictors as described in the PDF spec for FlateDecode. */
function applyPredictor(data: Uint8Array, parms: PDFDict | undefined, imageWidth: number): Uint8Array {
  if (!parms) return data;
  const predictor = numberOf(parms.lookup(PDFName.of("Predictor"))) ?? 1;
  if (predictor <= 1) return data;
  const colors = numberOf(parms.lookup(PDFName.of("Colors"))) ?? 1;
  const bpc = numberOf(parms.lookup(PDFName.of("BitsPerComponent"))) ?? 8;
  const columns = numberOf(parms.lookup(PDFName.of("Columns"))) ?? imageWidth;
  const bpp = Math.ceil((colors * bpc) / 8);
  const rowLen = Math.ceil((colors * bpc * columns) / 8);

  if (predictor === 2) {
    if (bpc !== 8) return data;
    const rows = Math.floor(data.length / rowLen);
    for (let r = 0; r < rows; r++) {
      const off = r * rowLen;
      for (let i = bpp; i < rowLen; i++) data[off + i] = (data[off + i] + data[off + i - bpp]) & 0xff;
    }
    return data;
  }

  const rows = Math.floor(data.length / (rowLen + 1));
  const out = new Uint8Array(rows * rowLen);
  const prev = new Uint8Array(rowLen);
  for (let r = 0; r < rows; r++) {
    const ft = data[r * (rowLen + 1)];
    const src = data.subarray(r * (rowLen + 1) + 1, (r + 1) * (rowLen + 1));
    const dst = out.subarray(r * rowLen, (r + 1) * rowLen);
    for (let i = 0; i < rowLen; i++) {
      const a = i >= bpp ? dst[i - bpp] : 0;
      const b = prev[i];
      const c = i >= bpp ? prev[i - bpp] : 0;
      let v = src[i];
      switch (ft) {
        case 1:
          v += a;
          break;
        case 2:
          v += b;
          break;
        case 3:
          v += (a + b) >> 1;
          break;
        case 4: {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
          break;
        }
        default:
          break;
      }
      dst[i] = v & 0xff;
    }
    prev.set(dst);
  }
  return out;
}

function samplesToRgba(
  data: Uint8Array,
  width: number,
  height: number,
  bpc: number,
  cs: ColorSpaceInfo,
  inverted: boolean,
): Raster | null {
  const out = new Uint8ClampedArray(width * height * 4);
  const comps = cs.components;
  const rowBits = width * comps * bpc;
  const rowBytes = Math.ceil(rowBits / 8);
  if (data.length < rowBytes * height) return null;
  const maxVal = (1 << bpc) - 1;

  const readSample = (row: number, index: number): number => {
    // index = sample index within the row (pixel * comps + component)
    const rowOff = row * rowBytes;
    switch (bpc) {
      case 8:
        return data[rowOff + index];
      case 16:
        return data[rowOff + index * 2];
      case 1:
      case 2:
      case 4: {
        const bit = index * bpc;
        const byte = data[rowOff + (bit >> 3)];
        const shift = 8 - bpc - (bit & 7);
        return (byte >> shift) & maxVal;
      }
      default:
        return 0;
    }
  };
  const scale = bpc === 16 ? 1 : 255 / maxVal;

  let o = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++, o += 4) {
      const s = x * comps;
      let r: number;
      let g: number;
      let b: number;
      switch (cs.kind) {
        case "gray": {
          let v = readSample(y, s) * scale;
          if (inverted) v = 255 - v;
          r = g = b = v;
          break;
        }
        case "rgb":
          r = readSample(y, s) * scale;
          g = readSample(y, s + 1) * scale;
          b = readSample(y, s + 2) * scale;
          break;
        case "cmyk": {
          const c = readSample(y, s) * scale;
          const m = readSample(y, s + 1) * scale;
          const yy = readSample(y, s + 2) * scale;
          const k = readSample(y, s + 3) * scale;
          r = 255 - Math.min(255, c + k);
          g = 255 - Math.min(255, m + k);
          b = 255 - Math.min(255, yy + k);
          break;
        }
        case "indexed": {
          const idx = readSample(y, s);
          const base = cs.base!;
          const lut = cs.lookup!;
          const n = base.components;
          const off = idx * n;
          if (base.kind === "gray") r = g = b = lut[off] ?? 0;
          else if (base.kind === "rgb") {
            r = lut[off] ?? 0;
            g = lut[off + 1] ?? 0;
            b = lut[off + 2] ?? 0;
          } else {
            const k = lut[off + 3] ?? 0;
            r = 255 - Math.min(255, (lut[off] ?? 0) + k);
            g = 255 - Math.min(255, (lut[off + 1] ?? 0) + k);
            b = 255 - Math.min(255, (lut[off + 2] ?? 0) + k);
          }
          break;
        }
      }
      out[o] = r;
      out[o + 1] = g;
      out[o + 2] = b;
      out[o + 3] = 255;
    }
  }
  return { width, height, data: out };
}

/** Copy the gray channel of `alpha` (resampled nearest-neighbour) into `raster`'s alpha. */
function applyAlpha(raster: Raster, alpha: Raster): void {
  for (let y = 0; y < raster.height; y++) {
    const ay = Math.min(alpha.height - 1, Math.floor((y * alpha.height) / raster.height));
    for (let x = 0; x < raster.width; x++) {
      const ax = Math.min(alpha.width - 1, Math.floor((x * alpha.width) / raster.width));
      raster.data[(y * raster.width + x) * 4 + 3] = alpha.data[(ay * alpha.width + ax) * 4];
    }
  }
}
