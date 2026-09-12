/** Cheap header-only dimension parsing so we can pre-filter candidates without decoding. */
export interface ImageSize {
  width: number;
  height: number;
}

export function readImageSize(bytes: Uint8Array): ImageSize | null {
  return readPng(bytes) ?? readJpeg(bytes) ?? readGif(bytes) ?? readBmp(bytes) ?? null;
}

function readPng(b: Uint8Array): ImageSize | null {
  if (b.length < 24 || b[0] !== 0x89 || b[1] !== 0x50 || b[2] !== 0x4e || b[3] !== 0x47) return null;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  return { width: dv.getUint32(16), height: dv.getUint32(20) };
}

function readJpeg(b: Uint8Array): ImageSize | null {
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = b[i + 1];
    if (marker === 0xff) {
      i++;
      continue;
    }
    // Standalone markers without length.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    const len = dv.getUint16(i + 2);
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      return { height: dv.getUint16(i + 5), width: dv.getUint16(i + 7) };
    }
    if (marker === 0xda) break; // start of scan; no SOF found before it
    i += 2 + len;
  }
  return null;
}

function readGif(b: Uint8Array): ImageSize | null {
  if (b.length < 10 || b[0] !== 0x47 || b[1] !== 0x49 || b[2] !== 0x46) return null;
  return { width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
}

function readBmp(b: Uint8Array): ImageSize | null {
  if (b.length < 26 || b[0] !== 0x42 || b[1] !== 0x4d) return null;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  return { width: dv.getInt32(18, true), height: Math.abs(dv.getInt32(22, true)) };
}

export function mimeFromExtension(path: string): string {
  const ext = path.toLowerCase().split(".").pop() ?? "";
  switch (ext) {
    case "jpg":
    case "jpeg":
    case "jpe":
    case "jfif":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
    case "dib":
      return "image/bmp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "webp":
      return "image/webp";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

/** Sniff the real type from magic bytes; docx content types are occasionally wrong. */
export function sniffMime(bytes: Uint8Array, fallback: string): string {
  if (bytes.length >= 4) {
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return "image/gif";
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
    if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46) return "image/webp";
    if ((bytes[0] === 0x49 && bytes[1] === 0x49) || (bytes[0] === 0x4d && bytes[1] === 0x4d)) return "image/tiff";
  }
  return fallback;
}
