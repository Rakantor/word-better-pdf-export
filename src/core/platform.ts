/**
 * Environment-specific services the core needs. The browser implementation uses
 * canvas + DOMParser, the Node implementation (tests / CLI) uses jpeg-js, pngjs
 * and @xmldom/xmldom. Everything else in `src/core` is environment agnostic.
 */

/** 8-bit RGBA pixel buffer, row-major, no padding. */
export interface Raster {
  width: number;
  height: number;
  data: Uint8ClampedArray | Uint8Array;
}

export type EncodableMime = "image/png" | "image/jpeg";

export interface Platform {
  /** Decode an encoded image (JPEG/PNG/GIF/BMP/...) into RGBA. Throws for unsupported formats. */
  decodeImage(bytes: Uint8Array, mime: string): Promise<Raster>;
  /** Encode RGBA to PNG (lossless) or JPEG (`quality` 0..1). */
  encodeImage(raster: Raster, mime: EncodableMime, quality?: number): Promise<Uint8Array>;
  /** Parse an XML string into a DOM `Document`. */
  parseXml(xml: string): Document;
}
