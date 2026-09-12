import { DOMParser } from "@xmldom/xmldom";
import jpeg from "jpeg-js";
import { PNG } from "pngjs";
import type { Platform } from "../core/platform";

/** Platform for Node (tests and the CLI): pure-JS JPEG/PNG codecs and xmldom. */
export const nodePlatform: Platform = {
  async decodeImage(bytes, mime) {
    if (mime === "image/jpeg") {
      const img = jpeg.decode(bytes, { useTArray: true, formatAsRGBA: true, maxMemoryUsageInMB: 4096, maxResolutionInMP: 400 });
      return { width: img.width, height: img.height, data: img.data };
    }
    if (mime === "image/png") {
      const png = PNG.sync.read(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
      return { width: png.width, height: png.height, data: new Uint8Array(png.data.buffer, png.data.byteOffset, png.data.byteLength) };
    }
    throw new Error(`Node platform cannot decode ${mime}`);
  },

  async encodeImage(raster, mime, quality = 0.95) {
    const data = Buffer.from(raster.data.buffer, raster.data.byteOffset, raster.data.byteLength);
    if (mime === "image/jpeg") {
      const out = jpeg.encode({ data, width: raster.width, height: raster.height }, Math.round(quality * 100));
      return new Uint8Array(out.data.buffer, out.data.byteOffset, out.data.byteLength);
    }
    const png = new PNG({ width: raster.width, height: raster.height });
    png.data = data;
    const out = PNG.sync.write(png);
    return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
  },

  parseXml(xml) {
    return new DOMParser().parseFromString(xml, "application/xml") as unknown as Document;
  },
};
