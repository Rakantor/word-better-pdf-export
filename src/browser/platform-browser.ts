import type { EncodableMime, Platform, Raster } from "../core/platform";

/** Platform backed by the WebView2/browser canvas and DOMParser. */
export const browserPlatform: Platform = {
  async decodeImage(bytes, mime) {
    const blob = new Blob([bytes as BlobPart], { type: mime });
    let bitmap: ImageBitmap;
    try {
      // Word ignores EXIF orientation when drawing (it stores rotation on the shape), so must we.
      bitmap = await createImageBitmap(blob, { imageOrientation: "none" } as ImageBitmapOptions);
    } catch {
      bitmap = await createImageBitmap(blob);
    }
    try {
      const canvas = makeCanvas(bitmap.width, bitmap.height);
      const ctx = canvas.getContext("2d", { willReadFrequently: true }) as OffscreenCanvasRenderingContext2D | CanvasRenderingContext2D;
      ctx.drawImage(bitmap, 0, 0);
      const imageData = ctx.getImageData(0, 0, bitmap.width, bitmap.height);
      return { width: bitmap.width, height: bitmap.height, data: imageData.data };
    } finally {
      bitmap.close();
    }
  },

  async encodeImage(raster, mime, quality) {
    const canvas = makeCanvas(raster.width, raster.height);
    const ctx = canvas.getContext("2d") as OffscreenCanvasRenderingContext2D | CanvasRenderingContext2D;
    const imageData = ctx.createImageData(raster.width, raster.height);
    imageData.data.set(raster.data);
    ctx.putImageData(imageData, 0, 0);
    const blob = await canvasToBlob(canvas, mime, quality);
    return new Uint8Array(await blob.arrayBuffer());
  },

  parseXml(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const error = doc.getElementsByTagName("parsererror")[0];
    if (error) throw new Error(`XML parse error: ${error.textContent ?? ""}`);
    return doc;
  },
};

function makeCanvas(width: number, height: number): OffscreenCanvas | HTMLCanvasElement {
  if (typeof OffscreenCanvas !== "undefined") return new OffscreenCanvas(width, height);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function canvasToBlob(canvas: OffscreenCanvas | HTMLCanvasElement, type: EncodableMime, quality?: number): Promise<Blob> {
  if ("convertToBlob" in canvas) return canvas.convertToBlob({ type, quality });
  return new Promise((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("canvas.toBlob failed"))), type, quality),
  );
}

export type { Raster };
