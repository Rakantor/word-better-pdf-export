/**
 * Promise wrapper around Office.context.document.getFileAsync, which hands the
 * document back in slices of at most 4 MB (Word on Windows/Mac).
 */
export type FileProgress = (received: number, total: number) => void;

const SLICE_SIZE = 4 * 1024 * 1024;

export async function getDocumentFile(fileType: Office.FileType, onProgress?: FileProgress): Promise<Uint8Array> {
  const file = await new Promise<Office.File>((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: SLICE_SIZE }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || `getFileAsync(${fileType}) failed`));
    });
  });

  try {
    const chunks: Uint8Array[] = [];
    let received = 0;
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await new Promise<Office.Slice>((resolve, reject) => {
        file.getSliceAsync(index, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message || `getSliceAsync(${index}) failed`));
        });
      });
      const bytes = toBytes(slice.data);
      chunks.push(bytes);
      received += bytes.length;
      onProgress?.(received, file.size);
    }
    const out = new Uint8Array(received);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.length;
    }
    return out;
  } finally {
    await new Promise<void>((resolve) => file.closeAsync(() => resolve()));
  }
}

/** Slice.data is documented as a "byte array"; in practice a plain number[] on desktop. */
function toBytes(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (Array.isArray(data)) return Uint8Array.from(data as number[]);
  if (typeof data === "string") {
    const out = new Uint8Array(data.length);
    for (let i = 0; i < data.length; i++) out[i] = data.charCodeAt(i) & 0xff;
    return out;
  }
  throw new Error(`Unexpected slice data type: ${Object.prototype.toString.call(data)}`);
}
