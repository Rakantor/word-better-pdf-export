/**
 * Getting a file out of a task pane: prefer the File System Access API (WebView2
 * supports it and it gives a proper "Save as" dialog), fall back to a download link.
 */

interface FileSystemWritableFileStreamLike {
  write(data: Uint8Array | Blob): Promise<void>;
  close(): Promise<void>;
}
interface FileSystemFileHandleLike {
  name: string;
  createWritable(): Promise<FileSystemWritableFileStreamLike>;
}
type ShowSaveFilePicker = (options: {
  suggestedName?: string;
  types?: { description: string; accept: Record<string, string[]> }[];
}) => Promise<FileSystemFileHandleLike>;

export interface SaveTarget {
  name: string;
  write(bytes: Uint8Array): Promise<void>;
}

export function supportsSavePicker(): boolean {
  return typeof (window as unknown as { showSaveFilePicker?: unknown }).showSaveFilePicker === "function";
}

/**
 * Ask the user where to save. Must be called from a user gesture (click handler)
 * before any long-running work, otherwise the browser rejects the request.
 * Returns null when the user cancels.
 */
export async function pickSaveTarget(suggestedName: string): Promise<SaveTarget | null> {
  const picker = (window as unknown as { showSaveFilePicker: ShowSaveFilePicker }).showSaveFilePicker;
  try {
    const handle = await picker({
      suggestedName,
      types: [{ description: "PDF document", accept: { "application/pdf": [".pdf"] } }],
    });
    return {
      name: handle.name,
      async write(bytes) {
        const writable = await handle.createWritable();
        await writable.write(bytes);
        await writable.close();
      },
    };
  } catch (err) {
    if ((err as Error).name === "AbortError") return null;
    throw err;
  }
}

export function downloadBytes(bytes: Uint8Array, name: string): void {
  const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    a.remove();
    URL.revokeObjectURL(url);
  }, 60_000);
}

/** Derive "<document name>.pdf" from Office.context.document.url (a path or URL on desktop). */
export function suggestedPdfName(documentUrl: string | null | undefined): string {
  if (!documentUrl) return "document.pdf";
  let last = documentUrl.split(/[\\/]/).pop() || "document";
  try {
    last = decodeURIComponent(last);
  } catch {
    // keep as-is
  }
  return `${last.replace(/\.(docx|docm|dotx|dotm|doc|rtf|odt)$/i, "") || "document"}.pdf`;
}
