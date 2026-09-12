import { PDFArray, PDFDict, PDFDocument, PDFName, PDFObject, PDFRawStream, PDFRef, PDFStream } from "pdf-lib";

/** A dictionary entry (`/XObject << /Im3 12 0 R >>`) that points at an image. */
export interface XObjectSlot {
  dict: PDFDict;
  key: PDFName;
}

export interface PdfImage {
  /** Indirect reference (images embedded by real producers are always indirect). */
  ref: PDFRef;
  stream: PDFRawStream;
  width: number;
  height: number;
  /** Filter names in application order, e.g. ["FlateDecode"] or ["DCTDecode"]. */
  filters: string[];
  bitsPerComponent: number;
  /** Where this image is referenced from. Replacing = rewriting every slot. */
  slots: XObjectSlot[];
  /** 0-based page indexes the image appears on (directly or via form XObjects). */
  pages: number[];
  smaskRef?: PDFRef;
}

/**
 * Find every image XObject reachable from the pages' resources, including images
 * nested inside form XObjects (Word uses those for text boxes and grouped shapes).
 */
export function collectPdfImages(doc: PDFDocument): PdfImage[] {
  const images = new Map<string, PdfImage>();
  const visitedForms = new Set<string>();

  const visitResources = (resources: PDFDict | undefined, pageIndex: number): void => {
    if (!resources) return;
    const xobjects = resources.lookup(PDFName.of("XObject"));
    if (!(xobjects instanceof PDFDict)) return;
    for (const [key, raw] of xobjects.entries()) {
      const ref = raw instanceof PDFRef ? raw : undefined;
      const obj = ref ? doc.context.lookup(ref) : raw;
      if (!(obj instanceof PDFStream)) continue;
      const subtype = obj.dict.lookup(PDFName.of("Subtype"));
      if (subtype === PDFName.of("Image")) {
        if (!ref || !(obj instanceof PDFRawStream)) continue; // direct image objects: not produced by Word
        let entry = images.get(ref.toString());
        if (!entry) {
          entry = describeImage(ref, obj);
          images.set(ref.toString(), entry);
        }
        entry.slots.push({ dict: xobjects, key });
        if (!entry.pages.includes(pageIndex)) entry.pages.push(pageIndex);
      } else if (subtype === PDFName.of("Form")) {
        const formKey = ref ? ref.toString() : `direct-${pageIndex}-${key.toString()}`;
        if (visitedForms.has(formKey)) continue;
        visitedForms.add(formKey);
        const res = obj.dict.lookup(PDFName.of("Resources"));
        visitResources(res instanceof PDFDict ? res : undefined, pageIndex);
      }
    }
  };

  doc.getPages().forEach((page, index) => visitResources(page.node.Resources(), index));
  return Array.from(images.values());
}

function describeImage(ref: PDFRef, stream: PDFRawStream): PdfImage {
  const dict = stream.dict;
  const smask = dict.get(PDFName.of("SMask"));
  return {
    ref,
    stream,
    width: numberOf(dict.lookup(PDFName.of("Width"))) ?? 0,
    height: numberOf(dict.lookup(PDFName.of("Height"))) ?? 0,
    filters: filterNames(dict),
    bitsPerComponent: numberOf(dict.lookup(PDFName.of("BitsPerComponent"))) ?? 8,
    slots: [],
    pages: [],
    smaskRef: smask instanceof PDFRef ? smask : undefined,
  };
}

export function filterNames(dict: PDFDict): string[] {
  const filter = dict.lookup(PDFName.of("Filter"));
  if (filter instanceof PDFName) return [filter.decodeText()];
  if (filter instanceof PDFArray) {
    return filter.asArray().map((f) => (f instanceof PDFName ? f.decodeText() : String(f)));
  }
  return [];
}

export function numberOf(obj: PDFObject | undefined): number | undefined {
  // pdf-lib exposes PDFNumber via asNumber(); avoid importing the class to keep instanceof simple.
  const anyObj = obj as { asNumber?: () => number } | undefined;
  return anyObj && typeof anyObj.asNumber === "function" ? anyObj.asNumber() : undefined;
}

/**
 * Collect every indirect reference used anywhere in the document (dictionaries,
 * arrays, stream dictionaries). Used to make sure an object is orphaned before
 * deleting it.
 */
export function collectReferencedRefs(doc: PDFDocument): Set<string> {
  const seen = new Set<string>();
  const visit = (obj: PDFObject | undefined, depth: number): void => {
    if (!obj || depth > 64) return;
    if (obj instanceof PDFRef) {
      seen.add(obj.toString());
    } else if (obj instanceof PDFDict) {
      for (const [, value] of obj.entries()) visit(value, depth + 1);
    } else if (obj instanceof PDFArray) {
      for (const value of obj.asArray()) visit(value, depth + 1);
    } else if (obj instanceof PDFStream) {
      visit(obj.dict, depth + 1);
    }
  };
  for (const [, obj] of doc.context.enumerateIndirectObjects()) visit(obj, 0);
  visit(doc.context.trailerInfo.Root, 0);
  visit(doc.context.trailerInfo.Info, 0);
  return seen;
}
