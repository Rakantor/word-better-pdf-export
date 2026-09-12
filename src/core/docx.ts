import { strFromU8, unzipSync } from "fflate";
import { mimeFromExtension, readImageSize, sniffMime } from "./image-size";
import type { Platform } from "./platform";
import type { CropFractions } from "./raster";

const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  a14: "http://schemas.microsoft.com/office/drawing/2010/main",
  asvg: "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
  v: "urn:schemas-microsoft-com:vml",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
};

const REL = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
};

/** A binary image part inside the package (word/media/imageN.ext). */
export interface DocxImagePart {
  /** Package path, e.g. "word/media/image1.jpeg". */
  path: string;
  bytes: Uint8Array;
  mime: string;
  width: number;
  height: number;
}

/** One place where an image part is displayed. The same part can be used several times. */
export interface DocxImageUsage {
  part: DocxImagePart;
  /** Crop taken from a:srcRect, if any. */
  crop?: CropFractions;
  /** Display size in EMUs from wp:extent (when known). */
  extentEmu?: { cx: number; cy: number };
  /** Name from wp:docPr / pic:cNvPr, for reporting. */
  name?: string;
  /** Story part the usage lives in (word/document.xml, word/header1.xml, ...). */
  storyPart: string;
  /** Document order within all scanned stories. */
  order: number;
  /**
   * Reasons why replacing this usage with the original file would change the
   * rendering (Word bakes these into the exported bitmap). Empty when safe.
   */
  blockers: string[];
}

export interface DocxImages {
  parts: DocxImagePart[];
  usages: DocxImageUsage[];
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/**
 * Extract all raster image parts from a .docx together with every place they are used.
 */
export async function extractDocxImages(docxBytes: Uint8Array, platform: Platform): Promise<DocxImages> {
  const zip = unzipSync(docxBytes);
  const files = new Map<string, Uint8Array>();
  for (const [name, data] of Object.entries(zip)) files.set(normalizePath(name), data);

  const contentTypes = parseContentTypes(files, platform);
  const rootRels = parseRels(files, "_rels/.rels", platform);
  const mainRel = rootRels.find((r) => r.type === REL.officeDocument);
  const mainPart = mainRel ? resolveTarget("", mainRel.target) : "word/document.xml";
  if (!files.has(mainPart)) throw new Error(`Main document part not found in package (${mainPart})`);

  const mainRels = parseRels(files, relsPathFor(mainPart), platform);
  const storyParts = [mainPart];
  for (const rel of mainRels) {
    if ([REL.header, REL.footer, REL.footnotes, REL.endnotes].includes(rel.type) && !rel.external) {
      const target = resolveTarget(mainPart, rel.target);
      if (files.has(target)) storyParts.push(target);
    }
  }

  const parts = new Map<string, DocxImagePart>();
  const usages: DocxImageUsage[] = [];
  let order = 0;

  const getPart = async (path: string): Promise<DocxImagePart | null> => {
    const cached = parts.get(path);
    if (cached) return cached;
    const bytes = files.get(path);
    if (!bytes) return null;
    const mime = sniffMime(bytes, contentTypes.get(path) ?? mimeFromExtension(path));
    if (!isRasterMime(mime)) return null;
    let size = readImageSize(bytes);
    if (!size) {
      try {
        const raster = await platform.decodeImage(bytes, mime);
        size = { width: raster.width, height: raster.height };
      } catch {
        return null;
      }
    }
    const part: DocxImagePart = { path, bytes, mime, ...size };
    parts.set(path, part);
    return part;
  };

  for (const storyPart of storyParts) {
    const xml = files.get(storyPart);
    if (!xml) continue;
    const rels = parseRels(files, relsPathFor(storyPart), platform);
    const relById = new Map(rels.map((r) => [r.id, r]));
    const doc = platform.parseXml(strFromU8(xml));

    // DrawingML pictures: <a:blip r:embed="rIdN"/> inside <pic:blipFill> (also shape fills).
    const blips = Array.from(doc.getElementsByTagNameNS(NS.a, "blip"));
    for (const blip of blips) {
      const embedId = blip.getAttributeNS(NS.r, "embed") || blip.getAttribute("r:embed");
      if (!embedId) continue; // r:link only (linked picture) -> nothing embedded to restore
      const rel = relById.get(embedId);
      if (!rel || rel.external) continue;
      const part = await getPart(resolveTarget(storyPart, rel.target));
      if (!part) continue;

      const blockers: string[] = [];
      let crop: CropFractions | undefined;
      const blipFill = blip.parentNode as Element | null;
      if (blipFill) {
        const srcRect = childNS(blipFill, NS.a, "srcRect");
        if (srcRect) {
          crop = {
            left: pct(srcRect.getAttribute("l")),
            top: pct(srcRect.getAttribute("t")),
            right: pct(srcRect.getAttribute("r")),
            bottom: pct(srcRect.getAttribute("b")),
          };
          if (Object.values(crop).some((v) => v < 0)) blockers.push("negative crop (picture padded beyond its frame)");
          if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) blockers.push("crop removes whole image");
        }
        if (childNS(blipFill, NS.a, "tile")) blockers.push("tiled fill");
      }
      // Any child of a:blip other than a:extLst is a colour/transparency effect Word bakes into the export.
      for (const child of Array.from(blip.childNodes)) {
        if (child.nodeType !== 1) continue;
        const el = child as Element;
        if (el.localName === "extLst") {
          if (el.getElementsByTagNameNS(NS.a14, "imgProps").length) blockers.push("artistic effect");
          if (el.getElementsByTagNameNS(NS.asvg, "svgBlip").length) blockers.push("SVG source (rendered as vector)");
        } else {
          blockers.push(`picture effect (${el.localName})`);
        }
      }
      const pic = closestNS(blip, NS.pic, "pic");
      if (pic) {
        const spPr = childNS(pic, NS.pic, "spPr");
        if (spPr) {
          const effectLst = childNS(spPr, NS.a, "effectLst");
          if (effectLst && childNS(effectLst, NS.a, "softEdge")) blockers.push("soft edges");
        }
      }

      const inlineOrAnchor = closestAny(blip, [NS.wp], ["inline", "anchor"]);
      let extentEmu: DocxImageUsage["extentEmu"];
      let name: string | undefined;
      if (inlineOrAnchor) {
        const extent = childNS(inlineOrAnchor, NS.wp, "extent");
        if (extent) extentEmu = { cx: Number(extent.getAttribute("cx")), cy: Number(extent.getAttribute("cy")) };
        name = childNS(inlineOrAnchor, NS.wp, "docPr")?.getAttribute("name") ?? undefined;
      }
      if (!name && pic) {
        name = childNS(pic, NS.pic, "nvPicPr")?.getElementsByTagNameNS(NS.pic, "cNvPr")[0]?.getAttribute("name") ?? undefined;
      }

      usages.push({ part, crop, extentEmu, name: name || undefined, storyPart, order: order++, blockers });
    }

    // Legacy VML pictures: <v:imagedata r:id="rIdN"/>
    const imagedatas = Array.from(doc.getElementsByTagNameNS(NS.v, "imagedata"));
    for (const im of imagedatas) {
      const relId = im.getAttributeNS(NS.r, "id") || im.getAttribute("r:id");
      if (!relId) continue;
      const rel = relById.get(relId);
      if (!rel || rel.external) continue;
      const part = await getPart(resolveTarget(storyPart, rel.target));
      if (!part) continue;
      const blockers: string[] = [];
      for (const attr of ["cropleft", "croptop", "cropright", "cropbottom", "gain", "blacklevel", "gamma", "grayscale", "bilevel", "chromakey"]) {
        if (im.hasAttribute(attr)) blockers.push(`VML ${attr}`);
      }
      usages.push({ part, storyPart, order: order++, blockers, name: im.getAttribute("o:title") ?? undefined });
    }
  }

  return { parts: Array.from(parts.values()), usages };
}

// ---------------------------------------------------------------------------

function isRasterMime(mime: string): boolean {
  return ["image/jpeg", "image/png", "image/gif", "image/bmp", "image/tiff", "image/webp"].includes(mime);
}

/** a:srcRect values are in 1/1000ths of a percent. */
function pct(v: string | null): number {
  return v ? Number(v) / 100000 : 0;
}

function normalizePath(p: string): string {
  return p.replace(/^\/+/, "");
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash + 1) : "";
  const name = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir}_rels/${name}.rels`;
}

/** Resolve a relationship target relative to the source part. */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return normalizePath(target);
  const slash = sourcePart.lastIndexOf("/");
  const base = slash >= 0 ? sourcePart.slice(0, slash).split("/") : [];
  for (const seg of target.split("/")) {
    if (seg === "..") base.pop();
    else if (seg !== "." && seg !== "") base.push(seg);
  }
  return base.join("/");
}

function parseRels(files: Map<string, Uint8Array>, relsPath: string, platform: Platform): Relationship[] {
  const xml = files.get(relsPath);
  if (!xml) return [];
  const doc = platform.parseXml(strFromU8(xml));
  return Array.from(doc.getElementsByTagNameNS(NS.rels, "Relationship")).map((el) => ({
    id: el.getAttribute("Id") ?? "",
    type: el.getAttribute("Type") ?? "",
    target: el.getAttribute("Target") ?? "",
    external: el.getAttribute("TargetMode") === "External",
  }));
}

function parseContentTypes(files: Map<string, Uint8Array>, platform: Platform): Map<string, string> {
  const result = new Map<string, string>();
  const xml = files.get("[Content_Types].xml");
  if (!xml) return result;
  const doc = platform.parseXml(strFromU8(xml));
  const defaults = new Map<string, string>();
  for (const el of Array.from(doc.getElementsByTagNameNS(NS.ct, "Default"))) {
    defaults.set((el.getAttribute("Extension") ?? "").toLowerCase(), el.getAttribute("ContentType") ?? "");
  }
  for (const el of Array.from(doc.getElementsByTagNameNS(NS.ct, "Override"))) {
    result.set(normalizePath(el.getAttribute("PartName") ?? ""), el.getAttribute("ContentType") ?? "");
  }
  for (const path of files.keys()) {
    if (result.has(path)) continue;
    const ext = path.toLowerCase().split(".").pop() ?? "";
    const ct = defaults.get(ext);
    if (ct) result.set(path, ct);
  }
  return result;
}

function childNS(parent: Element, ns: string, localName: string): Element | null {
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 1 && (child as Element).localName === localName && (child as Element).namespaceURI === ns) {
      return child as Element;
    }
  }
  return null;
}

function closestNS(el: Element, ns: string, localName: string): Element | null {
  return closestAny(el, [ns], [localName]);
}

function closestAny(el: Element, namespaces: string[], localNames: string[]): Element | null {
  let cur: Node | null = el.parentNode;
  while (cur && cur.nodeType === 1) {
    const e = cur as Element;
    if (namespaces.includes(e.namespaceURI ?? "") && localNames.includes(e.localName)) return e;
    cur = cur.parentNode;
  }
  return null;
}
