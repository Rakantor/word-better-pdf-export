# Better PDF Export for Word

A small Word add-in (Microsoft 365, Windows desktop) that exports the open document to PDF **without Word's picture compression**.

Word's built-in PDF export downsamples pictures to roughly 220 ppi and recompresses them as JPEG. This add-in still uses Word's converter for everything (layout, fonts, vector graphics, bookmarks), but afterwards swaps every compressed picture in the PDF for the untouched original image file that is stored inside the `.docx`.

```
┌────────────┐  getFileAsync(Pdf)         ┌───────────────┐
│            │ ─────────────────────────► │ Word's PDF    │──┐
│  Word      │                            └───────────────┘  │   match pictures      ┌──────────────┐
│  document  │  getFileAsync(Compressed)  ┌───────────────┐  ├──────────────────────►│ PDF with     │
│            │ ─────────────────────────► │ .docx package │──┘   swap XObjects        │ original     │
└────────────┘                            │ word/media/*  │       (pdf-lib)           │ pictures     │
                                          └───────────────┘                           └──────────────┘
```

Everything runs inside the task pane; nothing leaves the machine.

## How it works

1. `Office.context.document.getFileAsync(Office.FileType.Pdf)` — Word's own PDF export, read in 4 MB slices.
2. `Office.context.document.getFileAsync(Office.FileType.Compressed)` — the current document as a `.docx` (a zip). `word/media/*` holds the original image files exactly as they were inserted; the document XML says where each one is used, whether it is cropped (`a:srcRect`) and whether picture effects are applied.
3. The PDF is parsed with [pdf-lib](https://pdf-lib.js.org/). Every image XObject reachable from the pages (including those nested in form XObjects) is decoded and compared with the candidate originals: first by aspect ratio, then by normalised cross-correlation of small grayscale thumbnails. A match needs a similarity of ≥ 0.9 (typical: 0.999).
4. For each match the original file is embedded (JPEG bytes verbatim, PNG losslessly) and the XObject entry is pointed at it. The page content stream is untouched — the drawing matrix already defines the picture's size on the page, so a higher-resolution bitmap just renders sharper. Cropped pictures are cropped from the original at full resolution and re-encoded (JPEG q95 or lossless PNG). Replaced objects that nothing references any more are removed.
5. The result is saved through the browser's "Save as" dialog (File System Access API) or, if unavailable, as a download.

### What is deliberately left alone

| Situation | Behaviour |
| --- | --- |
| Picture effects Word bakes into the bitmap (recolour, brightness/contrast, transparency, artistic effects, soft edges) | skipped, reported as such — replacing would drop the effect |
| SVG / EMF / WMF pictures | not touched — Word already exports them as vectors |
| Charts, shapes, text | not touched — vector output of Word's converter |
| Pictures already stored losslessly at full resolution | skipped |
| Images in the PDF with no visually matching original (e.g. inserted via other means) | left as is |

Effects that don't alter the pixels (rotation, borders, shadows, glow, reflections) are fine: Word draws those around the bitmap.

## Requirements

- Word for Windows or Mac, Microsoft 365 (the PDF file type of `getFileAsync` is supported on Windows, Mac and iPad, **not** in Word on the web).
- Node.js 20–24 for building/serving the add-in.

## Install (production build from GitHub Pages)

The add-in is deployed to <https://rakantor.github.io/word-better-pdf-export/>. Until it is listed in Microsoft Marketplace, install it by registering the production manifest with Word once. On Windows, from a clone of this repo:

```bash
npm install
```

```bash
npm run build
```

```bash
npm run register:prod
```

Then restart Word; **Export PDF** is on the Home tab. `npm run unregister:prod` removes it again. Without Node, follow the manual registry instructions on the [site](https://rakantor.github.io/word-better-pdf-export/#install).

## Development

```bash
npm install
```

```bash
npm start
```

`npm start` (run on Windows or Mac) starts the HTTPS dev server on `https://localhost:3000`, creates and trusts a development certificate, registers `manifest.xml` for sideloading and opens `file-sample.docx` in Word. `npm run stop` unregisters it again.

Both the dev manifest and the production manifest can be registered at the same time; they have the same ID, so Word shows whichever was registered last.

### Developing from WSL with Word on Windows

`office-addin-debugging` cannot write to the Windows registry from inside WSL. Either copy the project to a Windows path and run `npm start` there (simplest), or run `npm run dev-server` in WSL, trust `~/.office-addin-dev-certs/ca.crt` on the Windows side (`certutil -addstore -user Root <path>`), and register the manifest with a `.reg` file:

```text
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer]
"03305ea7-f246-41ac-8451-c52f63a8f2ff"="C:\\path\\to\\manifest.xml"
```

The value name is the `<Id>` from `manifest.xml`.

## Deployment

Every push to `main` runs [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml): typecheck, tests, production build, manifest validation, then deploy `dist/` to GitHub Pages. The production build rewrites every `https://localhost:3000/` in the manifest to `urlProd` from [webpack.config.js](webpack.config.js) and copies the static pages in `public/` (landing page, support, privacy policy, terms of use — the URLs Microsoft Marketplace requires).

One-time repository setup: **Settings → Pages → Build and deployment → Source: GitHub Actions**.

## Testing the core against a real Word export (no add-in needed)

The image replacement is plain TypeScript with no Office dependency, so it can be run from the command line on a PDF that Word produced via *File → Save As → PDF*:

```bash
npm run fix-pdf -- file-sample.docx "path/to/word-export.pdf"
```

This writes `word-export.original-images.pdf` next to the input and prints one line per image (replaced / skipped / unmatched, with similarity scores). To look at what is actually embedded in either PDF:

```bash
npx tsx scripts/extract-images.ts "path/to/word-export.pdf" extracted
```

## Scripts

```bash
npm test          # node:test suite (crop, effects, decoy, passthrough, Flate/predictor decoding, form XObjects)
npm run typecheck
npm run build     # production bundle in dist/ (replaces https://localhost:3000/ with urlProd from webpack.config.js)
npm run validate  # office-addin-manifest validate
npm run fix-pdf -- doc.docx export.pdf   # CLI version of the pipeline, for testing against Word's own export
```

`npm run make-test-pdf` writes `file-sample.simulated-word.pdf`, a PDF that imitates Word's export (pictures downsampled to 220 ppi and recompressed, plus a decoy picture with the same aspect ratio) — handy for iterating without Word. `test/browser/run.sh` bundles a harness that runs the same pipeline with the canvas-based codecs in a real Chromium at `http://localhost:8123/test/browser/`.

### Layout

```
src/core/        environment-agnostic pipeline
  docx.ts        unzip .docx, resolve relationships, list picture usages (crop, effects, extent, name)
  pdf-images.ts  walk page resources → image XObjects (recursing into form XObjects)
  pdf-decode.ts  DCT / Flate(+PNG predictors) / Gray / RGB / CMYK / ICC / Indexed / SMask → RGBA
  raster.ts      thumbnails, normalised cross-correlation, cropping, resampling
  replace.ts     matching, embedding, slot rewriting, orphan cleanup, report
src/browser/     Platform impl. using createImageBitmap / OffscreenCanvas / DOMParser (the add-in)
src/node/        Platform impl. using jpeg-js / pngjs / @xmldom/xmldom (tests and CLI)
src/taskpane/    task pane UI, getFileAsync wrapper, save helpers
scripts/         fix-pdf (CLI), make-test-pdf, simulate-word-export, extract-images, inspect-pdf
public/          static site: landing page, support, privacy, terms (deployed next to the add-in)
test/            node:test suites and the browser harness
manifest.xml     add-in only (XML) manifest, Word desktop, ribbon button on Home
.github/         GitHub Pages deployment workflow
```

## Notes and limitations

- Output size grows by the size of the original pictures; that is the point, but a 20 MB photo becomes a 20 MB PDF.
- Cropped pictures are re-encoded from the decoded original (JPEG q95 for JPEG sources, lossless PNG otherwise), because Word exports only the visible region. Uncropped pictures are byte-for-byte originals.
- The PDF is rewritten by pdf-lib. If you rely on PDF/A conformance of Word's output, re-validate the result.
- Matching is visual. Two different pictures with the same aspect ratio are told apart reliably; two near-identical pictures with the same aspect ratio may be interchanged (harmless for identical files).
- An image the decoder cannot read (an unusual filter) is only replaced if exactly one original has its aspect ratio.
