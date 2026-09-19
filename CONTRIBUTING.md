# Contributing

Bug reports and pull requests are welcome. For export problems, include your Word version, operating system, reproduction steps, and relevant Picture details and Log output. Remove private document names and content before posting; use a minimal example you can share publicly.

Use the latest patch of Node.js 22 (22.15+), 24, or 26. CI checks all three release lines; Node.js 20 is end-of-life and is no longer supported. Before opening a pull request, run `npm run typecheck`, `npm test`, `npm run build`, and `npm audit`. Changes to the export pipeline should include a regression test and, where possible, a check against a real Word export.

## How the export works

1. Read Word’s PDF export and compressed document package through `getFileAsync`.
2. Extract embedded images and their crop/effect metadata from the document XML.
3. Walk PDF image XObjects, including those in form XObjects, and compare aspect ratios and grayscale thumbnails with candidate originals. The default minimum visual similarity is 0.9; undecodable PDF images use a unique aspect-ratio fallback.
4. Embed matching originals, replace image resource references, and remove unreferenced replaced objects. Page content streams remain unchanged. Uncropped JPEG bytes are preserved; PNGs are embedded losslessly, and cropped images are re-encoded.
5. Save through the file picker or offer a browser download.

## Development

```bash
npm ci
```

```bash
npm start
```

`npm start` (run on Windows or Mac) starts the HTTPS dev server on `https://localhost:3000`, creates and trusts a development certificate, registers `manifest.xml` for sideloading and opens `file-sample.docx` in Word. `npm run stop` unregisters it again.

The development and production manifests share an add-in ID. Unregister the production add-in before starting local development to avoid conflicts.

### Developing from WSL with Word on Windows

`office-addin-debugging` cannot write to the Windows registry from inside WSL. Either copy the project to a Windows path and run `npm start` there (simplest), or run `npm run dev-server` in WSL, trust `~/.office-addin-dev-certs/ca.crt` on the Windows side (`certutil -addstore -user Root <path>`), and register the manifest with a `.reg` file:

```text
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer]
"03305ea7-f246-41ac-8451-c52f63a8f2ff"="C:\\path\\to\\manifest.xml"
```

The value name is the `<Id>` from `manifest.xml`.

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
scripts/         fix-pdf (CLI), register-prod, make-test-pdf, simulate-word-export, extract-images, inspect-pdf
public/          static site: landing page, support, privacy, terms (deployed next to the add-in)
test/            node:test suites and the browser harness
manifest.xml     add-in only (XML) manifest, Word desktop, ribbon button on Home
.github/         GitHub Pages deployment workflow
```


## Dependency updates

Dependabot checks npm packages and GitHub Actions weekly. Run `npm outdated` to check available versions and `npm audit` to check the complete dependency tree, including development tools, for known vulnerabilities. CI fails if the audit reports a vulnerability or cannot complete.

TypeScript stays on the latest 6.x release because [ts-loader does not yet support TypeScript 7](https://github.com/TypeStrong/ts-loader/issues/1702). Revisit that constraint when the loader gains support; it does not prevent using Node.js 26.

The `adm-zip` and `qs` overrides enforce patched transitive versions in the Office tooling and development server. `adm-zip` 0.6.1 fixes ZIP extraction and memory-allocation issues ([release notes](https://github.com/cthackers/adm-zip/releases/tag/v0.6.1)); `qs` 6.16.0 addresses [GHSA-4mjr-xmp4-gh2g](https://github.com/advisories/GHSA-4mjr-xmp4-gh2g) and [GHSA-x5fp-wj9c-mxmx](https://github.com/advisories/GHSA-x5fp-wj9c-mxmx). Remove an override only when all parent dependencies resolve to patched versions without it.
