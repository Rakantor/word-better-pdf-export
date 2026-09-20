# Better PDF Export for Word

Export PDFs from Microsoft Word with sharper pictures. This free, MIT-licensed add-in uses Word’s PDF exporter, then restores matching pictures from the image files stored in your document.

Word can downsample or recompress pictures during PDF export. Better PDF Export replaces those pictures at their stored resolution while retaining Word’s page layout and drawing instructions. It cannot recover image quality already lost in the document itself.

Document processing runs locally inside Word’s task pane. The add-in does not upload your documents or pictures. Its program files load from GitHub Pages and Microsoft’s Office CDN; see the [privacy policy](https://rakantor.github.io/word-better-pdf-export/privacy.html).

## Install

You need **Microsoft 365 Word for Windows or Mac** and permission to sideload Office add-ins. Word on the web is not supported. The current installation flow targets desktop Word, not iPad.

Until the add-in is listed in Microsoft Marketplace, every installation below is a *sideload*, and Word does not load sideloaded add-ins on its own. After each start of Word, open **Home → Add-ins → Developer Add-ins** and choose **Better PDF Export**. The **Export PDF** button then appears in the *Better PDF Export* group on the Home tab for the rest of that Word session.

### Windows

Follow the [installation instructions](https://rakantor.github.io/word-better-pdf-export/#install) to download and register the manifest. You do not need Node.js or a copy of this repository.

Keep the downloaded manifest in its registered location: Word reads it again when it starts. If your organization blocks sideloading, contact your Microsoft 365 administrator.

### Mac

Download the [production manifest](https://rakantor.github.io/word-better-pdf-export/manifest.xml) and follow [Microsoft’s sideloading instructions for Word on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac). Use the production manifest, which loads the hosted add-in without a local development server.

### Install from source

With the latest patch of Node.js 22 (22.15+), 24, or 26 installed, clone this repository and run these commands from its directory:

```bash
npm ci
npm run build
npm run register:prod
```

Restart Word and open the add-in from **Developer Add-ins** as described above. This registers the hosted add-in using a manifest copied to a stable per-user location (`%LOCALAPPDATA%\BetterPdfExport\manifest.xml` on Windows), so the checkout can be moved or deleted afterwards. Run `npm run unregister:prod` from a checkout with dependencies installed to remove it.

## Export a PDF

1. Open your document in Word.
2. On the **Home** tab, choose **Export PDF** in the **Better PDF Export** group. If the button is missing, choose **Add-ins → Developer Add-ins → Better PDF Export** first (see [Install](#install)).
3. Click **Export PDF with original pictures**.
4. Choose a save location when prompted, or save the download offered when export finishes.

The task pane reports which pictures were replaced, skipped, or left unmatched. Review the exported PDF before sharing it.

## What changes?

| Content | Result |
| --- | --- |
| Matching JPEG pictures without cropping | Original JPEG bytes embedded |
| Matching PNG pictures | Embedded losslessly |
| Cropped pictures | Visible region restored from the original at full resolution; JPEG crops are re-encoded at quality 95 |
| Pictures with detected pixel-changing effects | Skipped to preserve Word’s rendering |
| Pictures already stored without loss at full resolution | Skipped when detected |
| Pictures with no suitable match or unsupported formats | Left as Word exported them |
| Text, shapes, charts, and vector graphics | Page drawing instructions retained |

## Limitations

- Higher-resolution pictures can make the PDF substantially larger and require more memory during export.
- Only images stored in the document can be restored. Previously compressed originals and linked images that are not embedded cannot be recovered.
- Matching is visual. Near-identical pictures may be confused, so check the picture report and output. If a PDF image cannot be decoded, the pipeline falls back to a unique aspect-ratio match, which is less reliable.
- Cropped JPEGs are re-encoded; only uncropped JPEGs retain the exact original file bytes. Other supported bitmap formats may be converted to PNG.
- Detected effects such as recolouring, brightness, transparency, and soft edges are left alone. Unsupported images or effects may limit restoration.
- The PDF is rewritten with pdf-lib. Re-validate the result if you require PDF/A conformance or other formal PDF properties.

## Help and feedback

See [Support](https://rakantor.github.io/word-better-pdf-export/support.html) or [open an issue](https://github.com/Rakantor/word-better-pdf-export/issues). Include your Word version, operating system, and relevant **Picture details** and **Log** output. Remove private names and content before posting; attach only documents you can share publicly.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for local setup, WSL instructions, tests, the command-line exporter, and the source layout.

## License

[MIT](LICENSE). This is an independent project, not affiliated with or endorsed by Microsoft.
