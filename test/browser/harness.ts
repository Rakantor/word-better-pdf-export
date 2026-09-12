/**
 * Browser harness: runs the core with the canvas-based platform in a real Chromium
 * (the same engine as Word's WebView2 task pane). Bundled by test/browser/run.sh.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { simulateWordExport } from "../../scripts/simulate-word-export";
import { browserPlatform } from "../../src/browser/platform-browser";
import { replaceImagesInPdf } from "../../src/core";

declare global {
  interface Window {
    harnessResult?: unknown;
  }
}

async function main(): Promise<void> {
  const out = document.getElementById("out")!;
  const log: string[] = [];
  try {
    const [pdf, docx] = await Promise.all([
      fetch("/file-sample.simulated-word.pdf").then((r) => r.arrayBuffer()),
      fetch("/file-sample.docx").then((r) => r.arrayBuffer()),
    ]);
    const started = performance.now();
    const result = await replaceImagesInPdf(new Uint8Array(pdf), new Uint8Array(docx), {
      platform: browserPlatform,
      log: (m) => log.push(m),
    });
    // Second scenario: a cropped picture, which exercises canvas JPEG encoding.
    const files = unzipSync(new Uint8Array(docx));
    files["word/document.xml"] = strToU8(
      strFromU8(files["word/document.xml"]).replace(
        '<a:blip r:embed="rId4"/>',
        '<a:blip r:embed="rId4"/><a:srcRect l="10000" t="20000" r="30000" b="5000"/>',
      ),
    );
    const croppedDocx = zipSync(files);
    const croppedPdf = await simulateWordExport(croppedDocx, { platform: browserPlatform });
    const cropped = await replaceImagesInPdf(croppedPdf, croppedDocx, { platform: browserPlatform, log: (m) => log.push("[crop] " + m) });
    const summary = {
      ms: Math.round(performance.now() - started),
      report: { ...result.report, results: result.report.results },
      outputHeader: String.fromCharCode(...result.pdf.subarray(0, 8)),
      cropped: cropped.report.results,
      croppedOutputBytes: cropped.pdf.length,
      log,
    };
    window.harnessResult = summary;
    out.textContent = JSON.stringify(summary, null, 2);
  } catch (err) {
    window.harnessResult = { error: String((err as Error).stack ?? err), log };
    out.textContent = JSON.stringify(window.harnessResult, null, 2);
  }
}

void main();
