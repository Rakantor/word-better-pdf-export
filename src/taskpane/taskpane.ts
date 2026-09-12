import { browserPlatform } from "../browser/platform-browser";
import { replaceImagesInPdf, type ImageResult, type ReplaceReport } from "../core";
import { getDocumentFile } from "./office-file";
import { downloadBytes, pickSaveTarget, suggestedPdfName, supportsSavePicker, type SaveTarget } from "./save";

type StepId = "pdf" | "docx" | "replace" | "save";
type StepState = "pending" | "active" | "done" | "error" | "skipped";

const el = {
  export: byId<HTMLButtonElement>("export"),
  download: byId<HTMLButtonElement>("download"),
  hint: byId<HTMLElement>("hint"),
  steps: byId<HTMLOListElement>("steps"),
  summary: byId<HTMLElement>("summary"),
  report: byId<HTMLDetailsElement>("report"),
  reportRows: byId<HTMLTableSectionElement>("report-rows"),
  logDetails: byId<HTMLDetailsElement>("log-details"),
  log: byId<HTMLPreElement>("log"),
  error: byId<HTMLElement>("error"),
};

let lastPdf: { bytes: Uint8Array; name: string } | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    showError("This add-in only works in Word.");
    return;
  }
  const canExportPdf = Office.context.requirements.isSetSupported("File", "1.1");
  const canGetDocx = Office.context.requirements.isSetSupported("CompressedFile", "1.1");
  if (!canExportPdf || !canGetDocx) {
    showError("This version of Word cannot hand documents to add-ins as PDF. Use Word for Windows or Mac (Microsoft 365).");
    return;
  }
  el.export.disabled = false;
  el.hint.textContent = supportsSavePicker()
    ? "You will be asked where to save the PDF first, then the export runs."
    : "The PDF is offered as a download when the export finishes.";
  el.export.addEventListener("click", () => void runExport());
  el.download.addEventListener("click", () => {
    if (lastPdf) downloadBytes(lastPdf.bytes, lastPdf.name);
  });
});

async function runExport(): Promise<void> {
  resetUi();
  el.export.disabled = true;
  const name = suggestedPdfName(Office.context.document.url);

  try {
    // Ask for the destination while we still have the click's user activation.
    let target: SaveTarget | null = null;
    if (supportsSavePicker()) {
      target = await pickSaveTarget(name);
      if (!target) {
        el.hint.textContent = "Export cancelled.";
        return;
      }
    }
    el.steps.hidden = false;

    setStep("pdf", "active");
    const pdfBytes = await getDocumentFile(Office.FileType.Pdf, (got, total) => setStep("pdf", "active", progress(got, total)));
    setStep("pdf", "done", formatBytes(pdfBytes.length));

    setStep("docx", "active");
    const docxBytes = await getDocumentFile(Office.FileType.Compressed, (got, total) => setStep("docx", "active", progress(got, total)));
    setStep("docx", "done", formatBytes(docxBytes.length));

    setStep("replace", "active");
    const { pdf, report } = await replaceImagesInPdf(pdfBytes, docxBytes, { platform: browserPlatform, log: appendLog });
    setStep("replace", "done", `${report.replaced} of ${report.pdfImages} picture(s) restored`);

    setStep("save", "active");
    lastPdf = { bytes: pdf, name: target?.name ?? name };
    if (target) {
      await target.write(pdf);
      setStep("save", "done", `${target.name} · ${formatBytes(pdf.length)}`);
    } else {
      setStep("save", "done", `${formatBytes(pdf.length)} ready`);
      el.download.hidden = false;
      downloadBytes(pdf, name);
    }
    showSummary(report);
  } catch (err) {
    const active = el.steps.querySelector<HTMLLIElement>('li[data-state="active"]');
    if (active) active.dataset.state = "error";
    showError((err as Error).message || String(err));
    appendLog(`ERROR ${(err as Error).stack ?? String(err)}`);
  } finally {
    el.export.disabled = false;
  }
}

// -- UI helpers --------------------------------------------------------------

function resetUi(): void {
  lastPdf = null;
  el.error.hidden = true;
  el.error.textContent = "";
  el.summary.hidden = true;
  el.report.hidden = true;
  el.reportRows.replaceChildren();
  el.download.hidden = true;
  el.log.textContent = "";
  el.logDetails.hidden = true;
  el.hint.textContent = "";
  for (const li of el.steps.querySelectorAll<HTMLLIElement>("li")) {
    li.dataset.state = "pending";
    li.querySelector(".detail")!.textContent = "";
  }
}

function setStep(id: StepId, state: StepState, detail = ""): void {
  const li = el.steps.querySelector<HTMLLIElement>(`li[data-step="${id}"]`);
  if (!li) return;
  li.dataset.state = state;
  li.querySelector(".detail")!.textContent = detail;
}

function showSummary(report: ReplaceReport): void {
  const ratio = report.inputBytes ? report.outputBytes / report.inputBytes : 1;
  const headline =
    report.replaced > 0
      ? `<strong class="ok">${report.replaced} picture${report.replaced === 1 ? "" : "s"} restored to original quality</strong>`
      : report.pdfImages === 0
        ? `<strong>No pictures in this PDF</strong>`
        : `<strong class="warn">No pictures were replaced</strong>`;
  const notes: string[] = [];
  if (report.skipped) notes.push(`${report.skipped} skipped`);
  if (report.unmatched) notes.push(`${report.unmatched} not matched to a picture in the document`);
  el.summary.innerHTML = `${headline}<span>${formatBytes(report.inputBytes)} → ${formatBytes(report.outputBytes)} (${ratio.toFixed(1)}×)${
    notes.length ? " · " + notes.join(", ") : ""
  }</span>`;
  el.summary.hidden = false;

  if (report.results.length) {
    for (const r of report.results) el.reportRows.appendChild(reportRow(r));
    el.report.hidden = false;
  }
  el.logDetails.hidden = false;
}

function reportRow(r: ImageResult): HTMLTableRowElement {
  const tr = document.createElement("tr");
  const pages = r.pages.map((p) => p + 1).join(", ") || "–";
  const inPdf = `${r.width}×${r.height} ${r.filters.includes("DCTDecode") ? "JPEG" : r.filters.join("+") || "raw"}`;
  let result: string;
  switch (r.status) {
    case "replaced":
      result = `<span class="replaced">Replaced</span> → ${escapeHtml(r.to)}<br><small>${escapeHtml(r.source)}</small>`;
      break;
    case "skipped":
      result = `<span class="skipped">Skipped</span> – ${escapeHtml(r.reason)}${r.source ? `<br><small>${escapeHtml(r.source)}</small>` : ""}`;
      break;
    default:
      result = `<span class="unmatched">Left as is</span> – ${escapeHtml(r.reason)}`;
  }
  tr.innerHTML = `<td>${pages}</td><td>${inPdf}</td><td>${result}</td>`;
  return tr;
}

function showError(message: string): void {
  el.error.textContent = message;
  el.error.hidden = false;
}

function appendLog(message: string): void {
  el.log.textContent += `${message}\n`;
  el.logDetails.hidden = false;
}

function progress(got: number, total: number): string {
  return total ? `${formatBytes(got)} / ${formatBytes(total)}` : formatBytes(got);
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}

function byId<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`missing element #${id}`);
  return node as T;
}
