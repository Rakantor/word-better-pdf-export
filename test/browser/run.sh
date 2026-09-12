#!/usr/bin/env sh
# Bundles the browser harness and serves the repo root on http://localhost:8123.
# Open http://localhost:8123/test/browser/ in a Chromium-based browser.
set -e
cd "$(dirname "$0")/../.."
npm run make-test-pdf -- file-sample.docx file-sample.simulated-word.pdf >/dev/null
npx esbuild test/browser/harness.ts --bundle --format=iife --outfile=test/browser/harness.js --log-level=warning
exec python3 -m http.server 8123 --bind 127.0.0.1
