# Repository Guidelines

## Project Structure & Module Organization

- `src/core/` implements DOCX extraction, PDF image decoding, visual matching, and replacement. Keep this pipeline independent of Office and browser APIs.
- `src/browser/` and `src/node/` provide platform-specific image and XML operations; `src/taskpane/` contains the Word UI, Office file access, and save helpers.
- `test/` contains unit tests and shared fixtures; `test/browser/` exercises browser codecs. `file-sample.docx` is the shared sample document.
- `scripts/` contains CLI and registration tools. `assets/` holds icons; `public/` holds website pages. `manifest.xml` defines the add-in; webpack generates `dist/`.

## Build, Test, and Development Commands

Use Node.js 22 (22.15+), 24, or 26 with current patches.

- `npm ci`: install the locked dependencies.
- `npm start`: start HTTPS development, register the manifest, and open the sample in desktop Word. `npm run stop` unregisters it.
- `npm run dev-server`: run only the development server.
- `npm run typecheck`: check strict TypeScript types.
- `npm test`: run the Node test suite through `tsx`.
- `npm run build`: produce the hosted add-in and production manifest in `dist/`.
- `npm run validate`: validate the source manifest; use `npx --no-install office-addin-manifest validate dist/manifest.xml` after building to validate production.
- `npm audit`: check runtime and development dependencies for known vulnerabilities.

See `CONTRIBUTING.md` for WSL setup and CLI export examples.

## Coding Style & Naming Conventions

Follow existing TypeScript style: two-space indentation, double quotes, semicolons, camelCase functions and variables, and PascalCase types. Use descriptive kebab-case filenames such as `pdf-images.ts`. Prefer explicit types at platform boundaries. No project formatter or lint command is configured; preserve surrounding formatting.

## Testing Guidelines

Use `node:test` with `node:assert/strict`; name suites `test/*.test.ts`. Add regression tests for pipeline changes, including crop, effect, matching, and PDF decoding behavior. No numeric coverage threshold is configured. Where possible, verify against a real Word export and run the browser harness for codec changes.

## Commit & Pull Request Guidelines

Group changes logically. Use single-line Conventional Commit messages, for example `fix: preserve cropped picture effects`; existing prefixes include `feat`, `fix`, `docs`, `build`, `ci`, and `chore`. Describe the problem, resulting behavior, and validation in PRs; link relevant issues and include screenshots for UI changes. Run typecheck, tests, build, and audit before submission.

## Security & Configuration

Keep documents local and redact private content from shared logs or fixtures. Never commit credentials, certificates, or generated output. Preserve security overrides in `package.json` until parent dependencies resolve patched versions without them.
