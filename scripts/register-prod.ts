/**
 * Install or remove the production add-in in the local Word, without the dev server.
 *
 *   npm run register:prod      (after `npm run build`)
 *   npm run unregister:prod
 *
 * Word re-reads registered manifests from disk at every start, so the registered
 * path must outlive this checkout: the manifest is copied to a per-user location
 * (%LOCALAPPDATA%\BetterPdfExport on Windows, ~/.better-pdf-export elsewhere) and
 * that copy is registered. Deleting or moving the project folder afterwards is fine.
 */
import { copyFileSync, existsSync, mkdirSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { registerAddIn, unregisterAddIn } from "office-addin-dev-settings";

const builtManifest = path.resolve("dist/manifest.xml");
const stableDir =
  process.platform === "win32"
    ? path.join(process.env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local"), "BetterPdfExport")
    : path.join(os.homedir(), ".better-pdf-export");
const stableManifest = path.join(stableDir, "manifest.xml");

async function main(): Promise<void> {
  const action = process.argv[2];
  if (action === "register") {
    if (!existsSync(builtManifest)) {
      throw new Error(`${builtManifest} not found. Run "npm run build" first.`);
    }
    mkdirSync(stableDir, { recursive: true });
    copyFileSync(builtManifest, stableManifest);
    await registerAddIn(stableManifest);
    console.log(`Registered ${stableManifest}\nRestart Word: "Export PDF" appears on the Home tab.`);
  } else if (action === "unregister") {
    // Unregistering needs the manifest (for its ID); fall back to the build output if the copy is gone.
    const manifest = existsSync(stableManifest) ? stableManifest : builtManifest;
    if (!existsSync(manifest)) throw new Error("No manifest found to unregister (neither the installed copy nor dist/manifest.xml).");
    await unregisterAddIn(manifest);
    rmSync(stableDir, { recursive: true, force: true });
    console.log("Unregistered. Restart Word to remove the ribbon button.");
  } else {
    throw new Error("usage: register-prod <register|unregister>");
  }
}

main().catch((err) => {
  console.error((err as Error).message);
  process.exit(1);
});
