import { execFileSync } from "node:child_process";
import { resolve } from "node:path";
import { defineConfig } from "vite";

/**
 * Which commit the pane was built from.
 *
 * PowerPoint caches the pane's HTML for ten minutes. Open it too soon after a
 * deploy and the round tests code the host never fetched — and there is no way
 * to tell from the result, because the pane looks identical and the run reads
 * as a clean run of the wrong build. A sibling project records whole rounds
 * lost to it, so the stamp is on screen and can be checked before spending ten
 * minutes of a real PowerPoint on it.
 *
 * "unknown" rather than a failure: a checkout with no git still has to produce
 * a pane, and a wrong commit hash would be worse than an honest blank.
 */
function buildStamp(): string {
  try {
    return execFileSync("git", ["rev-parse", "--short", "HEAD"], { encoding: "utf8" }).trim() || "unknown";
  } catch {
    return "unknown";
  }
}

export default defineConfig({
  // Stringified because `define` substitutes source text, not values.
  define: { __BUILD_STAMP__: JSON.stringify(buildStamp()) },
  // The pane is the root, so it builds to dist/taskpane.html rather than
  // dist/src/pane/taskpane.html. The manifest points at that URL, and a
  // manifest change is the one kind that costs a re-sideload.
  root: resolve(import.meta.dirname, "src/pane"),
  base: "./",
  publicDir: resolve(import.meta.dirname, "public"),
  // The DEV MANIFEST names https://localhost:3000/taskpane.html, so the dev
  // server has to answer there — a manifest pointing at a port nothing serves
  // is a blank pane with a generic error. `strictPort` because falling back to
  // 3001 is the same failure with an extra step.
  server: { port: 3000, strictPort: true },
  build: {
    outDir: resolve(import.meta.dirname, "dist"),
    emptyOutDir: true,
    rollupOptions: { input: resolve(import.meta.dirname, "src/pane/taskpane.html") },
  },
});
