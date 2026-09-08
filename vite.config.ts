import { execFileSync } from "node:child_process";
import { resolve } from "node:path";
import { defineConfig } from "vite";

/**
 * Which commit the pane was built from.
 *
 * PowerPoint caches the pane's HTML for ten minutes. Open it too soon after a
 * deploy and the round tests code the host never fetched — and there is no way
 * to tell from the result, because the pane looks identical and the run log
 * reads as a clean run of the wrong build. A sibling project records whole
 * rounds lost to it.
 *
 * So the stamp is on screen and in the run record, and the owner can check it
 * against the commit they meant to test before spending ten minutes of a real
 * PowerPoint on it.
 *
 * Falls back to "unknown" rather than failing the build: a checkout with no git
 * (a release tarball, a CI runner without history) still has to produce a pane.
 * "unknown" is honest and a wrong commit hash would not be.
 */
function buildStamp(): string {
  try {
    return execFileSync("git", ["rev-parse", "--short", "HEAD"], { encoding: "utf8" }).trim() || "unknown";
  } catch {
    return "unknown";
  }
}

/**
 * The task pane, built for GitHub Pages.
 *
 * `base: "./"` because the add-in is served from a path that the manifest
 * points at, and absolute asset URLs would break the moment the site moved.
 * `publicDir` carries the CNAME that binds the custom domain and the landing
 * page, so one build produces the whole site.
 */
/** A development certificate, when the environment names one. Vite reads the two paths itself. */
const devKey = process.env["SSF_DEV_KEY"];
const devCert = process.env["SSF_DEV_CERT"];

export default defineConfig({
  // Replaced at build time, so the pane can say what it is. Stringified because
  // `define` substitutes source text, not values.
  define: { __BUILD_STAMP__: JSON.stringify(buildStamp()) },
  // The pane is the root, so it builds to dist/taskpane.html rather than to
  // dist/src/pane/taskpane.html. The manifest points at that URL and a
  // manifest change is the one kind that costs the owner a re-sideload, so the
  // path is worth getting right before there is a manifest to change.
  root: resolve(import.meta.dirname, "src/pane"),
  base: "./",
  publicDir: resolve(import.meta.dirname, "public"),
  // The DEV MANIFEST names https://localhost:3002/taskpane.html, so the dev
  // server has to answer there — a manifest pointing at a port nothing serves
  // is a blank pane with a generic error. `strictPort` because falling back to
  // 3003 would be the same failure with an extra step. 3002 rather than the
  // 3000 both siblings hold, so two SSF panes can be open on one machine.
  //
  // Office requires HTTPS for a sideloaded add-in. `npm run dev` serves plain
  // HTTP, which is right for looking at the pane in a browser; sideloading it
  // needs a certificate — `npx office-addin-dev-certs install`, then the two
  // paths it prints in SSF_DEV_KEY and SSF_DEV_CERT. Through the environment
  // rather than CLI flags, because Vite 8 has no `--https` flag: a sibling's
  // CONTRIBUTING documented one for weeks and nobody could have run it.
  server: { port: 3002, strictPort: true, ...(devKey && devCert ? { https: { key: devKey, cert: devCert } } : {}) },
  build: {
    outDir: resolve(import.meta.dirname, "dist"),
    emptyOutDir: true,
    // Named explicitly because the root has no index.html: that name belongs to
    // the landing page in public/, and a pane called index.html would collide
    // with it in dist/.
    rollupOptions: { input: resolve(import.meta.dirname, "src/pane/taskpane.html") },
  },
});
