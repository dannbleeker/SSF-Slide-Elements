import { defineConfig } from "vitest/config";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { MEASURED } from "./scripts/coverage-scope.mjs";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    coverage: {
      provider: "v8",
      // The engine and the pure decisions are the product, and the pane's
      // renderer earns the same floor because jsdom can check it. What will
      // NOT be counted here is the Office.js-touching code, because a
      // well-tested engine hiding behind an untestable pane, or the reverse,
      // is exactly what a pooled number lets happen.
      // From `scripts/coverage-scope.mjs`, so a test can hold the list against
      // the directories that actually exist.
      include: MEASURED.map((dir) => `src/${dir}/**`),
      // The pane's entry point is the one file there that touches Office.js and
      // cannot run in the suite. Counting it would drag the number down for a
      // reason nobody can act on, which is how a threshold gets deleted.
      exclude: ["src/pane/main.ts", "src/pane/*.html", "src/pane/*.css", "src/pane/env.d.ts"],
      reporter: ["text", "lcov"],
      // Floors sit a little under what the suite achieves, so an ordinary
      // change does not go red, and they are raised deliberately rather than
      // tracking the current number. A threshold that follows coverage upward
      // on its own only ever ratchets, and the first hard week gets it deleted.
      //
      // Measured with the package layer in, 2026-09-08: 97.8 statements, 87.4
      // branches, 100 functions, 99.7 lines. The branches that are not reached
      // are `Pkg`'s and `xml.ts`'s defensive nulls — an attribute list that is
      // not there, a document with no root — which `@xmldom/xmldom` never
      // produces. Raise them the same way when the harvest and the splice land:
      // measure, then leave two or three points of headroom, and say what you
      // measured. (The scaffold alone measured 98.1 / 95.1 / 100 / 97.7.)
      thresholds: { statements: 95, branches: 85, functions: 97, lines: 97 },
    },
  },
});
