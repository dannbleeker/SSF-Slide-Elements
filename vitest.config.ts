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
      // Measured on the scaffold commit, 2026-09-08: 98.1 statements, 95.1
      // branches, 100 functions, 97.7 lines — over four small files, so one
      // uncovered branch moves the number by a point. Raise them the same way
      // when the engine lands: measure, then leave two or three points of
      // headroom, and say what you measured.
      thresholds: { statements: 95, branches: 90, functions: 97, lines: 95 },
    },
  },
});
