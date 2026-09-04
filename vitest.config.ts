import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    // The splice sweep opens a 1.5 MB package ninety-eight times.
    testTimeout: 120_000,
    coverage: {
      provider: "v8",
      // The engine and the pure decisions are the product. `src/office` is
      // deliberately absent: pooling a well-tested engine with untestable host
      // calls produces one number that hides both.
      include: ["src/core/**", "src/host/**", "src/pane/**"],
      exclude: ["src/pane/main.ts", "src/pane/*.html", "src/pane/*.css", "src/pane/env.d.ts"],
      reporter: ["text", "lcov"],
      // Floors sit under what the suite achieves, so an ordinary change does
      // not go red, and they are raised deliberately rather than tracking the
      // current number. A threshold that follows coverage upward on its own
      // only ever ratchets, and the first hard week gets it deleted.
      thresholds: { statements: 80, branches: 75, functions: 80, lines: 80 },
    },
  },
});
