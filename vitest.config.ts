import { defineConfig } from "vitest/config";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { MEASURED } from "./scripts/coverage-scope.mjs";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    /**
     * Vitest's default is five seconds, and this suite does real work: several
     * files harvest one or both committed library decks — 1.5 MB of .pptx, 117
     * elements, a megabyte of markup — and the splice sweep runs every one of
     * them through a package build and an integrity check.
     *
     * Those cases take six or seven seconds on a CI runner and about eleven
     * here, so the default killed them THERE and nowhere else: the same commit
     * was green locally and red on Linux twice, once for three cases and once
     * for a fourth nobody had noticed. Per-case timeouts fixed the three that
     * failed and left the fourth, which is what a per-case fix does.
     *
     * Raised for the whole suite instead, because a slow test here means a slow
     * MACHINE, never a wedged one: nothing in the suite waits on a network, a
     * host or a clock. A hung test now takes a minute to say so, which costs a
     * minute of CI and is worth it.
     */
    testTimeout: 60_000,
    hookTimeout: 180_000,
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
      // produces. (The scaffold alone measured 98.1 / 95.1 / 100 / 97.7.)
      //
      // Measured again with the splice, the host handshake and the picker in,
      // 2026-09-10: **97.7 statements, 91.9 branches, 100 functions, 98.9
      // lines**. Branches rose by four and a half points because the new code
      // is mostly decisions, and every decision is a pure function the suite
      // can put both ways. So the branch floor rises 85 → 88, which is three
      // points of headroom; the rest keep theirs, which they already had.
      //
      // Measured again on 2026-09-12, after a pass over the five files with the
      // thinnest branches: **98.4 statements, 93.5 branches, 100 functions,
      // 99.6 lines**. Branches rose a point on tests rather than on code —
      // `keepOnly` had no test of its own at all and was reached only through
      // its three callers, so its refusal had never run; `coloursOf` had never
      // been asked for a theme with no master, a master that is not in the
      // package, or a colour map naming a slot the theme does not carry;
      // `relatedOfType` had never met an external, dangling or targetless
      // relationship. So statements rise 95 → 96 and branches 88 → 90.
      //
      // What is LEFT uncovered in those files is one kind of thing, and it is
      // deliberately left: nullish guards TypeScript requires over DOM calls
      // that never return null in `@xmldom/xmldom` — a matched regex group,
      // `attributes.item(i)` inside its own length, an element's
      // `ownerDocument`. Reaching them means faking a DOM that cannot exist,
      // and a test that does that asserts about the fake.
      //
      // Raise them the same way next time: measure, leave two or three points,
      // and say what you measured. A threshold that follows coverage upward on
      // its own only ever ratchets, and the first hard week gets it deleted.
      thresholds: { statements: 96, branches: 90, functions: 97, lines: 97 },
    },
  },
});
