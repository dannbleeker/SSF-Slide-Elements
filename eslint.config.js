// @ts-check
import js from "@eslint/js";
import globals from "globals";
import tseslint from "typescript-eslint";

/**
 * Type-aware linting, deliberately.
 *
 * The engine will be async from top to bottom: every part read, every copy,
 * every write returns a promise. A forgotten `await` there does not throw, it
 * splices the wrong thing quietly, which is the failure mode this project can
 * least afford. `no-floating-promises` and `no-misused-promises` need type
 * information to see it at all, so the cost of the slower lint is the point.
 */
export default tseslint.config(
  // `probe/` is generated for Script Lab, not for this project: it has no
  // module system and a `main()` at the top level, and it is typechecked
  // against the real Office.js types by test/probe.test.ts rather than by the
  // linter.
  { ignores: ["dist/", "dist-lib/", "coverage/", "public/", "probe/"] },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        // The config files sit outside tsconfig's include, and type-aware
        // linting refuses to parse a file no project owns.
        //
        // The COUNT is raised because the default is eight and `scripts/` goes
        // past it — at which point every file over the line fails with a
        // parsing error rather than a finding, so `npm run lint` reports the
        // linter's own limit as though it were a defect in the code.
        projectService: {
          allowDefaultProject: ["*.js", "*.ts", "scripts/*.mjs"],
          // Raise this when a script is added and the linter starts reporting
          // its own limit as a parsing error.
          maximumDefaultProjectFileMatchCount_THIS_WILL_SLOW_DOWN_LINTING: 60,
        },
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      // A caught error is often re-thrown or reported; the engine does that a
      // lot and does not need a type argument each time.
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      // An unused argument named with a leading underscore is a documented
      // signature, not an oversight.
      "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
    },
  },
  {
    // The build scripts run under Node, not in a pane.
    files: ["scripts/**", "*.config.{js,ts}"],
    languageOptions: { globals: globals.node },
  },
  {
    // A Node script whose `page.evaluate` callbacks are serialised and run
    // inside the browser, so `document` there is real.
    files: ["scripts/pane-shots.mjs", "scripts/build-previews.mjs"],
    languageOptions: { globals: { ...globals.node, ...globals.browser } },
  },
  {
    // Untyped `.mjs` with no project behind it: every `RegExp.exec` result and
    // every `JSON.parse` is `any`, so a function returning one is an unsafe
    // return by construction rather than by mistake. The rules these files hold
    // are gated by `test/manifest.test.ts`, which proves each one can still
    // fail — a stronger check than the type of an intermediate.
    files: [
      "scripts/bench-engine.mjs",
      "scripts/build-previews.mjs",
      "scripts/build-probe.mjs",
      "scripts/catalogue-page.mjs",
      "scripts/harvest.mjs",
      "scripts/read-answers.mjs",
      "scripts/manifest-rules.mjs",
      "scripts/manifest-source.mjs",
      "scripts/release-assets.mjs",
      "scripts/without-prose.mjs",
    ],
    rules: { "@typescript-eslint/no-unsafe-return": "off" },
  },
  {
    // The tests reach into shapes on purpose.
    files: ["test/**"],
    rules: { "@typescript-eslint/no-explicit-any": "off" },
  },
);
