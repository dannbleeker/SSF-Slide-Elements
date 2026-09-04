// @ts-check
import js from "@eslint/js";
import globals from "globals";
import tseslint from "typescript-eslint";

/**
 * Type-aware linting, deliberately.
 *
 * The engine is async from top to bottom: every part read, every copy, every
 * write returns a promise. A forgotten `await` there does not throw, it splices
 * the wrong thing quietly, which is the failure mode this project can least
 * afford. `no-floating-promises` and `no-misused-promises` need type
 * information to see it at all, so the slower lint is the point.
 */
export default tseslint.config(
  { ignores: ["dist/", "dist-lib/", "coverage/", "public/", "public/catalogue/"] },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        // The config files sit outside tsconfig's include, and type-aware
        // linting refuses to parse a file no project owns.
        projectService: {
          allowDefaultProject: ["*.js", "*.ts", "scripts/*.mjs"],
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
      "@typescript-eslint/no-unsafe-return": "off",
      // An unused argument named with a leading underscore is a documented
      // signature, not an oversight.
      "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
      // xmldom's Element is structurally the DOM's but not nominally, and the
      // non-null assertions in the tests are on nodes the fixture just built.
      "@typescript-eslint/no-non-null-assertion": "off",
    },
  },
  {
    files: ["scripts/**", "*.config.{js,ts}"],
    languageOptions: { globals: globals.node },
  },
  {
    // A Node script whose `page.evaluate` callbacks are SERIALISED and run
    // inside the browser, where `document` is real. Untyped .mjs with no
    // project behind it, so every evaluate result is `any` by construction.
    files: ["scripts/pane-shots.mjs"],
    languageOptions: { globals: { ...globals.node, ...globals.browser } },
  },
  {
    // The fixture builder and the tests reach into XML shapes on purpose.
    files: ["test/**"],
    rules: { "@typescript-eslint/no-explicit-any": "off" },
  },
);
