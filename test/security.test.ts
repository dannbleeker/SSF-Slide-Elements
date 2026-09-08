import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The claims on the front of SECURITY.md, as tests.
 *
 * That page says the add-in makes no network calls and never writes markup.
 * Both are properties of the SOURCE, so both can be read off it — and a
 * security page whose claims nothing re-checks is the failure the page's own
 * preamble warns about. The package-level cases (values as text, part paths
 * that stay inside the package) arrive with the package layer and join this
 * file then.
 */
const sources = (): string[] => {
  const out: string[] = [];
  const walk = (dir: string): void => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = `${dir}/${entry.name}`;
      if (entry.isDirectory()) walk(path);
      else if (/\.(ts|html)$/.test(entry.name)) out.push(path);
    }
  };
  walk("src");
  return out;
};

const code = (path: string): string =>
  readFileSync(path, "utf8")
    // Comments are stripped, because this repo's files explain themselves and
    // several of them name these APIs in prose.
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/^\s*\/\/.*$/gm, "");

describe("the claims on the front of SECURITY.md are executable", () => {
  it("reads the source at all", () => {
    // The vacuity guard: an empty file list would satisfy every assertion
    // below forever.
    expect(sources().length).toBeGreaterThan(5);
  });

  it("makes no network call", () => {
    for (const path of sources()) {
      const src = code(path);
      for (const call of ["fetch(", "XMLHttpRequest", "WebSocket", "sendBeacon"]) {
        expect(src.includes(call), `${path} calls ${call}`).toBe(false);
      }
    }
  });

  it("writes text, never markup", () => {
    for (const path of sources()) {
      const src = code(path);
      for (const sink of ["outerHTML", "insertAdjacentHTML", "document.write", "eval(", "new Function"]) {
        expect(src.includes(sink), `${path} uses ${sink}`).toBe(false);
      }
      // `innerHTML` is READ once, in the page's own no-Office fallback, to ask
      // whether anything has been drawn yet. Assigning to it is what this
      // forbids.
      expect(/innerHTML\s*=/.test(src), `${path} assigns innerHTML`).toBe(false);
    }
  });
});
