import { describe, expect, it } from "vitest";
import { catalogueUrl, reportUrl, siteFrom } from "../src/host/links.js";

/**
 * What the two links out of the gear may carry.
 *
 * `docs/DESIGN.md` section 7 gives the pane a **Report a problem** link that
 * arrives at the support page with the build, the host and the platform already
 * filled in. That is a URL on the open web, written by code that is running
 * inside somebody's presentation — so the interesting property is not that the
 * three values arrive, it is that **nothing else can**.
 */

const site = { origin: "https://ssf-slide-elements.struktureretsundfornuft.dk" };

describe("the support link", () => {
  it("carries the build, the host and the platform", () => {
    const url = new URL(reportUrl(site, { build: "4faaae5", host: "PowerPoint", platform: "OfficeOnline" }));
    expect(url.origin).toBe(site.origin);
    expect(url.pathname).toBe("/support.html");
    expect(Object.fromEntries(url.searchParams)).toEqual({
      build: "4faaae5",
      host: "PowerPoint",
      platform: "OfficeOnline",
    });
  });

  it("leaves out what it does not know, rather than sending the word unknown", () => {
    // The page can then tell "not reported" from "reported as unknown", and a
    // pane opened outside a host — which is every time somebody looks at it in
    // a browser — does not produce a report full of placeholders.
    expect(reportUrl(site, {})).toBe(`${site.origin}/support.html`);
    expect(reportUrl(site, { build: "4faaae5" })).toBe(`${site.origin}/support.html?build=4faaae5`);
  });

  it("drops a value that is not what it claims to be", () => {
    // The host hands these over and the support page shows them. A value that
    // is not a commit, not one of Office's platform names, and not a host name
    // has no business travelling between the two — the allowlist is the whole
    // guard, and it is here rather than at the page because the page is not the
    // only thing that could read the URL.
    const url = new URL(
      reportUrl(site, {
        build: "<script>alert(1)</script>",
        host: "Neptune",
        platform: "https://elsewhere.example/steal",
      }),
    );
    expect([...url.searchParams.keys()]).toEqual([]);
  });

  it("cannot be made to carry anything from the deck", () => {
    // The strong version of the rule, stated as a test: the signature takes
    // three named fields, so a caller cannot add a fourth, and anything extra
    // handed over is ignored rather than passed through.
    const url = new URL(
      reportUrl(site, {
        build: "4faaae5",
        ...({ deck: "Q4 board pack.pptx", slide: "7", element: "white-box" } as Record<string, string>),
      }),
    );
    expect([...url.searchParams.keys()]).toEqual(["build"]);
    expect(url.toString()).not.toContain("board");
  });

  it("takes every platform Office names, and a long enough commit", () => {
    for (const platform of ["PC", "OfficeOnline", "Mac", "iOS", "Android", "Universal"]) {
      expect(reportUrl(site, { platform }), platform).toContain(`platform=${platform}`);
    }
    expect(reportUrl(site, { build: "a".repeat(40) })).toContain("build=aaaa");
    // Six characters is not a commit this repo ever shows: the header prints
    // seven.
    expect(reportUrl(site, { build: "4faaae" })).not.toContain("build");
  });
});

describe("the catalogue link", () => {
  it("is the catalogue page on the same site", () => {
    expect(catalogueUrl(site)).toBe(`${site.origin}/catalogue.html`);
  });
});

describe("which site the pane links to", () => {
  it("is the one the pane itself was served from", () => {
    // Never a hard-coded production address: a dev build links to the dev
    // origin, and nothing in the pane can send somebody to a site the add-in
    // did not come from.
    expect(siteFrom("https://ssf-slide-elements.struktureretsundfornuft.dk/taskpane.html")).toEqual(site);
    expect(siteFrom("https://localhost:3002/taskpane.html?x=1")).toEqual({ origin: "https://localhost:3002" });
  });

  it("answers nothing for a pane that was not served over the web", () => {
    expect(siteFrom("file:///C:/somewhere/taskpane.html")).toBeUndefined();
    expect(siteFrom("javascript:alert(1)")).toBeUndefined();
    expect(siteFrom("not a url at all")).toBeUndefined();
  });
});
