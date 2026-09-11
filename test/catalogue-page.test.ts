import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the harvest script.
import { catalogueHtml, dirOf, esc } from "../scripts/catalogue-page.mjs";

/**
 * The catalogue page on the site.
 *
 * `docs/DESIGN.md` section 3: every element with its picture and its name, for
 * browsing outside PowerPoint. It is generated at harvest and committed, so it
 * has the same two ways of going wrong every generated artifact here has —
 * edited by hand, or left behind when the decks changed — and the same gate:
 * regenerate it from the committed catalogue and compare.
 */

const INDEX = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
  version: string;
  sizes: Record<string, { categories: { key: string; name: string }[]; elements: { id: string; name: string }[] }>;
};
const PAGE = readFileSync("public/catalogue.html", "utf8");

describe("the committed page", () => {
  it("is what the generator makes from the committed catalogue, character for character", () => {
    // If this fails, either a deck changed and `npm run harvest` was not re-run,
    // or somebody edited the generated file. CI diffs it as well, on a fresh
    // harvest; this says the same thing without needing the decks.
    expect(catalogueHtml(INDEX.sizes, INDEX.version)).toBe(PAGE);
  });

  it("shows every element in both sizes", () => {
    for (const [size, catalogue] of Object.entries(INDEX.sizes)) {
      for (const element of catalogue.elements) {
        expect(PAGE, `${size} ${element.id}`).toContain(`id="${dirOf(size) as string}-${element.id}"`);
      }
    }
    // 117 per size today, and the assertion is against the index rather than
    // against 234, so adding an element to the decks does not need this edited.
    const total = Object.values(INDEX.sizes).reduce((n, c) => n + c.elements.length, 0);
    expect([...PAGE.matchAll(/<figure class="element"/g)]).toHaveLength(total);
  });

  it("names every category, with what it holds", () => {
    for (const [, catalogue] of Object.entries(INDEX.sizes)) {
      for (const category of catalogue.categories) {
        expect(PAGE, category.name).toContain(`>${esc(category.name) as string} <span class="count">`);
      }
    }
  });

  it("carries no script at all, which is what keeps it out of the security question", () => {
    // A page of 234 pictures on a public site does not need JavaScript, and
    // every line of it would be a line to audit against SECURITY.md's promise
    // that this project sends nothing anywhere.
    expect(PAGE).not.toMatch(/<script/i);
    expect(PAGE).not.toMatch(/\son\w+=/i);
  });

  it("points at pictures with a path relative to itself and a cache key", () => {
    // Relative, like the pane's own fetches: the page is served from the same
    // origin as the previews, and an absolute URL here would be the one thing
    // the security page says never happens.
    expect(PAGE).toContain(`src="catalogue/16x9/previews/`);
    expect(PAGE).toContain(`?v=${INDEX.version}`);
    expect(PAGE).not.toMatch(/src="https?:/);
  });
});

describe("the generator", () => {
  const one = (name: string): string =>
    catalogueHtml(
      {
        "16:9": {
          size: "16:9",
          categories: [{ key: "boxes", name: "White boxes" }],
          elements: [{ id: "a-box", name, category: { key: "boxes", name: "White boxes" } }],
        },
      },
      "v1",
    ) as string;

  it("escapes a name, because every name on the page comes out of the owner's deck", () => {
    // The deck is the owner's own, so this is not a hostile input — it is a
    // deck that one day holds an element called `Boxes & "quotes"`, which would
    // otherwise produce a page that is not valid HTML at all.
    const html = one(`Boxes & "quotes" <script>alert(1)</script>`);
    expect(html).toContain("Boxes &amp; &quot;quotes&quot; &lt;script&gt;");
    expect(html).not.toMatch(/<script/i);
  });

  it("puts the name in the alt text as well as under the picture", () => {
    // The picture IS the content here, so a reader who cannot see it gets the
    // name from the alt text rather than nothing. The pane's tiles do the
    // opposite — their pictures are decorative, because the name is beside them
    // as text and a screen reader announcing it twice is noise.
    expect(one("One box")).toContain(`alt="One box"`);
  });

  it("shows every member of a sized run, unlike the pane", () => {
    const html = catalogueHtml(
      {
        "16:9": {
          size: "16:9",
          categories: [{ key: "flow", name: "Process flows" }],
          elements: [1, 2, 3].map((n) => ({
            id: `flow-${n}`,
            name: `Process flow, ${n} boxes`,
            category: { key: "flow", name: "Process flows" },
            run: { key: "Process flow, N boxes", noun: "boxes", count: n },
          })),
        },
      },
      "v1",
    ) as string;
    for (const n of [1, 2, 3]) expect(html).toContain(`Process flow, ${n} boxes`);
    // And the count is NOT repeated beside the name: it is already in it.
    expect(html).not.toContain(`boxes <span class="run"`);
  });
});
