import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import * as refs from "../scripts/doc-refs.mjs";

/**
 * Where this repository's prose points, and whether anything is there.
 *
 * It explains itself by pointing — a comment says `docs/DESIGN.md` section 6, a
 * doc links the backlog, the manual links a heading inside itself — and until
 * this file none of those pointers was checked. Renumbering one heading turns
 * dozens of sentences into directions to the wrong place without a single thing
 * going red.
 *
 * Every case here asserts a COUNT as well as a verdict. A reference gate whose
 * pattern stops matching reports a clean sweep of nothing, and that is the
 * failure every guard in this repository that has ever been wrong shared.
 */

const brokenLinks = refs.brokenLinks as () => string[];
const brokenAnchors = refs.brokenAnchors as () => string[];
const brokenCitations = refs.brokenCitations as () => string[];
const markdownLinks = refs.markdownLinks as () => { anchor?: string }[];
const sectionCitations = refs.sectionCitations as () => unknown[];
const sectionsOf = refs.sectionsOf as (path: string) => Map<number, string>;
const anchorOf = refs.anchorOf as (heading: string) => string;

describe("where the prose points", () => {
  it("links only files that are there", () => {
    expect(brokenLinks()).toEqual([]);
    // The sweep is real: the docs carry dozens of relative links, and a pattern
    // that matched none of them would pass the line above.
    expect(markdownLinks().length, "no markdown links were found at all").toBeGreaterThan(30);
  });

  it("names only headings that are there", () => {
    expect(brokenAnchors()).toEqual([]);
    expect(markdownLinks().filter((link) => link.anchor).length, "no anchored links were found at all").toBeGreaterThan(
      20,
    );
  });

  it("cites only sections the document has", () => {
    expect(brokenCitations()).toEqual([]);
    // 130 of these when the gate was written, most of them in code comments.
    // They are the reference this project leans on hardest and the one nothing
    // was watching.
    expect(sectionCitations().length, "no section citations were found at all").toBeGreaterThan(100);
  });

  it("reads the numbered headings out of the design, rather than a list kept here", () => {
    // The gate's own input. If this stopped finding sections, every citation
    // would resolve against an empty map — which is the shape that fails open,
    // so it is asserted rather than assumed.
    const sections = sectionsOf("docs/DESIGN.md");
    expect(sections.size).toBeGreaterThan(10);
    expect(sections.get(1)).toBeTypeOf("string");
    expect(sectionsOf("docs/does-not-exist.md").size).toBe(0);
  });

  it("builds an anchor the way a renderer does", () => {
    expect(anchorOf("7. The gear, and what it holds")).toBe("7-the-gear-and-what-it-holds");
    expect(anchorOf("What is stored on your device")).toBe("what-is-stored-on-your-device");
  });
});
