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
const probeQuestions = refs.probeQuestions as () => number[];
const questionDrift = refs.questionDrift as () => string[];
const directoryRows = refs.directoryRows as () => { paths: string[]; files: string[] }[];
const directoryTableProblems = refs.directoryTableProblems as () => string[];
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

describe("the questions the probe asks", () => {
  /**
   * The probe and the two canonical lists, held equal.
   *
   * `docs/DESIGN.md` section 13 and `CLAUDE.md`'s "Open questions for the real
   * host" are the places a reader looks for what a round settles, and the
   * reader script is what a round is actually read through. On 2026-09-12 the
   * jump added a seventh question to the probe, to `docs/PROBE.md` and to the
   * reader, and to neither list — so a sheet could have answered a question no
   * document was waiting on, while `CLAUDE.md` said the snippet asked "all of
   * them". Same shape as `test/release.test.ts` holding the two workflows
   * equal: two lists that must agree, and nothing that made them.
   */
  it("are the questions both canonical lists carry", () => {
    expect(questionDrift()).toEqual([]);
  });

  it("were actually found, so an empty sweep cannot pass for agreement", () => {
    // The vacuity half. A reader whose headings stopped matching would report
    // no questions at all, and the case above would go green on nothing.
    expect(probeQuestions().length).toBeGreaterThan(5);
    expect(probeQuestions()).toContain(7);
  });
});

describe("the directory table in CLAUDE.md", () => {
  /**
   * What the memory file says each directory owns, against what is there.
   *
   * Read on 2026-09-12 it was two increments behind: `src/core/` promised the
   * splice was "next" while `src/core/splice/` had shipped, and `src/host/`
   * named three files out of eight. Nothing was checking it — `CLAUDE.md` is
   * read by one other test, for the dated-round-count rule. A row that has
   * chosen to enumerate its directory is held to enumerating all of it; a row
   * that describes one in prose is not made into a list.
   */
  it("names only paths that are there, and every source directory has a row", () => {
    expect(directoryTableProblems()).toEqual([]);
  });

  it("read rows at all, and rows that name their modules", () => {
    const rows = directoryRows();
    expect(rows.length, "no table rows were found").toBeGreaterThan(6);
    expect(rows.filter((row) => row.files.length > 0).length, "no row names any module").toBeGreaterThan(2);
  });
});
