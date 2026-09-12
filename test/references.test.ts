import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
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
const commandDrift = refs.commandDrift as () => string[];
const paneControlProblems = refs.paneControlProblems as (root?: string) => string[];
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

describe("the Commands block in CLAUDE.md", () => {
  /**
   * What a session is told it can run, against what `package.json` has.
   *
   * On 2026-09-12 the block was missing five scripts, and one of them —
   * `npm run previews` — is a step the Pages deploy runs to put the pictures
   * on the tiles. A command nobody knows about is a step nobody takes.
   */
  it("names every script, and only scripts that are there", () => {
    expect(commandDrift()).toEqual([]);
  });
});

describe("the controls the design record names", () => {
  /**
   * The gap this session kept finding by hand, made mechanical.
   *
   * Four behaviours built on 2026-09-12 — the jump, "Open all", "Move to a new
   * slide" and the pane remembering itself per deck — had all been described in
   * `docs/DESIGN.md` sections 4 and 6, approved by the owner, and never built.
   * Nothing went red for any of them. The lockstep rule holds the MANUAL to the
   * pane; nothing held the RECORD to it, and the record is written FIRST, so it
   * is the record that gets ahead of the code.
   *
   * Checked against the pane with its comments stripped, because a label
   * appears in a comment explaining why the pane does not have it — section 4
   * records that there is deliberately no "Close all" — and a raw sweep reads
   * that as built. Four guards in this family have gone red for exactly that
   * reason; this would have been the fifth.
   *
   * **Measured once, on 2026-09-12**: pointed at the repository as it stood at
   * `e34b379`, the commit before "Open all" was built, the sweep reported
   * exactly `"Open all"` and `"Move to a new slide"` and nothing else real. It
   * is written down here rather than re-run on every build. The case that did
   * re-run it rebuilt that tree with `git show`, which passes on a full clone
   * and fails on `actions/checkout`'s one-commit-deep one — so it measured the
   * checkout rather than the sweep. The two cases below take a fixture
   * instead, and hold more than that one did: the second is the
   * comment-stripping trap, which until now was only ever proven by hand.
   */
  it("draws every control the record quotes, or says why not", () => {
    expect(paneControlProblems()).toEqual([]);
  });

  /**
   * A fixture the sweep is pointed at, rather than the repository itself.
   *
   * `paneControlProblems` takes a root for this. Two files are the whole of
   * what it reads: a `docs/DESIGN.md` with sections 4 and 6, and the `.ts`
   * under `src/pane` and `src/host`. Both source directories have to exist —
   * `paneCode` calls `readdirSync` on each — and sections 4 and 6 have to carry
   * at least six quoted labels between them, or the sweep's own vacuity guard
   * fires and answers a complaint about itself instead of about the fixture.
   */
  function sweepOver(record: string, pane: string): string[] {
    const dir = mkdtempSync(join(tmpdir(), "ssf-slide-elements-record-"));
    try {
      mkdirSync(join(dir, "docs"), { recursive: true });
      mkdirSync(join(dir, "src/pane"), { recursive: true });
      mkdirSync(join(dir, "src/host"), { recursive: true });
      writeFileSync(join(dir, "docs/DESIGN.md"), record);
      writeFileSync(join(dir, "src/pane/render.ts"), pane);
      writeFileSync(join(dir, "src/host/nothing.ts"), "export const NOTHING = 0;\n");
      return paneControlProblems(dir);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }

  /** Sections 4 and 6 as the record writes them: control labels in quotes. */
  const RECORD =
    `## 4. The pane\n\n` +
    `- **Categories** start collapsed, with "Spread them out" beside the count.\n` +
    `- **Footer**: "Do it again", "Take it back" and "Put it somewhere else".\n` +
    `- A slide number says "Slide N" once the host confirms it.\n\n` +
    `## 5. Where an element lands\n\nNot swept.\n\n` +
    `## 6. Inserting\n\n` +
    `- The tile says "Working on it" while an insert runs.\n` +
    `- A part already in the deck offers "Take it off N slides".\n\n` +
    `## 7. Colours\n\nNot swept either.\n`;

  /** A pane drawing all of them but one, with the odd one out named in prose. */
  const PANE =
    `const a = button("Spread them out");\n` +
    `const b = button("Do it again");\n` +
    `const c = button("Take it back");\n` +
    `const d = \`Slide \${n}\`;\n` +
    `const e = "Working on it";\n` +
    `const f = \`Take it off \${n} slides\`;\n`;

  it("names a control the record describes and the pane does not draw", () => {
    // The sweep's whole job, on a fixture rather than on this repository — so
    // it runs on a shallow checkout, in a fresh worktree, out of a tarball.
    // The first version of this case rebuilt the real tree at `e34b379` with
    // `git show`, passed here and failed in CI, where `actions/checkout` clones
    // one commit deep. A gate that only works on the machine that wrote it is
    // the thing this file exists to prevent.
    const found = sweepOver(RECORD, PANE);
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('"Put it somewhere else"');
  });

  it("does not count a label the pane only mentions in a COMMENT", () => {
    // The trap four guards in this family have fallen into, and the reason the
    // sweep reads the pane through `withoutTsComments`. Section 4 really does
    // carry a sentence saying there is deliberately no "Close all", and a raw
    // read of the source counts that as the control being built.
    const excused = `${PANE}// "Put it somewhere else" is deliberately not offered here.\n`;
    const found = sweepOver(RECORD, excused);
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('"Put it somewhere else"');
  });
});
