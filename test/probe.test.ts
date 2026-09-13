import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";
import JSZip from "jszip";
import ts from "typescript";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsComments, withoutTsProse } from "../scripts/without-prose.mjs";
import { Pkg } from "../src/core/pptx/pkg.js";
import { P_NS, element, elements } from "../src/core/pptx/xml.js";
import { API_FLOOR } from "../src/host/capability.js";
import {
  API_SETS,
  PROBE_MARKER,
  PROBE_TAG,
  PROBE_UNDO_VALUE,
  exportPartsVerdict,
  floorLine,
  insertVerdict,
  insertionBlame,
  leftBehind,
  masterVerdict,
  notAsked,
  orderVerdict,
  pruningReading,
  jumpProbeVerdict,
  selectionVerdict,
  summarizeParts,
  targetAddedVerdict,
  timingLine,
  undoVerdict,
  type InsertVerdict,
  type PartsSummary,
} from "../src/host/probe.js";

/**
 * These read the ARTIFACT, not a recipe for it.
 *
 * The probe's whole value is that its arms differ in exactly one way each, and
 * a builder that quietly produced identical decks would give an answer sheet
 * that looks complete and says nothing. Reading the generated snippet means
 * there is no second copy of the recipe to drift from the first; CI rebuilds
 * it and diffs, so it also cannot be stale.
 */
const snippet = readFileSync("probe/probe-snippet.ts", "utf8");
const code = withoutTsComments(snippet) as string;

function deckFromSnippet(name: string): string {
  // Tolerant of whitespace on purpose: a formatter that wrapped the
  // assignment once broke a sibling's version of this silently.
  const match = new RegExp(`const ${name}\\s*=\\s*"([A-Za-z0-9+/=]+)"`).exec(snippet);
  if (!match?.[1]) throw new Error(`the snippet has no ${name}`);
  return match[1];
}

const DECKS = ["LISTED_DECK", "PRUNED_DECK", "UNLISTED_DECK", "SINGLE_DECK", "UNDO_DECK"];

describe("the probe's fixture decks", () => {
  it("lists two slides in the listed deck and one in every other", async () => {
    for (const [name, want] of [
      ["LISTED_DECK", 2],
      ["PRUNED_DECK", 1],
      ["UNLISTED_DECK", 1],
      ["SINGLE_DECK", 1],
      ["UNDO_DECK", 1],
    ] as const) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      expect(await pkg.slidePaths(), name).toHaveLength(want);
    }
  });

  it("keeps the second slide's PART in both pruned decks, which is the whole question", async () => {
    // Question 1 is whether an unlisted part is tolerated on the way in. A
    // pruned deck that had also lost the part would test nothing.
    for (const name of ["PRUNED_DECK", "UNLISTED_DECK"]) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      expect(pkg.has("ppt/slides/slide2.xml"), name).toBe(true);
      expect(await pkg.contentTypeOf("ppt/slides/slide2.xml"), name).toContain("slide+xml");
    }
  });

  it("differs between the two pruned decks in the relationship alone", async () => {
    const pruned = await Pkg.open(deckFromSnippet("PRUNED_DECK"));
    const unlisted = await Pkg.open(deckFromSnippet("UNLISTED_DECK"));
    expect(await pruned.relatedParts("ppt/presentation.xml")).not.toContain("ppt/slides/slide2.xml");
    expect(await unlisted.relatedParts("ppt/presentation.xml")).toContain("ppt/slides/slide2.xml");
  });

  it("gives every slide its own creation id", async () => {
    const ids = new Set<string>();
    for (const name of DECKS) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      for (const path of await pkg.slidePaths()) {
        const m = /p14:creationId[^>]*val="(\d+)"/.exec(await pkg.text(path));
        expect(m?.[1], `${name} ${path}`).toBeDefined();
        ids.add(m?.[1] ?? "");
      }
    }
    // Two slides in the listed deck plus one each in single and undo: four distinct ids across the arms.
    expect(ids.size).toBeGreaterThanOrEqual(4);
  });

  it("tags the undo deck's slide in the PACKAGE, which is what the second run looks for", async () => {
    const pkg = await Pkg.open(deckFromSnippet("UNDO_DECK"));
    const slide = await pkg.doc("ppt/slides/slide1.xml");
    const tags = element(slide, P_NS, "tags");
    expect(tags, "the slide has no <p:tags> reference").toBeDefined();
    const rId = tags?.getAttribute("r:id") ?? "";
    const target = await pkg.relTarget("ppt/slides/slide1.xml", rId);
    expect(target).toBe("ppt/tags/tag1.xml");
    const tagLst = await pkg.doc("ppt/tags/tag1.xml");
    const tag = elements(tagLst, P_NS, "tag").find((t) => t.getAttribute("name") === PROBE_TAG);
    expect(tag?.getAttribute("val")).toBe(PROBE_UNDO_VALUE);
    expect(await pkg.contentTypeOf("ppt/tags/tag1.xml")).toContain("tags+xml");
  });

  it("pins every zip entry's timestamp, so the build is the same on any machine at any hour", async () => {
    // JSZip writes dates in LOCAL time and the engine re-files an edited part
    // with the clock, so the first CI run diffed a snippet that differed from
    // the committed one in every deck. `stableZip` overwrites the DOS fields;
    // JSZip reads them back through Date.UTC, so the UTC components are what
    // carry the pinned value. Read as local components this was 14:00 on the
    // owner's UTC+2 box and 12:00 on CI, and the suite was red at home only.
    for (const name of DECKS) {
      const zip = await JSZip.loadAsync(Buffer.from(deckFromSnippet(name), "base64"));
      const entries = Object.values(zip.files).filter((f) => !f.dir);
      expect(entries.length, name).toBeGreaterThan(5);
      for (const entry of entries) {
        const d = entry.date;
        const stamp = [
          d.getUTCFullYear(),
          d.getUTCMonth(),
          d.getUTCDate(),
          d.getUTCHours(),
          d.getUTCMinutes(),
          d.getUTCSeconds(),
        ];
        expect(stamp, `${name} ${entry.name}`).toEqual([2026, 8, 8, 12, 0, 0]);
      }
    }
  });

  it("carries a theme with all three of its required children", async () => {
    // CT_BaseStyles requires clrScheme, fontScheme and fmtScheme. A sibling's
    // first real sheet came back InvalidArgument from every insert because
    // this part was `<a:themeElements/>`.
    const pkg = await Pkg.open(deckFromSnippet("LISTED_DECK"));
    const theme = await pkg.text("ppt/theme/theme1.xml");
    for (const child of ["clrScheme", "fontScheme", "fmtScheme"]) expect(theme).toContain(`<a:${child} `);
  });

  it("names every part it declares a content type for, in every deck", async () => {
    // A content-type Override for a part that is not in the zip is a package
    // no reader will open, and it is the easy mistake when pruning.
    for (const name of DECKS) {
      const pkg = await Pkg.open(deckFromSnippet(name));
      const types = await pkg.text("[Content_Types].xml");
      const parts = [...types.matchAll(/<Override PartName="\/([^"]+)"/g)].map((m) => m[1] ?? "");
      expect(parts.length, name).toBeGreaterThan(4);
      for (const part of parts) expect(pkg.has(part), `${name} declares ${part}`).toBe(true);
    }
  });
});

describe("the probe snippet", () => {
  it("bounds every host call, because a stall here is death rather than slowness", () => {
    expect(code).toContain("withTimeout");
  });

  it("cleans up by position and never by id", () => {
    // A slide this run added does not round-trip through slides.getItem(id)
    // on the web, and a sibling's by-id clean-up reported 45 successful
    // deletes having removed nothing.
    expect(code).toContain("getItemAt(i).delete()");
    expect(code).not.toContain("getItem(id).delete");
  });

  it("passes a targetSlideId, so the probe's slides do not land in front of the user's", () => {
    expect(code).toContain("targetSlideId");
  });

  it("carries a floor that stops the sweep reaching the user's own slides", () => {
    expect(code).toContain("from < deckAtStart");
  });

  it("never sets a SHAPE selection, and puts the slide selection back after the jump arm", () => {
    // `setSelectedShapes` wedges the web host's selection subsystem and is
    // never called. `setSelectedSlides` IS called, by question 7 alone, on the
    // sibling's measurement (src/host/jump.ts); the arm restores what was
    // selected. Against the CODE: the prose names both calls.
    const prose = withoutTsProse(snippet) as string;
    expect(prose).not.toContain("setSelectedShapes");
    expect(code).toContain("getSelectedSlides");
    const arm = code.slice(code.indexOf("async function jumpProbe"), code.indexOf("async function exportProbe"));
    expect(arm).toContain("setSelectedSlides([first])");
    expect(arm).toContain("setSelectedSlides(before)");
    // Nowhere else writes a selection.
    expect(code.split("setSelectedSlides(").length - 1).toBe(2);
  });

  it("asks its control arm BEFORE it inserts anything of its own", () => {
    // The control inserts the presentation's own bytes. Run later it would be
    // inserting the probe's slides back too.
    const control = code.indexOf("answers.insertOwn ");
    const first = code.indexOf("answers.insertListed");
    expect(control).toBeGreaterThan(-1);
    expect(control).toBeLessThan(first);
  });

  it("leaves the Ctrl+Z slide LAST, after the final sweep, and never two in a row", () => {
    const sweep = code.indexOf("answers.sweep =");
    const undo = code.indexOf("answers.undo =");
    expect(sweep).toBeGreaterThan(-1);
    expect(undo).toBeGreaterThan(sweep);
    // A second run is one that found the previous run's slide OR its marker:
    // the slide alone was the test once, and a second run after a successful
    // Ctrl+Z looked like a first and left a slide of its own.
    expect(code).toContain("const secondRun = undoAtStart.found === true || marker.found");
    expect(code).toContain("const leaveBehind = !secondRun");
  });

  it("writes the marker BEFORE the slide it leaves, and touches the settings nowhere after", () => {
    // The document's settings are outside the undo stack, which is what lets
    // the marker outlive the slide. A save AFTER the insert is reported to
    // disable undo on Excel (office-js#3141), which would take the user's one
    // Ctrl+Z away from the insert this question is about.
    const leave = code.indexOf("insertDeck(UNDO_DECK");
    const mark = code.indexOf("await writeMarker(");
    expect(mark).toBeGreaterThan(-1);
    expect(mark).toBeLessThan(leave);
    expect(code.indexOf("Marker(", leave)).toBe(-1);
    expect(code.indexOf("saveAsync", leave)).toBe(-1);
  });

  it("clears the marker on the second run, so the run after is a first run again", () => {
    const clear = code.indexOf("await clearMarker()");
    const start = code.indexOf("answers.deckAtStart");
    expect(clear).toBeGreaterThan(-1);
    expect(clear).toBeLessThan(start);
  });

  it("removes a previous run's slide by position before measuring anything", () => {
    const removal = code.indexOf("undoAtStart.index as number).delete()");
    const start = code.indexOf("answers.deckAtStart");
    expect(removal).toBeGreaterThan(-1);
    expect(removal).toBeLessThan(start);
  });

  it("compares the API's slide order with the file's <p:sldIdLst>", () => {
    expect(code).toContain("sldIdLstMatches");
    expect(code).toContain("ppt/presentation.xml");
    expect(code).toContain('split("#")[0]');
  });

  it("aims the order arm at a slide it just added, and falls back when that is refused", () => {
    expect(code).toContain("targetAdded");
    expect(code).toContain("sweepAfterRefusal");
  });

  it("probes for each method rather than trusting the requirement set alone", () => {
    for (const set of ['supports("1.3")', 'supports("1.5")', 'supports("1.10")']) expect(code).toContain(set);
  });

  it("carries the same constants as the engine", () => {
    expect(snippet).toContain(`const PROBE_TAG = "${PROBE_TAG}"`);
    expect(snippet).toContain(`const PROBE_MARKER = "${PROBE_MARKER}"`);
    expect(snippet).toContain(`const API_FLOOR = "${API_FLOOR}"`);
    expect(snippet).toContain(`const API_SETS = ${JSON.stringify(API_SETS)}`);
  });

  it("summarises a part list the way the reader does", () => {
    // The snippet cannot import the engine, so it carries a typed copy of
    // `summarizeParts`. Transpiled and run here against the real one, so a
    // regex changed in one place goes red rather than making the two ends of
    // the comparison disagree about what a comment part is.
    const at = snippet.indexOf("function summarizeParts");
    const body = snippet.slice(at, snippet.indexOf("\n}\n", at) + 3);
    const js = ts.transpileModule(body, { compilerOptions: { target: ts.ScriptTarget.ES2020 } }).outputText;
    // Run in a bare context: the snippet's copy must not lean on anything of this file's.
    const theirs = runInNewContext(`${js}; summarizeParts`) as (names: string[]) => PartsSummary;
    const names = [
      "[Content_Types].xml",
      "ppt/presentation.xml",
      "ppt/slides/slide1.xml",
      "ppt/slides/slide12.xml",
      "ppt/slides/_rels/slide1.xml.rels",
      "ppt/slideMasters/slideMaster1.xml",
      "ppt/slideLayouts/slideLayout3.xml",
      "ppt/theme/theme1.xml",
      "ppt/theme/theme2.xml",
      "ppt/comments/comment1.xml",
      "ppt/modernComments/modernComment_1.xml",
      "ppt/authors.xml",
      "ppt/media/image1.png",
      "ppt/notesSlides/notesSlide1.xml",
    ];
    expect(theirs(names)).toEqual(summarizeParts(names));
    expect(summarizeParts(names)).toEqual({
      total: 14,
      slides: 2,
      masters: 1,
      layouts: 1,
      themes: 2,
      comments: 2,
      authors: true,
      media: 1,
    });
  });

  it("typechecks against the real Office.js types", () => {
    // The snippet is outside tsconfig's include and is pasted into an editor
    // that will run it before anyone reads it, so nothing else would catch a
    // misspelled option key or a call that does not exist. `--ignoreConfig`
    // is required from TypeScript 6, which errors on a file named beside a
    // tsconfig.json.
    let output = "";
    let failed = false;
    try {
      execFileSync(
        process.execPath,
        [
          "./node_modules/typescript/bin/tsc",
          "--noEmit",
          "--ignoreConfig",
          "--lib",
          "es2020,dom",
          "--types",
          "office-js",
          "--skipLibCheck",
          "probe/probe-snippet.ts",
        ],
        { encoding: "utf8", stdio: "pipe" },
      );
    } catch (e) {
      failed = true;
      const err = e as { stdout?: string; stderr?: string };
      output = `${err.stdout ?? ""}${err.stderr ?? ""}`.trim();
    }
    expect(failed, `tsc rejected the snippet:\n${output}`).toBe(false);
  }, 60000);

  it("calls its entry point at the top level, and never reaches for jQuery or a #run button", () => {
    // Script Lab's SAMPLE snippet ends `$("#run").click(...)`, and a blank
    // snippet has neither jQuery nor that button. A sibling lost a round trip
    // to exactly that.
    expect(code).toMatch(/^\s*main\(\)/m);
    expect(code).not.toMatch(/\$\s*\(/);
    expect(code).not.toContain("#run");
  });
});

describe("every arm the probe collects is READ", () => {
  // An arm nobody reads is the same nothing as an arm nobody asks, and worse,
  // because the sheet says it is there. Both halves come from source.
  const reader = withoutTsComments(readFileSync("scripts/read-answers.mjs", "utf8")) as string;
  const collected = [...new Set([...code.matchAll(/\banswers\.(\w+)\s*=/g)].map((m) => m[1] ?? ""))];

  it("finds arms to check", () => {
    expect(collected.length).toBeGreaterThan(10);
  });

  it.each(collected)("reads %s", (arm) => {
    expect(reader, `${arm} is collected and never read`).toContain(`sheet.${arm}`);
  });
});

describe("CI keeps the committed snippet equal to a fresh build", () => {
  it("rebuilds and diffs it in the gate", () => {
    const ci = readFileSync(".github/workflows/ci.yml", "utf8");
    expect(ci).toContain("npm run probe");
    expect(ci).toContain("git diff --exit-code -- probe/probe-snippet.ts");
    // test/release.test.ts holds the Pages gate to the same list.
  });
});

// ---------------------------------------------------------------------------
// The verdicts: every reading the reader prints, decided over plain values.
// ---------------------------------------------------------------------------

describe("insertVerdict", () => {
  it("grades by the delta, never by the absence of an error", () => {
    expect(insertVerdict({ before: 3, after: 5, expected: 2 })).toMatchObject({ verdict: "yes", landed: 2 });
    expect(insertVerdict({ before: 3, after: 3, expected: 2 })).toMatchObject({ verdict: "no", landed: 0 });
    expect(insertVerdict({ before: 3, after: 4, expected: 2 }).detail).toContain("partial");
  });

  it("believes a landed delta over a late error", () => {
    const v = insertVerdict({ before: 3, after: 5, expected: 2, error: "gave up waiting" });
    expect(v.verdict).toBe("yes");
    expect(v.detail).toContain("budget");
  });

  it("names a throw, and how much landed anyway", () => {
    expect(insertVerdict({ before: 3, after: 3, expected: 2, error: "InvalidArgument" })).toMatchObject({
      verdict: "threw",
      detail: "the call threw: InvalidArgument",
    });
    expect(insertVerdict({ before: 3, after: 4, expected: 2, error: "x" }).detail).toContain(
      "1 slide(s) landed anyway",
    );
  });

  it("will not call more-than-listed or a shrink a partial insert", () => {
    expect(insertVerdict({ before: 3, after: 5, expected: 1 })).toMatchObject({ verdict: "unknown" });
    expect(insertVerdict({ before: 3, after: 5, expected: 1 }).detail).toContain("more than the slide list");
    expect(insertVerdict({ before: 3, after: 2, expected: 1 }).detail).toContain("SHRANK");
  });

  it("says NOT ASKED for an arm the sheet does not carry", () => {
    expect(notAsked("x arm")).toMatchObject({ verdict: "unknown", landed: 0 });
    expect(notAsked("x arm").detail).toContain("NOT ASKED");
  });
});

describe("insertionBlame", () => {
  it("separates our package from the host by the control arm", () => {
    expect(insertionBlame("yes", "no")).toContain("works");
    expect(insertionBlame("threw", "yes")).toContain("OURS");
    expect(insertionBlame("threw", "unknown")).toContain("CANNOT TELL");
    expect(insertionBlame("threw", "threw")).toContain("THE HOST");
  });
});

const arm = (verdict: InsertVerdict["verdict"], landed = 0): InsertVerdict => ({
  verdict,
  landed,
  detail: `${verdict} detail`,
});

describe("pruningReading", () => {
  it("reads nothing about pruning when the listed deck itself did not land", () => {
    expect(pruningReading(arm("threw"), arm("yes", 1), arm("yes", 1))).toContain("NOT ANSWERED");
  });

  it("names which pruning the host takes", () => {
    expect(pruningReading(arm("yes", 2), arm("yes", 1), arm("yes", 1))).toContain("Both prunings");
    expect(pruningReading(arm("yes", 2), arm("yes", 1), arm("threw"))).toContain("drop the relationship");
    expect(pruningReading(arm("yes", 2), arm("threw"), arm("yes", 1))).toContain("Unexpected");
    expect(pruningReading(arm("yes", 2), arm("threw"), arm("no"))).toContain("Pkg.removeSlide");
  });
});

describe("masterVerdict", () => {
  it("cannot say without a master count or without a landed insert", () => {
    expect(masterVerdict({ landed: 3, formatting: "KeepSourceFormatting" }).detail).toContain("NOT ASKED");
    expect(masterVerdict({ mastersBefore: 1, mastersAfter: 2, landed: 0, formatting: "x" }).detail).toContain(
      "landed nothing",
    );
  });

  it("reads an added master, an unchanged count, and a fall", () => {
    expect(
      masterVerdict({ mastersBefore: 1, mastersAfter: 2, landed: 3, formatting: "KeepSourceFormatting" }),
    ).toMatchObject({ verdict: "yes" });
    expect(masterVerdict({ mastersBefore: 1, mastersAfter: 1, landed: 3, formatting: "x" })).toMatchObject({
      verdict: "no",
    });
    expect(masterVerdict({ mastersBefore: 2, mastersAfter: 1, landed: 3, formatting: "x" })).toMatchObject({
      verdict: "unknown",
    });
  });

  it("cannot say when the sheet carries one master count and not the other", () => {
    // A sheet is copied out of the host by hand, so half an arm is a real
    // shape. Either half missing is NOT ASKED, never a count against nothing.
    expect(masterVerdict({ mastersBefore: 1, landed: 3, formatting: "x" }).detail).toContain("NOT ASKED");
    expect(masterVerdict({ mastersAfter: 2, landed: 3, formatting: "x" }).detail).toContain("NOT ASKED");
  });

  it("reads a ONE-slide insert, which is the smallest that can add a master", () => {
    // The bar is "landed nothing", not "landed little": one slide is an insert
    // that happened and its master count is evidence.
    expect(masterVerdict({ mastersBefore: 1, mastersAfter: 2, landed: 1, formatting: "x" })).toMatchObject({
      verdict: "yes",
    });
    expect(masterVerdict({ mastersBefore: 1, mastersAfter: 2, landed: 0, formatting: "x" }).detail).toContain(
      "landed nothing",
    );
  });
});

describe("orderVerdict", () => {
  const P1 = "259#1";
  const P2 = "260#2";
  const S = "261#3";

  it("names a throw and an arm that did not read all three positions", () => {
    expect(orderVerdict({ error: "boom" })).toMatchObject({ verdict: "threw" });
    expect(orderVerdict({ afterFirst: [P1, P2] }).detail).toContain("NOT ASKED");
  });

  it("says yes when S lands after its target and the delete leaves [S, P2]", () => {
    const targeted = orderVerdict({
      afterFirst: [P1, P2],
      targetAdded: { ok: true },
      afterSecond: [P1, S, P2],
      deletedIndex: 3,
      afterDelete: [S, P2],
    });
    expect(targeted.verdict).toBe("yes");
    expect(targeted.detail).toContain("just added");
    const fallback = orderVerdict({
      afterFirst: [P1, P2],
      targetAdded: { ok: false, error: "SlideNotFound" },
      afterSecond: [S, P1, P2],
      deletedIndex: 4,
      afterDelete: [S, P2],
    });
    expect(fallback.verdict).toBe("yes");
    expect(fallback.detail).toContain("user's last slide");
  });

  it("separates a misplaced insert from a wrong delete", () => {
    expect(
      orderVerdict({ afterFirst: [P1, P2], targetAdded: { ok: true }, afterSecond: [S, P1, P2], afterDelete: [S, P2] })
        .detail,
    ).toContain("did not place it");
    expect(
      orderVerdict({ afterFirst: [P1, P2], targetAdded: { ok: true }, afterSecond: [P1, S, P2], afterDelete: [P1, P2] })
        .detail,
    ).toContain("positional delete left");
  });

  it("says S never landed when no new id appears", () => {
    expect(orderVerdict({ afterFirst: [P1, P2], afterSecond: [P1, P2, P1], afterDelete: [P1, P2] })).toMatchObject({
      verdict: "no",
    });
  });

  it("counts all three reads, and says NOT ASKED when any one of them is short", () => {
    // Each position count is checked on its own: a sheet with the right number
    // of ids in two of the three reads has still not asked the question.
    expect(orderVerdict({}).detail).toContain("(0, 0, 0 ids)");
    const shortMiddle = orderVerdict({ afterFirst: [P1, P2], afterSecond: [P1, S], afterDelete: [S, P2] });
    expect(shortMiddle).toMatchObject({ verdict: "unknown" });
    expect(shortMiddle.detail).toContain("(2, 2, 2 ids)");
  });
});

describe("targetAddedVerdict", () => {
  it("reads whether a just-added slide can be a target", () => {
    expect(targetAddedVerdict({ targetAdded: { ok: true } })).toMatchObject({ verdict: "yes" });
    expect(targetAddedVerdict({ targetAdded: { ok: false, error: "SlideNotFound" } }).detail).toContain(
      "SlideNotFound",
    );
    expect(targetAddedVerdict({ targetAdded: { ok: false } }).detail).toContain("no error text");
    expect(targetAddedVerdict({})).toMatchObject({ verdict: "unknown" });
    expect(targetAddedVerdict({ error: "boom" })).toMatchObject({ verdict: "threw" });
  });
});

describe("selectionVerdict", () => {
  it("cannot say on a host without the call, on a throw, or with nothing selected", () => {
    expect(selectionVerdict({ supported: false }).detail).toContain("1.5");
    expect(selectionVerdict({ supported: true, error: "x" })).toMatchObject({ verdict: "threw" });
    expect(selectionVerdict({ supported: true, selectedIds: [] }).detail).toContain("answered nothing");
  });

  it("says no when a selected id is at no position", () => {
    expect(
      selectionVerdict({ supported: true, selectedIds: ["a"], selectedIndexes: [-1], positionalRead: 5 }),
    ).toMatchObject({ verdict: "no" });
  });

  it("reports the slide number and whether the file's order agrees", () => {
    const yes = selectionVerdict({
      supported: true,
      selectedIds: ["257#9"],
      selectedIndexes: [1],
      deckSize: 3,
      sldIdLstMatches: true,
      sldIdLstEntries: 3,
    });
    expect(yes.verdict).toBe("yes");
    expect(yes.detail).toContain("slide number(s) 2 of 3");
    expect(
      selectionVerdict({ supported: true, selectedIds: ["a"], selectedIndexes: [0], sldIdLstMatches: false }),
    ).toMatchObject({ verdict: "no" });
    expect(selectionVerdict({ supported: true, selectedIds: ["a"], selectedIndexes: [0] }).detail).toContain(
      "could not be compared",
    );
  });

  it("reads a sheet a hand copy left a field short rather than throwing on it", () => {
    // Sheets are pasted out of the pane by hand (blob downloads are blocked in
    // WebView2), so a missing field is an input this reader meets, not a bug.
    expect(selectionVerdict({ supported: true }).detail).toContain("answered nothing");
    expect(selectionVerdict({ supported: true, selectedIds: ["a"] }).detail).toContain("1 slide(s) selected");
  });
});

describe("exportPartsVerdict", () => {
  const parts = (comments: number, authors: boolean, total = 40): PartsSummary => ({
    total,
    slides: 3,
    masters: 1,
    layouts: 11,
    themes: 1,
    comments,
    authors,
    media: 0,
  });

  it("cannot say without the call, after a throw, or without both lists", () => {
    expect(exportPartsVerdict({ supported: false }).detail).toContain("getFileAsync is the only read");
    expect(exportPartsVerdict({ supported: true, error: "x" })).toMatchObject({ verdict: "threw" });
    expect(exportPartsVerdict({ supported: true }).detail).toContain("no part lists");
  });

  it("says NOT ASKED on a deck with nothing to drop", () => {
    expect(
      exportPartsVerdict({ supported: true, source: parts(0, false), exported: parts(0, false, 38) }).detail,
    ).toContain("Re-run on a deck with comments");
  });

  it("names what the export dropped", () => {
    const v = exportPartsVerdict({ supported: true, source: parts(4, true), exported: parts(0, false, 35) });
    expect(v.verdict).toBe("yes");
    expect(v.detail).toContain("ppt/authors.xml and 4 comment part(s)");
  });

  it("says no when the comments came through", () => {
    expect(
      exportPartsVerdict({ supported: true, source: parts(4, true), exported: parts(4, true), missing: ["a"] }).detail,
    ).toContain("1 other part(s)");
    expect(exportPartsVerdict({ supported: true, source: parts(4, true), exported: parts(4, true) })).toMatchObject({
      verdict: "no",
    });
  });

  it("cannot say when the sheet carries one part list and not the other", () => {
    expect(exportPartsVerdict({ supported: true, source: parts(4, true) }).detail).toContain("no part lists");
    expect(exportPartsVerdict({ supported: true, exported: parts(4, true) }).detail).toContain("no part lists");
  });

  it("names a single loss on its own, without the other half of the sentence", () => {
    // office-js#6867 names two casualties; a host that drops one of them is
    // still a host that drops. Each is read alone as well as together.
    const authorsOnly = exportPartsVerdict({ supported: true, source: parts(0, true), exported: parts(0, false) });
    expect(authorsOnly.verdict).toBe("yes");
    expect(authorsOnly.detail).toContain("DROPS ppt/authors.xml —");
    const commentsOnly = exportPartsVerdict({ supported: true, source: parts(4, false), exported: parts(0, false) });
    expect(commentsOnly.verdict).toBe("yes");
    expect(commentsOnly.detail).toContain("DROPS 4 comment part(s) —");
  });
});

describe("timingLine", () => {
  it("prints size, time and rate, or why it cannot", () => {
    expect(timingLine("x", undefined)).toBe("x: not asked");
    expect(timingLine("x", { error: "boom" })).toBe("x: threw — boom");
    expect(timingLine("x", { ms: 5 })).toBe("x: not measured");
    expect(timingLine("x", { ms: 500, bytes: 1048576 })).toBe("x: 1.00 MB in 500 ms (2.0 MB/s)");
    expect(timingLine("x", { ms: 0, bytes: 1048576 })).toBe("x: 1.00 MB in 0 ms");
  });

  it("counts a megabyte as 1024 x 1024 bytes", () => {
    // The MB/s figures in docs/DESIGN.md section 15 (a 14.13 MB deck in 2816
    // ms, 4.6 to 5.0 MB/s) came off this line, so the divisor is part of the
    // measurement. 131072 bytes is exactly an eighth of a MiB: 0.13 against
    // 1048576, and 0.12 against any larger divisor.
    expect(timingLine("x", { ms: 125, bytes: 131072 })).toBe("x: 0.13 MB in 125 ms (1.0 MB/s)");
  });

  it("prints a rate for every measured time, and skips it only at zero", () => {
    // The guard is division by zero and nothing else: 1 ms is a rate.
    expect(timingLine("x", { ms: 1, bytes: 1048576 })).toBe("x: 1.00 MB in 1 ms (1000.0 MB/s)");
    expect(timingLine("x", { ms: 0, bytes: 1048576 })).toBe("x: 1.00 MB in 0 ms");
  });
});

describe("undoVerdict", () => {
  it("names a throw", () => {
    expect(undoVerdict({ error: "x" })).toMatchObject({ verdict: "threw" });
  });

  it("falls back to counts when tags cannot be read", () => {
    expect(undoVerdict({ foundAtStart: "unsupported" }).detail).toContain("NOT ASKED");
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 3, previousDeckAtEnd: 4 })).toMatchObject({
      verdict: "yes",
    });
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 4, previousDeckAtEnd: 4 })).toMatchObject({
      verdict: "no",
    });
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 2, previousDeckAtEnd: 4 })).toMatchObject({
      verdict: "unknown",
    });
  });

  it("will not compare counts with one end of the pair missing", () => {
    // Both ends or nothing: a delta against an absent number is arithmetic on
    // undefined, and this file says NOT ASKED rather than grading nothing.
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 3 }).detail).toContain("NOT ASKED");
    expect(undoVerdict({ foundAtStart: "unsupported", previousDeckAtEnd: 4 }).detail).toContain("NOT ASKED");
  });

  it("reads the tagged slide's presence on the second run", () => {
    expect(undoVerdict({ foundAtStart: true })).toMatchObject({ verdict: "no" });
    expect(undoVerdict({ foundAtStart: false, deckAtStart: 3, previousDeckAtEnd: 4 })).toMatchObject({
      verdict: "yes",
    });
    expect(undoVerdict({ foundAtStart: false, deckAtStart: 3, previousDeckAtEnd: 6 }).detail).toContain(
      "something else happened",
    );
  });

  it("says NOT YET on a first run, and NOT ASKED on a lone second sheet without a marker", () => {
    expect(undoVerdict({ foundAtStart: false, leftBehind: true }).detail).toContain("NOT YET");
    expect(undoVerdict({ foundAtStart: false, leftBehind: false }).detail).toContain("NOT ASKED");
  });

  it("says NOT ASKED when the slide is gone and only ONE of the two counts is on the sheet", () => {
    // The tag arm needs BOTH counts to subtract. Every case above carries both
    // or neither, and neither is what the `&&` reads the same way as an `||` —
    // so the second `&&` on that line survived the sweep, and survived it
    // twice: the first full run reported it killed, which was the sweep's own
    // whole-suite re-check failing for a reason that was not the mutation.
    //
    // What it costs: with `||` the arm is entered on one count alone and
    // subtracts against `undefined`, so the sheet reads "the deck changed by
    // NaN slide(s)" — a sentence that goes into `docs/host-answers/` as this
    // add-in's answer to whether Ctrl+Z reverts an insert.
    for (const half of [{ previousDeckAtEnd: 4 }, { deckAtStart: 3 }]) {
      const out = undoVerdict({ foundAtStart: false, ...half });
      expect(out.detail, `one count alone: ${JSON.stringify(half)}`).toContain("NOT ASKED");
      expect(out.detail, "and it never subtracts against a missing count").not.toContain("NaN");
    }
  });

  it("answers a lone second sheet from the marker, and prefers the previous sheet when it has it", () => {
    // The marker holds the deck size BEFORE the slide was left; the previous
    // sheet holds the size after. One more than the marker is the same
    // number, so a second sheet on its own can say what happened.
    const alone = undoVerdict({ foundAtStart: false, deckAtStart: 3, previousDeckBeforeLeave: 3 });
    expect(alone).toMatchObject({ verdict: "yes" });
    expect(alone.detail).toContain("marker");
    expect(undoVerdict({ foundAtStart: false, deckAtStart: 4, previousDeckBeforeLeave: 3 }).detail).toContain(
      "something else happened",
    );
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 4, previousDeckBeforeLeave: 3 })).toMatchObject({
      verdict: "no",
    });
    expect(undoVerdict({ foundAtStart: "unsupported", deckAtStart: 3, previousDeckBeforeLeave: 3 })).toMatchObject({
      verdict: "yes",
    });
    // Both given: the sheet is the measurement, the marker the note; the sheet wins.
    const both = undoVerdict({ foundAtStart: false, deckAtStart: 3, previousDeckAtEnd: 4, previousDeckBeforeLeave: 9 });
    expect(both).toMatchObject({ verdict: "yes" });
    expect(both.detail).toContain("the previous sheet");
  });
});

describe("floorLine", () => {
  it("says what the host has against the floor and the reads", () => {
    const full = floorLine([...API_SETS], "OfficeOnline", "1.2");
    expect(full).toContain("up to 1.10: clears the floor 1.2");
    expect(full).toContain("exportAsBase64Presentation (1.10)");
    const bare = floorLine(["1.1"], undefined, "1.2");
    expect(bare).toContain("unknown platform");
    expect(bare).toContain("BELOW the floor");
    expect(bare).toContain("no getSelectedSlides");
    expect(floorLine([], "iOS", "1.2")).toContain("up to none");
  });
});

describe("leftBehind", () => {
  const parts = (over: Partial<PartsSummary> = {}): PartsSummary => ({
    total: 40,
    slides: 3,
    masters: 1,
    layouts: 11,
    themes: 1,
    comments: 0,
    authors: false,
    media: 2,
    ...over,
  });

  it("reports a missing summary, an unchanged package and what changed", () => {
    expect(leftBehind(undefined, parts())).toContain("not measured");
    expect(leftBehind(parts(), parts())).toContain("ended as it started");
    const changed = leftBehind(parts(), parts({ masters: 2, themes: 2, authors: true, total: 44 }));
    expect(changed).toContain("masters 1 → 2");
    expect(changed).toContain("themes 1 → 2");
    expect(changed).toContain("authors part gained");
  });
});

describe("jumpProbeVerdict", () => {
  const asked = { supported: true, wanted: "256#1", ms: 480, afterwards: { answered: true, ms: 12 } };

  it("cannot say without the call or without the arm, and names a throw", () => {
    expect(jumpProbeVerdict({ supported: false }).detail).toContain("1.5");
    expect(jumpProbeVerdict({ supported: true }).detail).toContain("NOT ASKED");
    expect(jumpProbeVerdict({ supported: true, error: "boom" })).toMatchObject({ verdict: "threw" });
  });

  it("says yes only when the view moved AND the host still answered afterwards", () => {
    expect(jumpProbeVerdict({ ...asked, selected: ["256#1"] })).toMatchObject({ verdict: "yes" });
    expect(jumpProbeVerdict({ ...asked, selected: ["256"] })).toMatchObject({ verdict: "yes" });
    const wedged = jumpProbeVerdict({
      ...asked,
      selected: ["256#1"],
      afterwards: { answered: false, error: "gave up" },
    });
    expect(wedged.verdict).toBe("no");
    expect(wedged.detail).toContain("NEXT selection read");
  });

  it("says no when the view did not move or the read-back went silent", () => {
    expect(jumpProbeVerdict({ ...asked, selected: ["257#2"] }).detail).toContain("did not move");
    const silent = jumpProbeVerdict({ ...asked, selected: null, afterwards: { answered: false } });
    expect(silent.verdict).toBe("no");
    expect(silent.detail).toContain("wedge");
  });

  it("says NOT ASKED when the sheet carries only half the jump arm", () => {
    // The id asked for and the id answered are one question; either alone is
    // no reading at all.
    expect(jumpProbeVerdict(asked).detail).toContain("NOT ASKED");
    expect(jumpProbeVerdict({ supported: true, selected: ["256#1"] }).detail).toContain("NOT ASKED");
  });
});
