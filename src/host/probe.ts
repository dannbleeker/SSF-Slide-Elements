/**
 * What the host probe's observations mean, decided away from the host.
 *
 * The snippet that runs inside PowerPoint (`probe/probe-snippet.ts`, generated
 * by `scripts/build-probe.mjs`) collects raw observations and makes no
 * judgements: counts, ids, milliseconds, part names and the errors it caught.
 * Every reading happens here, as a pure function over those values, so the
 * rules can go red in CI while the round trip through a real PowerPoint stays
 * a matter of reading values out. A probe that reasons inside a
 * `PowerPoint.run` callback is a probe whose conclusions nobody can check.
 *
 * The wording is part of the contract. SSF-Merge's rounds produced verdicts on
 * questions the run had not asked, so every function here either names a
 * mechanism or says plainly that it CANNOT TELL, and a question the run could
 * not put is `unknown`, never `no`.
 *
 * The questions are `docs/DESIGN.md` section 13; `docs/PROBE.md` says how the
 * probe asks each one.
 */

export type Verdict = "yes" | "no" | "unknown" | "threw";

export interface Reading {
  verdict: Verdict;
  detail: string;
}

/** The tag key the probe writes into the slide it leaves behind for the Ctrl+Z question. */
export const PROBE_TAG = "SSF_SLIDE_ELEMENTS_PROBE";
/** Its value on the slide left behind. */
export const PROBE_UNDO_VALUE = "undo";
/**
 * The document-settings key the probe writes BEFORE it leaves that slide. The
 * settings live outside the undo stack, so the marker survives the Ctrl+Z the
 * slide does not, and the next run can tell it is the second of a pair even
 * when the slide is gone. Without it the second run looked like a first and
 * left a slide of its own (the web round of 2026-09-10).
 */
export const PROBE_MARKER = "SSF_SLIDE_ELEMENTS_PROBE_UNDO";

// ---------------------------------------------------------------------------
// Inserts: the delta is the evidence, never the absence of an error.
// ---------------------------------------------------------------------------

export interface InsertObservation {
  /** Deck size before the call, measured. */
  before: number;
  /** Deck size after the call, measured in its own context. */
  after: number;
  /** How many slides the package being inserted LISTS. */
  expected: number;
  /** The error, if the call threw. */
  error?: string;
}

export interface InsertVerdict extends Reading {
  landed: number;
}

/**
 * Whether an insert did what it said.
 *
 * Ported from SSF-Merge, where three real sheets shaped it: a call can raise
 * and still have done the work (an insert timed out with both slides landed),
 * more can arrive than the package listed, and the deck can shrink. Each is
 * named rather than folded into "partial".
 */
export function insertVerdict(o: InsertObservation): InsertVerdict {
  const landed = o.after - o.before;
  if (o.error !== undefined) {
    if (landed === o.expected) {
      return {
        verdict: "yes",
        landed,
        detail: `all ${o.expected} slide(s) landed, but not before the probe stopped waiting (${o.error}). That is the probe's budget being short, not the host refusing.`,
      };
    }
    return {
      verdict: "threw",
      landed,
      detail: `the call threw: ${o.error}${landed === 0 ? "" : `, and ${landed} slide(s) landed anyway`}`,
    };
  }
  if (landed === o.expected) return { verdict: "yes", landed, detail: `all ${o.expected} slide(s) landed` };
  if (landed > o.expected) {
    return {
      verdict: "unknown",
      landed,
      detail: `the deck grew by ${landed} while the package listed ${o.expected} slide(s), so the host inserted more than the slide list names`,
    };
  }
  if (landed < 0) {
    return {
      verdict: "unknown",
      landed,
      detail: `the deck SHRANK by ${-landed} slide(s) across an insert of ${o.expected}, so something else changed it`,
    };
  }
  if (landed === 0) {
    return {
      verdict: "no",
      landed,
      detail: `the call raised nothing and the deck did not grow, so ${o.expected} slide(s) were dropped silently`,
    };
  }
  return {
    verdict: "no",
    landed,
    detail: `${landed} of ${o.expected} slide(s) landed, which is a partial insert rather than a refusal`,
  };
}

/** An arm the sheet does not carry did not RUN, and says so rather than grading nothing. */
export function notAsked(what: string): InsertVerdict {
  return {
    verdict: "unknown",
    landed: 0,
    detail: `NOT ASKED — this sheet carries no ${what}. Evidence about nothing.`,
  };
}

/**
 * Whose fault a refused insert is.
 *
 * The control arm inserts the presentation's OWN bytes, read back through
 * `getFileAsync`. That deck is a package PowerPoint wrote seconds earlier, so
 * it cannot be malformed, and a host that refuses it is refusing insertion
 * itself. Do not reason about which of two readings holds; ask a question only
 * one of them survives.
 */
export function insertionBlame(ours: Verdict, self: Verdict): string {
  if (ours === "yes") return "The insert path works. Whose package it is does not arise.";
  if (self === "yes") {
    return "OURS: this host inserted a deck PowerPoint itself wrote and refused the one this repo generates. The package writer is at fault; nothing in the failing arms is a fact about the host.";
  }
  if (self === "unknown") {
    return "CANNOT TELL: the control arm never ran, so a refused insert is equally our package or this host. Re-run before concluding anything.";
  }
  return "THE HOST: it refused a deck it wrote itself, so nothing can be inserted here and the package path is blocked for reasons no change in this repo can reach.";
}

/**
 * Question 1: which pruning of a package the host accepts on the way in.
 *
 * The engine will read the user's whole deck, splice an element into one
 * slide and send ONE slide back. The cheap way to send one slide is to unlist
 * the others and leave their parts in the zip; the expensive way is
 * `Pkg.removeSlide` over every other slide, orphans walked. Three arms tell
 * them apart:
 *
 * - `listed`:   two slides, both listed — the package as PowerPoint would write it;
 * - `pruned`:   the second slide's `<p:sldId>` AND its relationship removed, part and content type kept;
 * - `unlisted`: only the `<p:sldId>` removed, the relationship still pointing at the part.
 *
 * Each pruned arm lists ONE slide, so the expected landing is one. Two landing
 * means the host walked the relationships rather than the slide list, which
 * is a different fact and is named.
 */
export function pruningReading(listed: InsertVerdict, pruned: InsertVerdict, unlisted: InsertVerdict): string {
  if (listed.verdict !== "yes") {
    return "NOT ANSWERED: the fully listed package did not land, so nothing about pruning can be read from the pruned arms. Read the errors and the control arm first.";
  }
  const cheap = pruned.verdict === "yes";
  const cheaper = unlisted.verdict === "yes";
  if (cheap && cheaper) {
    return "Both prunings land as one slide: the host reads the slide list and ignores an unlisted part, related or not. The engine can send one slide by unlisting the rest and leaving their parts in place.";
  }
  if (cheap) {
    return `Removing the slide id AND its relationship lands as one slide; leaving the relationship does not (${unlisted.detail}). The engine must drop the relationship as well as the id, and can leave the part.`;
  }
  if (cheaper) {
    return `Unexpected: leaving the relationship lands and removing it does not (${pruned.detail}). Do not build on either reading; re-run before drawing a conclusion.`;
  }
  return `Neither pruning lands (pruned: ${pruned.detail}; unlisted: ${unlisted.detail}). The host wants a package whose parts match its slide list, so the engine must remove every other slide properly — Pkg.removeSlide, orphans walked.`;
}

// ---------------------------------------------------------------------------
// Masters: does inserting a deck's own slide back into it add a master?
// ---------------------------------------------------------------------------

export interface MasterObservation {
  mastersBefore?: number;
  mastersAfter?: number;
  /** How many slides the insert landed; a refused insert cannot add a master. */
  landed: number;
  /** Which formatting option the insert used. */
  formatting: string;
}

/**
 * Whether an insert of the deck's own package adds a slide master.
 *
 * Every insert this add-in makes is a copy of one of the user's own slides,
 * on the user's own master. If `insertSlidesFromBase64` under
 * `KeepSourceFormatting` brings that master in AGAIN, every click grows the
 * deck by a master and a theme, and a deck edited for an hour carries dozens.
 * `UseDestinationTheme` is the alternative if so; the pair tells them apart.
 */
export function masterVerdict(o: MasterObservation): Reading {
  if (o.mastersBefore === undefined || o.mastersAfter === undefined) {
    return {
      verdict: "unknown",
      detail: `NOT ASKED — the host would not count its masters (slideMasters is PowerPointApi 1.3), so whether ${o.formatting} adds one is unmeasured here.`,
    };
  }
  if (o.landed <= 0) {
    return {
      verdict: "unknown",
      detail: `NOT ASKED — the ${o.formatting} insert landed nothing, so it could not have added a master either way.`,
    };
  }
  const added = o.mastersAfter - o.mastersBefore;
  if (added > 0) {
    return {
      verdict: "yes",
      detail: `${o.formatting} ADDED ${added} master(s) (${o.mastersBefore} → ${o.mastersAfter}) for slides on a master the deck already had. Every insert would grow the deck by a master.`,
    };
  }
  if (added < 0) {
    return {
      verdict: "unknown",
      detail: `the master count FELL (${o.mastersBefore} → ${o.mastersAfter}) across a ${o.formatting} insert, which nothing here predicts. Record it and re-run.`,
    };
  }
  return {
    verdict: "no",
    detail: `${o.formatting} added no master (${o.mastersBefore} before and after), so a slide on the deck's own master comes back onto it.`,
  };
}

// ---------------------------------------------------------------------------
// Question 2: insert after the current slide, then delete the current slide.
// ---------------------------------------------------------------------------

export interface OrderObservation {
  /** The first probe slide's index, which is the deck's size when the arm began. */
  n?: number;
  /** Positional ids at n and n+1 after inserting the two-slide deck: [P1, P2]. */
  afterFirst?: string[];
  /** Whether the single-slide insert could target P1 (a slide this run had just added) by id. */
  targetAdded?: { ok: boolean; error?: string };
  /** Positional ids at n, n+1, n+2 after the single-slide insert. */
  afterSecond?: string[];
  /** Which index the probe then deleted (P1's, wherever the second insert put it). */
  deletedIndex?: number;
  /** Positional ids at n and n+1 after the delete. */
  afterDelete?: string[];
  error?: string;
}

/**
 * Whether insert-then-positional-delete keeps the order the engine expects.
 *
 * The engine's "onto this slide" is: insert the rebuilt slide AFTER the
 * current one (`targetSlideId`), prove the deck grew, then delete the current
 * one by POSITION, so the copy takes its place. This arm does exactly that
 * with slides of its own: P1 stands in for the current slide, S for the
 * rebuilt copy, P2 for the slide after.
 *
 * Two facts, read separately because they fail separately: whether S landed
 * right after its target (and P2 moved down one), and whether deleting P1 by
 * index left `[S, P2]` in that order.
 */
export function orderVerdict(o: OrderObservation): Reading {
  if (o.error !== undefined) return { verdict: "threw", detail: `the arm threw: ${o.error}` };
  const first = o.afterFirst ?? [];
  const second = o.afterSecond ?? [];
  const after = o.afterDelete ?? [];
  if (first.length !== 2 || second.length !== 3 || after.length !== 2) {
    return {
      verdict: "unknown",
      detail: `NOT ASKED — the arm did not read all three positions (${first.length}, ${second.length}, ${after.length} ids), so nothing about order can be read from it.`,
    };
  }
  const [p1, p2] = first as [string, string];
  const sIndex = second.findIndex((id) => id !== p1 && id !== p2);
  if (sIndex < 0) {
    return {
      verdict: "no",
      detail: "the single-slide insert added no new id among the three positions, so S never landed",
    };
  }
  const s = second[sIndex] as string;
  // Where S should be: right after P1 when P1 was the target, else first
  // (the target was the slide before P1, the user's own last slide).
  const targetedP1 = o.targetAdded?.ok === true;
  const wantSecond = targetedP1 ? [p1, s, p2] : [s, p1, p2];
  const placed = second.every((id, i) => id === wantSecond[i]);
  const kept = after[0] === s && after[1] === p2;
  const target = targetedP1 ? "the slide this run had just added" : "the user's last slide";
  if (placed && kept) {
    return {
      verdict: "yes",
      detail: `S landed right after ${target} and the slides behind it moved down one; deleting P1 at index ${o.deletedIndex ?? "?"} left [S, P2] in order. Insert-then-positional-delete is safe here.`,
    };
  }
  if (!placed) {
    return {
      verdict: "no",
      detail: `S landed at position ${sIndex} of the three (wanted ${wantSecond.indexOf(s)}), so targetSlideId did not place it right after ${target}. Read: ${second.join(", ")}.`,
    };
  }
  return {
    verdict: "no",
    detail: `S landed where expected but the positional delete left [${after.join(", ")}] rather than [${s}, ${p2}], so the delete took the wrong slide or reordered the rest.`,
  };
}

/** Whether a slide this run just added can be a `targetSlideId`, which "as a new slide" twice in a row needs. */
export function targetAddedVerdict(o: OrderObservation): Reading {
  if (o.error !== undefined && o.targetAdded === undefined) {
    return { verdict: "threw", detail: `the arm threw before it could try: ${o.error}` };
  }
  if (o.targetAdded === undefined)
    return { verdict: "unknown", detail: "NOT ASKED — this sheet carries no targetAdded reading." };
  if (o.targetAdded.ok) {
    return {
      verdict: "yes",
      detail:
        "a slide this run had just added was accepted as targetSlideId, so a second insert can land after the first.",
    };
  }
  return {
    verdict: "no",
    detail: `a slide this run had just added was REFUSED as targetSlideId (${o.targetAdded.error ?? "no error text"}). A second insert must target a slide the host already knew, or re-read ids after the first.`,
  };
}

// ---------------------------------------------------------------------------
// Question 3: does getSelectedSlides name the slide the user is on?
// ---------------------------------------------------------------------------

export interface SelectionObservation {
  /** False when the host has no getSelectedSlides (PowerPointApi 1.5). */
  supported?: boolean;
  /** What getSelectedSlides answered, ids in its order. */
  selectedIds?: string[];
  /** Each selected id's index in the positional read, -1 when not found. */
  selectedIndexes?: number[];
  /** Deck size by getCount. */
  deckSize?: number;
  /** How many positional ids were read (capped by the probe). */
  positionalRead?: number;
  /** Whether the positional ids' first halves match `<p:sldIdLst>` in the file, entry for entry. */
  sldIdLstMatches?: boolean;
  /** How many entries `<p:sldIdLst>` had, or undefined when the file could not be read. */
  sldIdLstEntries?: number;
  error?: string;
}

/**
 * Whether the selection read names the slide the user is looking at, and
 * whether its API position is its file position.
 *
 * The probe cannot know what the user selected; the instructions ask for
 * slide 2 so a person can compare. The reading here is what CAN be decided
 * from the numbers: that something was selected, that every selected id is at
 * a position in the deck, and that the API's order is the file's order — which
 * is what lets the engine turn a selected id into a `<p:sldId>` to splice.
 */
export function selectionVerdict(o: SelectionObservation): Reading {
  if (o.supported === false) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — this host has no getSelectedSlides (PowerPointApi 1.5), so the pane cannot know the current slide here and must ask the user to say which.",
    };
  }
  if (o.error !== undefined) return { verdict: "threw", detail: `the read threw: ${o.error}` };
  const ids = o.selectedIds ?? [];
  if (ids.length === 0) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — getSelectedSlides answered nothing. Either no slide was selected when the probe ran or the host answers empty; select a slide in the strip and re-run.",
    };
  }
  const idx = o.selectedIndexes ?? [];
  const lost = idx.filter((i) => i < 0).length;
  if (lost > 0) {
    return {
      verdict: "no",
      detail: `${ids.length} slide(s) selected and ${lost} of them are at no position among the ${o.positionalRead ?? 0} read by getItemAt, so a selected id cannot be turned into a slide number here.`,
    };
  }
  const positions = idx.map((i) => i + 1).join(", ");
  const file =
    o.sldIdLstMatches === true
      ? `and the API's order is the file's <p:sldIdLst> order (${o.sldIdLstEntries ?? "?"} entries)`
      : o.sldIdLstMatches === false
        ? `BUT the API's order is NOT the file's <p:sldIdLst> order, so a position from the API cannot index the file's list`
        : "and the file's order could not be compared (no sldIdLst read)";
  return {
    verdict: o.sldIdLstMatches === false ? "no" : "yes",
    detail: `${ids.length} slide(s) selected, at slide number(s) ${positions} of ${o.deckSize ?? "?"}, ${file}. Check that the number is the slide you had selected.`,
  };
}

// ---------------------------------------------------------------------------
// Question 4: which read of the deck, and what each drops.
// ---------------------------------------------------------------------------

export interface PartsSummary {
  total: number;
  slides: number;
  masters: number;
  layouts: number;
  themes: number;
  comments: number;
  authors: boolean;
  media: number;
}

/**
 * A part list boiled down to what the reader wants to compare.
 *
 * NAMES only ever reach the sheet, never content: a part list is structure,
 * and the sheet is written to be pasted into an issue.
 */
export function summarizeParts(names: string[]): PartsSummary {
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length;
  return {
    total: names.length,
    slides: count(/^ppt\/slides\/slide\d+\.xml$/),
    masters: count(/^ppt\/slideMasters\/slideMaster\d+\.xml$/),
    layouts: count(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/),
    themes: count(/^ppt\/theme\/theme\d+\.xml$/),
    comments: count(/^ppt\/(comments|modernComments)\//),
    authors: names.includes("ppt/authors.xml"),
    media: count(/^ppt\/media\//),
  };
}

export interface ExportPartsObservation {
  /** False when the host has no exportAsBase64Presentation (PowerPointApi 1.10). */
  supported?: boolean;
  source?: PartsSummary;
  exported?: PartsSummary;
  /** In the file-route package and not in the export, capped by the probe. */
  missing?: string[];
  error?: string;
}

/**
 * Whether the export drops parts the file route keeps.
 *
 * office-js#6867 reports the slide-level export omitting modern comments and
 * `ppt/authors.xml`; SSF-Merge's sixth sheet found the presentation-level
 * export does the same on the web. This add-in has not chosen its read yet:
 * `getFileAsync` (the floor) hands back the whole deck, the export (1.10)
 * only the slides asked for. What each drops on THIS host decides it.
 *
 * A deck with no comments cannot answer, and says so.
 */
export function exportPartsVerdict(o: ExportPartsObservation): Reading {
  if (o.supported === false) {
    return {
      verdict: "unknown",
      detail:
        "NOT ASKED — this host has no exportAsBase64Presentation, so getFileAsync is the only read here and the question does not arise.",
    };
  }
  if (o.error !== undefined) return { verdict: "threw", detail: `the comparison threw: ${o.error}` };
  if (!o.source || !o.exported) return { verdict: "unknown", detail: "NOT ASKED — this sheet carries no part lists." };
  const s = o.source;
  const e = o.exported;
  if (!s.authors && s.comments === 0) {
    return {
      verdict: "unknown",
      detail: `NOT ASKED — this deck carries no comments and no authors part, so there was nothing for the export to drop. Re-run on a deck with comments. (${s.total} parts in, ${e.total} out.)`,
    };
  }
  const lostAuthors = s.authors && !e.authors;
  const lostComments = e.comments < s.comments;
  if (lostAuthors || lostComments) {
    const lost = [
      lostAuthors ? "ppt/authors.xml" : "",
      lostComments ? `${s.comments - e.comments} comment part(s)` : "",
    ]
      .filter(Boolean)
      .join(" and ");
    return {
      verdict: "yes",
      detail: `the export DROPS ${lost} — office-js#6867 reaches the presentation-level call here too. A package rebuilt from the export loses them; getFileAsync keeps them. ${s.total} parts in, ${e.total} out.`,
    };
  }
  return {
    verdict: "no",
    detail: `the export kept the comments and the authors part this deck carries (${s.total} parts in, ${e.total} out${(o.missing ?? []).length > 0 ? `, ${(o.missing ?? []).length} other part(s) not carried over` : ""}).`,
  };
}

export interface ReadTiming {
  ms?: number;
  bytes?: number;
  error?: string;
}

/** One line per read: how long, how big, at what rate. Measured, and dated by the sheet. */
export function timingLine(what: string, t: ReadTiming | undefined): string {
  if (!t) return `${what}: not asked`;
  if (t.error !== undefined) return `${what}: threw — ${t.error}`;
  if (t.ms === undefined || t.bytes === undefined) return `${what}: not measured`;
  const mb = t.bytes / 1048576;
  const rate = t.ms > 0 ? ` (${(mb / (t.ms / 1000)).toFixed(1)} MB/s)` : "";
  return `${what}: ${mb.toFixed(2)} MB in ${t.ms} ms${rate}`;
}

// ---------------------------------------------------------------------------
// Question 5: does PowerPoint's own Ctrl+Z revert an insert?
// ---------------------------------------------------------------------------

export interface UndoObservation {
  /** Whether the probe found its tagged slide from a previous run when it started. "unsupported" when tags cannot be read (1.3). */
  foundAtStart?: boolean | "unsupported";
  /** Whether this run left a tagged slide behind at the end, for the next run to look for. */
  leftBehind?: boolean;
  /** The deck size this run started at. */
  deckAtStart?: number;
  /** The deck size the PREVIOUS sheet ended at, when the reader was given two sheets. */
  previousDeckAtEnd?: number;
  /**
   * The deck size the previous run wrote into its marker right BEFORE it left
   * the slide, when this run found the marker. One more than this is what
   * that run left, so a lone second sheet can answer without the first.
   */
  previousDeckBeforeLeave?: number;
  error?: string;
}

/**
 * Whether PowerPoint's own undo reverts `insertSlidesFromBase64`.
 *
 * Two runs, because the probe cannot press Ctrl+Z. The first writes a marker
 * into the document's settings, leaves one tagged slide at the end of the deck
 * and says so; the user presses Ctrl+Z once on the slide canvas and runs the
 * snippet again. The second run looks for the tag. Gone, with the deck one
 * smaller than the first run left it: Ctrl+Z reverts the insert and the
 * pane's Undo must not fight it. Still there: it does not, and the pane's
 * Undo is the only way back.
 *
 * "One smaller than the first run left it" comes from the previous sheet when
 * the reader has it, else from the marker, which the first run wrote before
 * the slide and Ctrl+Z cannot reach. A first run cannot answer this and says
 * so; the second cannot tell a Ctrl+Z from a hand delete, which is why the
 * instructions say not to.
 */
export function undoVerdict(o: UndoObservation): Reading {
  if (o.error !== undefined) return { verdict: "threw", detail: `the tag read threw: ${o.error}` };
  const previousEnd =
    o.previousDeckAtEnd ?? (o.previousDeckBeforeLeave === undefined ? undefined : o.previousDeckBeforeLeave + 1);
  const source =
    o.previousDeckAtEnd !== undefined
      ? "the previous sheet"
      : `the marker the previous run wrote before it left the slide (${o.previousDeckBeforeLeave ?? "?"} slides, so ${previousEnd ?? "?"} after it)`;
  if (o.foundAtStart === "unsupported") {
    if (previousEnd === undefined || o.deckAtStart === undefined) {
      return {
        verdict: "unknown",
        detail:
          "NOT ASKED — this host cannot read tags (PowerPointApi 1.3), and neither the previous sheet nor a marker from the previous run was there, so the deck counts cannot be compared either.",
      };
    }
    const delta = previousEnd - o.deckAtStart;
    if (delta === 1)
      return {
        verdict: "yes",
        detail: `by count alone, against ${source}: the deck is one slide smaller than the previous run left it, which is the slide Ctrl+Z took back. Tags could not be read to confirm it was ours.`,
      };
    if (delta === 0)
      return {
        verdict: "no",
        detail: `by count alone, against ${source}: the deck is the size the previous run left it, so Ctrl+Z took nothing back. Tags could not be read to confirm the slide is ours.`,
      };
    return {
      verdict: "unknown",
      detail: `the deck changed by ${-delta} slide(s) since the previous run (against ${source}), which is not the one slide this question is about.`,
    };
  }
  if (o.foundAtStart === true) {
    return {
      verdict: "no",
      detail:
        "the tagged slide from the previous run was still there, so PowerPoint's Ctrl+Z did NOT revert the insert (or was not pressed). The pane's Undo is the only way back; this run removed the slide.",
    };
  }
  if (o.foundAtStart === false && previousEnd !== undefined && o.deckAtStart !== undefined) {
    const delta = previousEnd - o.deckAtStart;
    if (delta === 1) {
      return {
        verdict: "yes",
        detail: `the tagged slide from the previous run is gone and the deck is exactly one smaller than that run left it, by ${source}: PowerPoint's Ctrl+Z reverts an insert, so the pane's Undo must not fight it.`,
      };
    }
    return {
      verdict: "unknown",
      detail: `the tagged slide is gone but the deck changed by ${-delta} slide(s) rather than one since the previous run (against ${source}), so something else happened between the runs.`,
    };
  }
  if (o.leftBehind === true) {
    return {
      verdict: "unknown",
      detail:
        "NOT YET — this is the first run: one tagged slide was left at the end of the deck. Press Ctrl+Z once on the slide canvas, run the snippet again, and read both sheets together.",
    };
  }
  return {
    verdict: "unknown",
    detail:
      "NOT ASKED — no tagged slide was found at the start and none was left behind, so this sheet says nothing about Ctrl+Z on its own. Read it together with the sheet before it.",
  };
}

// ---------------------------------------------------------------------------
// The floor, and what the sweep left behind.
// ---------------------------------------------------------------------------

/** The PowerPointApi sets a host can report, lowest first. */
export const API_SETS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"] as const;

/** Question 6's second half: whether this host clears the floor, and which reads it has. */
export function floorLine(sets: string[], platform: string | undefined, floor: string): string {
  const has = (v: string) => sets.includes(v);
  const highest = [...API_SETS].reverse().find(has) ?? "none";
  const reads = [
    "getFileAsync (Common API, every host)",
    has("1.10") ? "exportAsBase64Presentation (1.10)" : "no exportAsBase64Presentation (needs 1.10)",
  ].join("; ");
  const extras = [
    has("1.3") ? "tags and masters readable (1.3)" : "no tags, no master count (needs 1.3)",
    has("1.5") ? "getSelectedSlides (1.5)" : "no getSelectedSlides (needs 1.5)",
  ].join("; ");
  const ok = has(floor) ? `clears the floor ${floor}` : `is BELOW the floor ${floor}`;
  return `${platform ?? "unknown platform"} reports PowerPointApi up to ${highest}: ${ok}; ${reads}; ${extras}.`;
}

/** What the deck looks like after the sweep, against how it started — slides first, then the parts the sweep cannot reach. */
export function leftBehind(start: PartsSummary | undefined, end: PartsSummary | undefined): string {
  if (!start || !end) return "not measured — the sheet carries no part summary for both ends";
  const lines: string[] = [];
  for (const key of ["slides", "masters", "layouts", "themes", "media", "comments"] as const) {
    if (end[key] !== start[key]) lines.push(`${key} ${start[key]} → ${end[key]}`);
  }
  if (end.authors !== start.authors) lines.push(`authors part ${start.authors ? "lost" : "gained"}`);
  if (lines.length === 0) return `the package ended as it started (${end.total} parts)`;
  return `the package changed: ${lines.join(", ")} (${start.total} → ${end.total} parts). A sweep removes slides; what it cannot remove is listed here.`;
}
