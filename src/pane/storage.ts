/**
 * What the pane remembers and where, as rules rather than as storage.
 *
 * `src/host/memory.ts` decides WHICH bucket a deck remembers itself in — the
 * key. This decides what goes in it, what comes back out, and what a value that
 * is not what it should be counts as.
 *
 * It lives here rather than in `main.ts`, where it was written, because
 * `main.ts` is excluded from the coverage floor (`scripts/coverage-scope.mjs`)
 * on the grounds that it is Office.js calls and DOM plumbing — untestable by
 * construction. These are neither. They are ordinary decisions over plain
 * values, and every one of them was going unmeasured inside a file exempted for
 * being something else.
 *
 * The reading and the writing stay in `main.ts`: `localStorage` throws on the
 * accessor itself in a WebView with site data blocked, and a `try` around a
 * browser API is exactly the kind of thing this file is not for.
 */
import { GLOBAL_KEY } from "../host/memory.js";
import { DEFAULT_SETTINGS, type PaneState } from "./steps.js";

/**
 * One bucket as it comes back out of storage.
 *
 * `Partial<PaneState>` is what it is MEANT to be. It is whatever JSON was
 * there, so every field is read defensively below — a hand-edited value, or
 * one written by an older build, must not take the pane down on open.
 */
export type Stored = Partial<PaneState> & { scroll?: unknown };

/**
 * How far down the list the deck's bucket says the user had scrolled.
 *
 * Zero for anything that is not a positive finite number, which is the same
 * answer as "never scrolled" — so a corrupt value costs the restore rather than
 * scrolling somebody to `NaN`.
 */
export function storedScroll(deck: Stored): number {
  const scroll = deck.scroll;
  return typeof scroll === "number" && Number.isFinite(scroll) && scroll > 0 ? scroll : 0;
}

/**
 * The state a pair of buckets restores to (`docs/DESIGN.md` section 4).
 *
 * Favourites and the first-run flag come from the MACHINE's bucket — a star is
 * a statement about the library and the coach marks are "dismissed once". The
 * rest is per deck, because the way you were reading the library for one deck
 * is rarely the way you want to read it for the next.
 *
 * Recent is per deck too, which the record does not settle either way. Decided
 * in the build: Recent exists so the thing just used is easy to reach again,
 * and "just used" is a fact about a deck. Favourites is the per-machine half of
 * that pair by the record's own wording.
 *
 * When the host would not name a deck, the caller passes the same object twice:
 * one bucket holds both halves, and that is the fallback working rather than a
 * case to special-case here.
 */
export function restored(machine: Stored, deck: Stored): Partial<PaneState> {
  return {
    favourites: machine.favourites ?? [],
    coached: machine.coached === true,
    settings: { ...DEFAULT_SETTINGS, ...(deck.settings ?? {}) },
    recent: deck.recent ?? [],
    open: deck.open ?? [],
    // Section 4 asks for the search and the tags back too, and for the size
    // picked in a stepper — which is `chosen`. It is written only by the
    // operations that already persist, so what comes back is the step last
    // INSERTED rather than wherever the keyboard was left: the pane marks both
    // with the same field, and a focus cursor is not a pick.
    query: typeof deck.query === "string" ? deck.query : "",
    tags: deck.tags ?? [],
    ...(typeof deck.chosen === "string" ? { chosen: deck.chosen } : {}),
  };
}

/**
 * Everything to write, as the pairs a caller hands to storage in order.
 *
 * A LIST rather than two calls, so the one-bucket case is a rule here instead
 * of a branch at the call site: when there is no deck to tell apart, the two
 * halves are merged into a single write, because writing them separately under
 * the same key leaves the second overwriting the first a moment later.
 */
export function writes(bucket: string, state: PaneState, scroll: number): [string, unknown][] {
  const machine = { favourites: state.favourites, coached: state.coached === true };
  const deck = {
    settings: state.settings,
    recent: state.recent,
    open: state.open,
    query: state.query,
    tags: state.tags,
    ...(state.chosen === undefined ? {} : { chosen: state.chosen }),
    ...(scroll > 0 ? { scroll } : {}),
  };
  if (bucket === GLOBAL_KEY) return [[GLOBAL_KEY, { ...machine, ...deck }]];
  return [
    [GLOBAL_KEY, machine],
    [bucket, deck],
  ];
}

/**
 * Whether the scroll offset should be put back on this draw.
 *
 * Once, and only once there are tiles to put it back into: the loading screen
 * is one short paragraph, and scrolling that to 800 px leaves the user looking
 * at nothing. A later draw is the user's own doing, and re-scrolling them to
 * where they were an hour ago is the pane fighting them — so is putting the old
 * offset back after they have already scrolled themselves, which is why the
 * caller sets `restored` on a scroll of their own as well as on a restore.
 */
export function shouldRestoreScroll(at: { restored: boolean; kept: number; hasTiles: boolean }): boolean {
  return !at.restored && at.kept > 0 && at.hasTiles;
}
