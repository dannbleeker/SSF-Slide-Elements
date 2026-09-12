/**
 * Which elements the picker shows, and what it offers when a search finds none.
 *
 * Split out of `steps.ts` on 2026-09-12. It is one question asked several ways —
 * does this element match, which category is it under, is that category open,
 * how many tiles are left, and what did you probably mean — and in one file with
 * `steps.ts` the halves of it sat 640 lines apart: `matches` near the top and
 * `didYouMean`, which is the same feature's last resort, near the bottom.
 *
 * Everything here is a pure function over plain values, like the rest of the
 * pane's decisions, so the suite can hold it to an answer with no browser.
 */
import type { Element } from "../core/catalogue/types.js";
import type { Library, PaneState } from "./steps.js";

/**
 * Whether a query matches an element.
 *
 * `docs/DESIGN.md` section 8: the English name, the Danish key, the category
 * and the tags. The key is searched even though the pane never shows it,
 * because the owner's decks are Danish and somebody who knows the library by
 * its Danish names should be able to find things.
 *
 * Every WORD has to match, not the whole string as a substring: "white box"
 * should find "White boxes, 2x1 vertical", where a substring match on the
 * phrase finds nothing at all.
 */
export function matches(element: Element, query: string): boolean {
  const q = query.trim().toLowerCase();
  if (q === "") return true;
  const hay = [element.name, element.key, element.category.name, element.category.key, ...element.tags]
    .join(" ")
    .toLowerCase();
  return q.split(/\s+/).every((word) => hay.includes(word));
}

/** Every tag in the library, most used first, for the tag line. */
export function tagsOf(library: Library): string[] {
  const counts = new Map<string, number>();
  for (const el of library.elements) for (const tag of el.tags) counts.set(tag, (counts.get(tag) ?? 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0])).map(([tag]) => tag);
}

/** One category with the elements a search left in it. */
export interface Group {
  key: string;
  name: string;
  elements: Element[];
}

/**
 * The elements to show, grouped by category, in the deck's own order.
 *
 * A sized element is ONE tile with a stepper (`docs/DESIGN.md` section 2), so a
 * run of "process flow with 1 box" through "with 6 boxes" collapses to its
 * first surviving member and the rest are reachable through the stepper. The
 * 16:9 deck has twelve such runs over 45 elements, so 117 elements show as 84
 * tiles: 117 − 45 + 12. Counted by running this function over the committed
 * catalogue, and confirmed against the count the pane itself draws.
 */
export function groups(library: Library, state: PaneState): Group[] {
  const wanted = library.elements.filter(
    (el) =>
      matches(el, state.query) &&
      state.tags.every((tag) => el.tags.includes(tag)) &&
      (state.category === undefined || el.category.key === state.category),
  );
  const seen = new Set<string>();
  const tiles = wanted.filter((el) => {
    if (!el.run) return true;
    if (seen.has(el.run.key)) return false;
    seen.add(el.run.key);
    return true;
  });
  const out: Group[] = [];
  for (const category of library.categories) {
    const elements = tiles.filter((el) => el.category.key === category.key);
    if (elements.length === 0) continue;
    out.push({ key: category.key, name: category.name, elements });
  }
  return out;
}

/** A category a search found something in, and how many tiles it left there. */
export interface CategoryHit {
  key: string;
  name: string;
  count: number;
}

/**
 * Which categories a search found something in, with counts, for the chips
 * `docs/DESIGN.md` section 8 puts above the results while searching.
 *
 * Counted WITHOUT the category the user has already picked applied, which is
 * the whole point: the chips have to keep showing the other categories and
 * their counts, or picking one would hide the way back and the way across. The
 * query and the tags DO apply, because those are what the counts are counts of.
 *
 * Counts are tiles, not elements, so they agree with the number beside the
 * search and with what a person can see: a run of six sizes is one tile.
 */
export function categoryHits(library: Library, state: PaneState): CategoryHit[] {
  const free: PaneState = { ...state, category: undefined };
  return groups(library, free).map((g) => ({ key: g.key, name: g.name, count: g.elements.length }));
}

/**
 * Whether this member of a sized run matches the search.
 *
 * Section 8: a sized tile greys out the counts that do not match. The tile
 * itself is the first member that survived the search, and its stepper still
 * offers every size — so without this, a search for "3 boxes" shows a stepper
 * of six numbers with nothing to say which one was searched for.
 */
export function stepMatches(member: Element, state: PaneState): boolean {
  return matches(member, state.query) && state.tags.every((tag) => member.tags.includes(tag));
}

/**
 * Every member of a sized element's run, in count order, for the stepper.
 *
 * The filter is a type predicate rather than a plain test, and that is what
 * lets the sort read `a.run.count` outright. It used to say `a.run?.count ?? 0`
 * on both sides — two fallbacks the filter above had already made unreachable,
 * since an element can only match `el.run?.key === key` by HAVING a run. They
 * were the only thing the compiler would accept and they were also two branches
 * nothing could ever execute, which coverage duly reported on 2026-09-12.
 */
export function runOf(library: Library, element: Element): Element[] {
  if (!element.run) return [element];
  const key = element.run.key;
  const sized = (el: Element): el is Element & { run: NonNullable<Element["run"]> } => el.run?.key === key;
  return library.elements.filter(sized).sort((a, b) => a.run.count - b.run.count);
}

/** How many tiles a search left, for the count beside the categories. */
export function tileCount(found: Group[]): number {
  return found.reduce((n, g) => n + g.elements.length, 0);
}

/**
 * Whether a category is open.
 *
 * Categories start collapsed (`docs/DESIGN.md` section 4), and a search or a
 * tag opens what it finds — otherwise a query that matched three things in two
 * categories would show two closed headers and look like nothing was found.
 */
export function isOpen(state: PaneState, key: string): boolean {
  if (state.query.trim() !== "" || state.tags.length > 0) return true;
  return state.open.includes(key);
}

/**
 * Whether to offer "Open all" beside the count (`docs/DESIGN.md` section 4).
 *
 * Only while there is something to open. A search or a tag already opens
 * everything it found, so the control would do nothing then; and once every
 * category is open it has done its job and hides rather than staying on screen
 * as a button that changes nothing. There is no "Close all" beside it: the
 * design record asks for one control, and one category closes by its own
 * header the way it always has.
 */
export function offersOpenAll(state: PaneState, library: Library): boolean {
  if (state.query.trim() !== "" || state.tags.length > 0) return false;
  return library.categories.some((category) => !state.open.includes(category.key));
}

/** The element a tile id names, if the library still has it. */
export function elementOf(library: Library | undefined, id: string | undefined): Element | undefined {
  if (!library || id === undefined) return undefined;
  return library.elements.find((el) => el.id === id);
}

/**
 * How far apart two strings are, counting single-character edits.
 *
 * The ordinary Levenshtein distance, one row at a time so a 117-element library
 * costs a few thousand numbers rather than a matrix per name.
 *
 * No empty-string shortcuts. It had `if (a.length === 0) return b.length` and
 * its mirror, and both went on 2026-09-12: the first cannot run at all, because
 * `didYouMean` gives up on a query under three characters before it gets here,
 * and the second was doing nothing the loops below do not already do — with `b`
 * empty the row starts and ends at `[0]` and the answer is `a.length` anyway.
 * A shortcut nothing can take is a line that rots; an element with no name is
 * the case in `pane-search.test.ts` that holds the general path to the same
 * answer.
 */
function distance(a: string, b: string): number {
  if (a === b) return 0;
  let row = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const next = [i];
    for (let j = 1; j <= b.length; j++) {
      const substitute = (row[j - 1] as number) + (a[i - 1] === b[j - 1] ? 0 : 1);
      next[j] = Math.min(substitute, (row[j] as number) + 1, (next[j - 1] as number) + 1);
    }
    row = next;
  }
  return row[b.length] as number;
}

/**
 * What the user might have meant, for a query that found nothing.
 *
 * `docs/DESIGN.md` section 8. A search that returns an empty list is a dead end;
 * this turns it into a route, and it only ever offers names the library really
 * has, so picking one cannot fail.
 *
 * Measured against each WORD of a name as well as the whole of it, because a
 * query is usually one word and "triangel" should reach "Triangle, simple, with
 * text at the corners" — which as a whole string is 30 edits away from it.
 *
 * The threshold grows with the query: one edit for a short word, more for a
 * long one. Without that, a three-letter typo would reach half the library.
 */
export function didYouMean(library: Library, query: string, limit = 3): string[] {
  const q = query.trim().toLowerCase();
  if (q.length < 3) return [];
  const allowed = Math.max(1, Math.floor(q.length / 3));

  const scored: { name: string; cost: number }[] = [];
  for (const element of library.elements) {
    const name = element.name;
    const words = name
      .toLowerCase()
      .split(/[^\p{L}\p{N}]+/u)
      .filter(Boolean);
    let best = distance(q, name.toLowerCase());
    for (const word of words) best = Math.min(best, distance(q, word));
    if (best <= allowed) scored.push({ name, cost: best });
  }

  scored.sort((a, b) => a.cost - b.cost || a.name.localeCompare(b.name, "en"));
  const seen = new Set<string>();
  const out: string[] = [];
  for (const { name } of scored) {
    if (seen.has(name)) continue;
    seen.add(name);
    out.push(name);
    if (out.length === limit) break;
  }
  return out;
}
