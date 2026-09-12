/**
 * Going to a slide from "Used in this deck", decided away from the host.
 *
 * `docs/DESIGN.md` section 4: a slide number in that list jumps to the slide.
 * The call is `presentation.setSelectedSlides([id])`, PowerPointApi 1.5, and
 * the reason it is safe to make is BORROWED, dated, and read back on every
 * click rather than trusted:
 *
 * - SSF-Charts ships the same call (`showSlide`) and measures it in every
 *   round of its self-test battery. Its archive from 2026-08-13 to 2026-09-04
 *   holds 2,429 selection-ladder rungs over about 347 rounds on PowerPoint for
 *   the web, every one answered, none silent, none refused. The wedge that
 *   family designed around was `setSelectedShapes([id])`, measured once on
 *   build `55011a3`; `setSelectedSlides` went silent only downstream of it and
 *   has never wedged anything on its own (`docs/SIBLING.md`).
 * - The tracker has no open issue against it on the web. On Windows and Mac it
 *   THROWS while the notes pane has focus (office-js#3552, build 2306); a throw
 *   is a sentence here, not a hang. On desktop `getSelectedSlides` answers the
 *   active slide even when nothing is selected (office-js#4877), which is what
 *   a read-back wants.
 * - No sibling measured the call on Windows or Mac. So this file never claims a
 *   jump it did not see: the office layer reads the selection back after the
 *   call, and the sentence says "Slide N" only when the read-back names slide
 *   N. Everything else says which slide to click.
 *
 * Pure on purpose: `src/office/powerpoint.ts` makes the call and returns what
 * the host answered; the reading of that answer is here, where the suite can
 * hold every branch.
 */

/**
 * Whether two slide ids name the same slide.
 *
 * A selection id can lack the `#suffix` the deck's own list carries
 * (office-js#2474, reported on Windows desktop and closed not planned), so the
 * two are compared on the half before the `#` when either side has none. Both
 * suffixed and different is different; both prefixes different is different.
 */
export function sameSlideId(a: string, b: string): boolean {
  if (a === b) return true;
  const [ap, as] = split(a);
  const [bp, bs] = split(b);
  if (ap !== bp || ap === "") return false;
  return as === undefined || bs === undefined;
}

function split(id: string): [string, string | undefined] {
  const at = id.indexOf("#");
  return at < 0 ? [id, undefined] : [id.slice(0, at), id.slice(at + 1)];
}

export interface JumpObservation {
  /** The slide number the user clicked, counting from one. */
  slide: number;
  /** The id the pane asked the host to select, read positionally a moment before. */
  wanted: string;
  /** Whether the host has the call at all (PowerPointApi 1.5). */
  supported: boolean;
  /** What `getSelectedSlides` answered AFTER the call; null when the host did not answer inside the budget. */
  selected: string[] | null;
  /** The reason the call raised, bounded, if it did. */
  error?: string;
}

export interface JumpOutcome {
  /** True only when the host was seen on the slide afterwards. */
  ok: boolean;
  /** The sentence for the footer or the live region. */
  detail: string;
}

/** What to tell the user when the pane could not see the jump happen. */
const clickIt = (slide: number): string => `Click slide ${slide} in the strip.`;

/**
 * Whether the jump happened, from what the host answered afterwards.
 *
 * The first selected id is the active slide, which is the one the editing
 * area shows (the API's own contract for `getSelectedSlides`), so that is the
 * one compared. A host that answered something else did not move; a host that
 * answered nothing inside the budget is not assumed to have moved either.
 */
export function jumpOutcome(o: JumpObservation): JumpOutcome {
  if (!o.supported) {
    return {
      ok: false,
      detail: `This PowerPoint cannot move to a slide from the pane (it needs PowerPointApi 1.5). ${clickIt(o.slide)}`,
    };
  }
  if (o.error !== undefined) {
    return { ok: false, detail: `Could not go to slide ${o.slide}: ${o.error}. ${clickIt(o.slide)}` };
  }
  if (o.selected === null) {
    return { ok: false, detail: `PowerPoint did not say which slide it is on. ${clickIt(o.slide)}` };
  }
  const first = o.selected[0];
  if (first !== undefined && sameSlideId(first, o.wanted)) {
    return { ok: true, detail: `Slide ${o.slide}` };
  }
  return { ok: false, detail: `PowerPoint did not move to slide ${o.slide}. ${clickIt(o.slide)}` };
}
