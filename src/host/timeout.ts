/**
 * Every call to the host is bounded.
 *
 * Ported from SSF-Merge (`src/host/timeout.ts`) on 2026-09-10. On PowerPoint
 * for the web a stall is death rather than slowness: the pane has no way to
 * cancel a queued batch, so a call that never settles leaves the add-in busy
 * for the rest of the session with a spinner and no sentence. The probe carries
 * the same wrapper for the same reason, and its own rounds have hit it — the
 * second pair of answer sheets on 2026-09-10 recorded an insert that took
 * longer than the probe's 120-second budget on a degraded document session, and
 * came back as a timed-out call that had nonetheless landed its slide.
 *
 * Which is the whole subtlety, and why this file only ever produces a REASON
 * rather than a verdict: **a call can raise and still have done the work**.
 * Nothing may read a timeout as "it did not happen". `src/host/insert.ts` reads
 * the deck's size instead.
 */

/**
 * How long each kind of call may take before the pane stops waiting.
 *
 * Numbers rather than one budget, because the calls differ by two orders of
 * magnitude and a single value would either abandon a large deck's read or let
 * a wedged count hang for two minutes. Measured on the web on 2026-09-10:
 * `getFileAsync` answered a 34 KB deck in 874 ms on a healthy session and took
 * 40 seconds for 40 KB on a session that had been through a timeout reload, so
 * the read's budget is generous on purpose — the alternative is telling a user
 * their deck is unreadable because their network had a bad minute.
 */
export const BUDGET = {
  /** A count or an id read: a batch with nothing in it. */
  read: 20_000,
  /** `getFileAsync` over the user's whole presentation. */
  deck: 180_000,
  /** `insertSlidesFromBase64` with that whole presentation as the argument. */
  insert: 180_000,
  /** A positional delete: one batch, one slide. */
  remove: 60_000,
  /**
   * A read whose answer only keeps a LINE fresh, never a read an insert
   * depends on.
   *
   * Twenty seconds is right for a read the user is waiting on and wrong for one
   * they are not. Measured on the web on 2026-09-11: a selection read fired
   * straight after an insert can sit unanswered for most of `read`'s budget
   * while the host finishes writing the deck — and because the pane answers one
   * selection event at a time, that one read froze the line naming the current
   * slide for the whole window. Three clicks on three different slides went by
   * with the pane still naming the first.
   *
   * So a glance gives up early. The line then keeps the number it has, which is
   * the right answer to "I could not find out": it is still the last thing the
   * host actually said, and the insert reads the selection again for itself
   * regardless.
   */
  glance: 4_000,
} as const;

/**
 * The pauses between reads when confirming that the deck changed size.
 *
 * **The slide count can lag an insert that has already happened.** Measured on
 * PowerPoint for the web on 2026-09-11, polling `slides.getCount()` every
 * 300 ms through a real insert: the count stayed at its old value for 2.8
 * seconds after the click and then went up. The undo read it once, immediately
 * after `insertSlidesFromBase64` resolved, got the old number, concluded the
 * insert had not landed and stopped — leaving the user with the slide it had
 * just put back AND the rebuilt one it never removed. It reported the failure
 * honestly and the deck was still wrong.
 *
 * So a size that decides anything is read more than once. Five reads over about
 * six and a half seconds, stopping the moment the deck agrees, which it almost
 * always does on the first. Backed off rather than evenly spaced because **the
 * web forces a full presentation save on every `context.sync()`**
 * (`CLAUDE.md`), read-only ones included: a tight poll would be fifteen saves
 * of the user's deck to answer one question.
 *
 * This is re-reading the MEASUREMENT, never re-trying the call. Nothing here
 * inserts or removes twice; a second insert on a deck whose shape this code has
 * already misread is how the wrong slide goes.
 */
export const CONFIRM = [500, 1000, 2000, 3000] as const;

/** A promise that rejects by name if it has not settled in time. */
export function withTimeout<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`ssf-slide-elements: gave up waiting for ${what} after ${Math.round(ms / 1000)} seconds`));
    }, ms);
    work.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error: unknown) => {
        clearTimeout(timer);
        reject(error instanceof Error ? error : new Error(String(error)));
      },
    );
  });
}
