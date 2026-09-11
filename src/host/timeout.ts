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
} as const;

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
