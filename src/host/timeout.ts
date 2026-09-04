/**
 * What each host call is allowed to cost.
 *
 * A call that stops answering never comes back: the pane sits on a spinner and
 * the user closes PowerPoint. Every Office.js call this add-in makes is bounded,
 * and the bound is a decision rather than a magic number at the call site.
 */
export const BUDGET = {
  /** Reading the deck's bytes. Big decks are slow and this is one pass. */
  file: 120_000,
  /** Handing the spliced slide back. */
  insert: 90_000,
  /** Removing the slide the insert replaced. */
  remove: 30_000,
  /** Small reads: the slide count, which slide is selected. */
  query: 15_000,
} as const;

/** Reject if the promise has not settled in time, naming what was waiting. */
export async function withTimeout<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      work,
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error(`${what} did not finish within ${Math.round(ms / 1000)}s`)), ms);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}
