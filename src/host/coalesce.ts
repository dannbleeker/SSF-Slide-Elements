/**
 * One job at a time, and always one more after the last request.
 *
 * PowerPoint tells a task pane that the selection changed, and it tells it
 * often: every slide clicked, every shape touched. Answering each one means a
 * `PowerPoint.run`, and the web forces a full presentation save on every
 * `context.sync()` (`CLAUDE.md`), so the pane may not simply read on each.
 *
 * The obvious guard is to drop an event that arrives while a read is already
 * running, and it is WRONG in a way that looks right in a slow test and fails
 * in front of a user. Measured on PowerPoint for the web on 2026-09-11:
 * clicking slides 2, 3 and 4 a quarter of a second apart left the pane saying
 * "Slide 3." for as long as anyone cared to watch, because the event for slide
 * 4 arrived while the read for slide 3 was still in flight and was thrown away.
 * There is no later event to put it right — the user has stopped clicking, and
 * the pane is now confidently naming the wrong slide.
 *
 * So the last request is remembered rather than dropped, and the job runs once
 * more when the one in flight finishes. Between a burst of ten events and one,
 * the difference is at most two runs: the one that was already going and the
 * one that settles it. What it never does is run ten.
 */
export function coalescing(work: () => Promise<void>): () => void {
  let running = false;
  let again = false;

  const run = async (): Promise<void> => {
    if (running) {
      // Not a queue. A second request while a third is waiting is the same
      // request: they all ask for the same thing, which is the CURRENT state.
      again = true;
      return;
    }
    running = true;
    try {
      do {
        // Cleared before the work, never after: an event that arrives DURING
        // this run must survive it. Clearing afterwards would swallow exactly
        // the event this function exists to keep.
        again = false;
        await work();
      } while (again);
    } finally {
      running = false;
    }
  };

  // The caller is an event handler and cannot await. A job that raises must not
  // leave `running` true for the rest of the session, which is what the
  // `finally` above is for; this only keeps the rejection from reaching the
  // host as an unhandled one.
  return () => {
    void run().catch(() => undefined);
  };
}
