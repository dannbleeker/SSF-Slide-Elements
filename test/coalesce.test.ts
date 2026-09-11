import { describe, expect, it } from "vitest";
import { coalescing } from "../src/host/coalesce.js";

/**
 * The rule that keeps the pane from naming a slide the user has left.
 *
 * PowerPoint tells a task pane that the selection changed far more often than
 * the pane can afford to answer, so the answering has to be throttled — and the
 * obvious throttle is the wrong one. Dropping an event because a read is
 * already running loses the LAST event of a burst, which is the only one whose
 * answer the user is waiting for. Measured on the web on 2026-09-11: slides 2,
 * 3 and 4 clicked a quarter of a second apart left the pane saying "Slide 3."
 * for as long as anyone cared to watch.
 *
 * Every case below drives the work by hand rather than by timer, so what is
 * being asserted is the ORDER of things and not a duration.
 */
function pending(): { work: () => Promise<void>; runs: number; finish: () => void; waiting: number } {
  const gates: (() => void)[] = [];
  const state = {
    runs: 0,
    get waiting() {
      return gates.length;
    },
    work: () =>
      new Promise<void>((resolve) => {
        state.runs += 1;
        gates.push(resolve);
      }),
    finish: () => {
      const next = gates.shift();
      if (!next) throw new Error("nothing was waiting to be finished");
      next();
    },
  };
  return state;
}

/** Let every already-resolved promise settle, without waiting on a clock. */
const settle = async (): Promise<void> => {
  for (let i = 0; i < 5; i++) await Promise.resolve();
};

describe("one job at a time, and one more after the last request", () => {
  it("runs straight away when nothing is running", async () => {
    const job = pending();
    const ask = coalescing(job.work);
    ask();
    await settle();
    expect(job.runs).toBe(1);
  });

  it("runs ONCE MORE for requests that arrived while it was busy", async () => {
    // The defect, as a case. Three clicks, one read in flight: the second and
    // third must not vanish, and must not become two more reads either.
    const job = pending();
    const ask = coalescing(job.work);
    ask();
    await settle();
    expect(job.runs, "the first request runs immediately").toBe(1);

    ask();
    ask();
    await settle();
    expect(job.runs, "nothing new starts while one is in flight").toBe(1);

    job.finish();
    await settle();
    expect(job.runs, "the requests made during the first run collapse into one").toBe(2);

    job.finish();
    await settle();
    expect(job.runs, "and then it stops, because nothing asked again").toBe(2);
  });

  it("is listening again once a run has finished", async () => {
    // Not about the burst: this is the release. A request that arrives after
    // everything settled has to start a fresh run, which it cannot do if
    // `running` was left true.
    const job = pending();
    const ask = coalescing(job.work);
    ask();
    await settle();
    job.finish();
    await settle();
    expect(job.runs).toBe(1);

    ask();
    await settle();
    expect(job.runs, "a request after the run finished starts a new one").toBe(2);
  });

  it("is ready again after a job that raised", async () => {
    // A read that fails must not leave the pane deaf for the rest of the
    // session, which is what a flag set outside a `finally` would do.
    let calls = 0;
    const ask = coalescing(() => {
      calls += 1;
      return Promise.reject(new Error("the host would not say"));
    });
    ask();
    await settle();
    expect(calls).toBe(1);
    ask();
    await settle();
    expect(calls, "the next event is still answered").toBe(2);
  });

  it("collapses a long burst into two runs, not into ten", async () => {
    const job = pending();
    const ask = coalescing(job.work);
    ask();
    await settle();
    for (let i = 0; i < 9; i++) ask();
    await settle();
    job.finish();
    await settle();
    expect(job.runs, "the one that was going, and the one that settles it").toBe(2);
    expect(job.waiting, "exactly one is still in flight").toBe(1);
  });
});
