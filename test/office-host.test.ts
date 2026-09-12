import { afterEach, describe, expect, it, vi } from "vitest";
import { BUDGET, CONFIRM } from "../src/host/timeout.js";

/**
 * The one call the pane makes before it has decided anything.
 *
 * `src/office` mostly cannot run in the suite — it calls Office.js — but
 * `hostSupports` is a single expression over `Office.context`, and what it does
 * when that object is not what it expects decides whether an unsupported host
 * gets a sentence or a blank pane. `ready()` uses it to render "this
 * PowerPoint is too old", before anything else, so a raise here leaves the
 * user with nothing to read on exactly the host that needed the message.
 */
async function supportsWith(context: unknown): Promise<(v: string) => boolean> {
  vi.resetModules();
  (globalThis as unknown as { Office: unknown }).Office = { context };
  const mod = await import("../src/office/powerpoint.js");
  return mod.hostSupports;
}

afterEach(() => {
  delete (globalThis as unknown as { Office?: unknown }).Office;
});

describe("asking the host what it supports", () => {
  it("answers what the host says when the host can answer", async () => {
    // The ordinary path, so the guard below is not just "always false".
    const asked: string[] = [];
    const supports = await supportsWith({
      requirements: {
        isSetSupported: (name: string, v: string) => {
          asked.push(`${name} ${v}`);
          return v === "1.2";
        },
      },
    });
    expect(supports("1.2")).toBe(true);
    expect(supports("1.10")).toBe(false);
    expect(asked).toEqual(["PowerPointApi 1.2", "PowerPointApi 1.10"]);
  });

  it("answers false rather than throwing when there is no requirements object", async () => {
    const supports = await supportsWith({});
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when there is no context at all", async () => {
    const supports = await supportsWith(undefined);
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when the host raises", async () => {
    const supports = await supportsWith({
      requirements: {
        isSetSupported: () => {
          throw new Error("no");
        },
      },
    });
    expect(supports("1.2")).toBe(false);
  });

  it("is what ready() asks, so the pane's first sentence is the floor's", async () => {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: (_n: string, v: string) => v === "1.2" } },
    };
    const mod = await import("../src/office/powerpoint.js");
    expect(mod.ready().ok).toBe(true);
    (globalThis as unknown as { Office: unknown }).Office = { context: {} };
    vi.resetModules();
    const again = await import("../src/office/powerpoint.js");
    const refused = again.ready();
    expect(refused.ok).toBe(false);
    expect(refused.detail).toContain("1.2");
  });
});

/**
 * A deck read that comes back SHORT, and the two calls that turn a list into an
 * index.
 *
 * `docs/SIBLING.md` triaged office-js#4272 — a collection load over about fifty
 * items answers short on the web — as relevant to this add-in, and promised the
 * sibling's defence: page `getItemAt` at twenty. Nothing was paged. Three plain
 * collection loads shipped, and both of the calls below take an INDEX out of the
 * list they get: which slide the user is on, and which slide an undo aims its
 * insert at. The removal that follows an insert is positional, so an index off a
 * list that dropped a slide in the middle names somebody else's slide.
 *
 * The fake answers the load at SYNC, never at `load()`, because a double that
 * fills a value the moment it is asked for cannot produce the failure this is
 * about. It also lets the list and the count disagree on purpose, which is the
 * whole point: the guard is a comparison, and a fake that kept them in step
 * would make it untestable.
 */

interface Fake {
  deck: string[];
  /** What the collection load ANSWERS with. Short, whole, or reordered. */
  read?: string[];
  selected?: string[];
}

function install(fake: Fake): void {
  const read = fake.read ?? fake.deck;
  const selected = fake.selected ?? [];

  function collection(answers: string[]) {
    const box = {
      items: [] as { id: string }[],
      count: { value: -1 },
      load(_names: string) {
        return box;
      },
      getCount() {
        return box.count;
      },
      fill() {
        // On the RESPONSE, which is where a real host fills one.
        box.items = answers.map((id) => ({ id }));
        box.count.value = fake.deck.length;
      },
    };
    return box;
  }

  const all = collection(read);
  const chosen = collection(selected);

  const context = {
    presentation: {
      slides: all,
      getSelectedSlides: () => chosen,
    },
    sync: () => {
      all.fill();
      chosen.fill();
      // The selection is its own collection and its count is its own length,
      // not the deck's.
      chosen.count.value = selected.length;
      return Promise.resolve();
    },
  };

  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements: { isSetSupported: () => true } },
  };
  (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
    run: async (cb: (c: typeof context) => unknown) => await cb(context),
  };
}

async function host(fake: Fake) {
  vi.resetModules();
  install(fake);
  return await import("../src/office/powerpoint.js");
}

describe("asking the host to name the open deck", () => {
  /**
   * The other single-expression read over `Office.context`, and the same
   * question about it: what it does when the object is not what it expects.
   * `docs/DESIGN.md` section 4 keys the pane's per-deck memory on this, and
   * `deckKey` in `src/host/memory.ts` decides what an absent answer means — so
   * everything this has to get right is answering `undefined` rather than
   * raising, and never handing back an empty string as if it were a deck.
   */
  async function urlWith(context: unknown): Promise<string | undefined> {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = { context };
    const mod = await import("../src/office/powerpoint.js");
    return mod.deckUrl();
  }

  it("answers the URL when the host gives one", async () => {
    expect(await urlWith({ document: { url: "https://example.sharepoint.com/a.pptx" } })).toBe(
      "https://example.sharepoint.com/a.pptx",
    );
  });

  it("answers nothing for an unsaved deck, whether that is an empty string or no property", async () => {
    // Both shapes have been seen: the property can be present and empty.
    expect(await urlWith({ document: { url: "" } })).toBeUndefined();
    expect(await urlWith({ document: {} })).toBeUndefined();
    expect(await urlWith({})).toBeUndefined();
    expect(await urlWith(undefined)).toBeUndefined();
  });

  it("answers nothing rather than throwing when the host raises", async () => {
    expect(
      await urlWith({
        get document(): never {
          throw new Error("no");
        },
      }),
    ).toBeUndefined();
  });

  it("answers nothing for something that is not a string", async () => {
    expect(await urlWith({ document: { url: 7 } })).toBeUndefined();
  });
});

describe("a deck read that came back short", () => {
  afterEach(() => {
    delete (globalThis as unknown as { PowerPoint?: unknown }).PowerPoint;
  });

  it("names the slide the user is on when the read is whole", async () => {
    // The ordinary path first, so the refusals below are not just "always
    // undefined".
    const mod = await host({ deck: ["a", "b", "c"], selected: ["b"] });
    expect(await mod.currentSlide()).toEqual({ index: 1, id: "b" });
  });

  it("refuses to name an index when the list is shorter than the count", async () => {
    // The host dropped "c". Nothing about "b" is wrong here — the point is
    // that the next read might drop "a" instead, and then "b" is index 0.
    const mod = await host({ deck: ["a", "b", "c"], read: ["a", "b"], selected: ["b"] });
    expect(await mod.currentSlide()).toBeUndefined();
  });

  it("refuses even when the slide it wants is in the short list", async () => {
    // The tempting mistake: the id was found, so the index looks usable. It is
    // not. A read that dropped the FRONT makes every index after it wrong, and
    // the deck is what a removal is counted against.
    const mod = await host({ deck: ["a", "b", "c", "d"], read: ["b", "c"], selected: ["c"] });
    expect(await mod.currentSlide()).toBeUndefined();
  });

  it("names the slide at a position when the read is whole", async () => {
    const mod = await host({ deck: ["a", "b", "c"] });
    expect(await mod.slideIdAt(2)).toBe("c");
  });

  it("will not name the slide at a position off a short read", async () => {
    // Undo's insert aims at this id. Aiming it at whatever happens to sit at
    // index 2 of a list missing a slide is how a restored slide lands in the
    // wrong place, and the undo says so out loud rather than guessing.
    const mod = await host({ deck: ["a", "b", "c", "d"], read: ["a", "b", "c"] });
    expect(await mod.slideIdAt(2)).toBeUndefined();
  });

  it("still answers undefined for a position past the end of a whole read", async () => {
    const mod = await host({ deck: ["a", "b"] });
    expect(await mod.slideIdAt(5)).toBeUndefined();
  });
});

/**
 * A count that lags the call it is meant to be evidence for.
 *
 * On PowerPoint for the web the slide count can still be the OLD number after
 * an insert that has already happened — measured on 2026-09-11 by polling
 * `getCount()` every 300 ms through a real insert, where it stayed put for 2.8
 * seconds and then went up. The undo read it once, got the old number,
 * concluded the insert had not landed, and stopped between putting the user's
 * slide back and removing the rebuilt one. It said so honestly and the deck was
 * still wrong, with six slides where five belonged.
 *
 * The fake answers a SERIES, one value per read, so the lag is something a case
 * can state rather than something it has to wait for the real host to produce.
 * It also counts the reads, because "it eventually got there" is only half of
 * it: a confirm that kept asking after the deck agreed would be saves of the
 * user's presentation nobody asked for.
 */
function counting(series: number[]): { reads: () => number } {
  let read = 0;
  const context = {
    presentation: {
      slides: {
        getCount() {
          const box = { value: -1 };
          // On the RESPONSE. A double that answers when ASKED cannot show a
          // lag, because the lag is the gap between the two.
          pending.push(() => {
            box.value = series[Math.min(read, series.length - 1)] ?? -1;
            read += 1;
          });
          return box;
        },
      },
    },
    sync: () => {
      for (const fill of pending.splice(0)) fill();
      return Promise.resolve();
    },
  };
  const pending: (() => void)[] = [];
  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements: { isSetSupported: () => true } },
  };
  (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
    run: async (cb: (c: typeof context) => unknown) => await cb(context),
  };
  return { reads: () => read };
}

async function counter(series: number[]) {
  vi.resetModules();
  const spy = counting(series);
  const mod = await import("../src/office/powerpoint.js");
  return { countReaching: mod.countReaching, reads: spy.reads };
}

describe("confirming that the deck changed size", () => {
  afterEach(() => {
    delete (globalThis as unknown as { PowerPoint?: unknown }).PowerPoint;
  });

  it("costs one read when the deck already agrees", async () => {
    const { countReaching, reads } = await counter([6]);
    expect(await countReaching(6)).toBe(6);
    expect(reads(), "the ordinary case must not pay for the rare one").toBe(1);
  });

  it("keeps asking while the count lags, and stops the moment it agrees", async () => {
    // The failure this exists for: two stale answers, then the truth.
    const { countReaching, reads } = await counter([5, 5, 6, 6, 6]);
    expect(await countReaching(6)).toBe(6);
    expect(reads(), "it must stop at the answer, not run the backoff out").toBe(3);
  });

  it("answers what it last saw when the deck never agrees", async () => {
    // A real no-op, and it must still be reported as one. Five reads is the
    // whole backoff; the number that comes back is the deck's, not the hope.
    const { countReaching, reads } = await counter([5]);
    expect(await countReaching(6)).toBe(5);
    expect(reads()).toBe(CONFIRM.length + 1);
  });

  it("answers a deck that overshot, rather than waiting for a number it will never see", async () => {
    // Two slides where one was expected. `outcomeOf` has a sentence for it, and
    // it can only say it if this hands back what is really there.
    const { countReaching } = await counter([7]);
    expect(await countReaching(6)).toBe(7);
  });

  it("backs off rather than polling evenly, because every read saves the deck", () => {
    // PowerPoint on the web forces a full presentation save on every
    // `context.sync()`, read-only ones included (`CLAUDE.md`). A tight poll
    // over the same span would be three times the saves for the same answer.
    expect(CONFIRM.length).toBeGreaterThan(2);
    for (let i = 1; i < CONFIRM.length; i++) {
      expect(CONFIRM[i], "each wait is longer than the one before").toBeGreaterThan(CONFIRM[i - 1] ?? 0);
    }
    const total = CONFIRM.reduce((a, b) => a + b, 0);
    expect(total, "the whole backoff has to fit inside one read's budget").toBeLessThan(BUDGET.read);
  });
});

/**
 * "There is no selection" and "the host did not answer" are different answers.
 *
 * They were one value, `undefined`, and the pane's slide line cannot live with
 * that: a read that runs out of time would replace a perfectly good slide
 * number with "PowerPoint did not say which slide you are on", which the host
 * never said. Measured on the web on 2026-09-11, where a selection read fired
 * just after an insert sat unanswered for most of a twenty-second budget.
 */
describe("a selection read that does not come back", () => {
  afterEach(() => {
    delete (globalThis as unknown as { PowerPoint?: unknown }).PowerPoint;
  });

  it("answers null when the host does not reply inside the budget", async () => {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => true } },
    };
    (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
      run: () => new Promise(() => undefined),
    };
    const mod = await import("../src/office/powerpoint.js");
    // A budget of its own, so the case costs milliseconds rather than seconds.
    expect(await mod.currentSlide(40)).toBeNull();
  });

  it("answers undefined when the host replies that nothing is selected", async () => {
    // The other half, and the reason the two may not be the same value: this
    // one IS the host speaking, and the pane should say so.
    vi.resetModules();
    install({ deck: ["a", "b"], selected: [] });
    const mod = await import("../src/office/powerpoint.js");
    expect(await mod.currentSlide()).toBeUndefined();
  });

  it("answers null when the host raises rather than answering", async () => {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => true } },
    };
    (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
      run: () => Promise.reject(new Error("the host is busy")),
    };
    const mod = await import("../src/office/powerpoint.js");
    expect(await mod.currentSlide()).toBeNull();
  });

  it("keeps a glance well under the budget of a read somebody is waiting on", () => {
    // The point of the separate number: a line nobody is waiting on may not
    // hold up the thing that keeps it fresh.
    expect(BUDGET.glance).toBeLessThan(BUDGET.read);
    expect(BUDGET.glance).toBeGreaterThan(1000);
  });
});

/**
 * The one selection WRITE this add-in makes, and the read-back it never trusts
 * the call without.
 *
 * `src/host/jump.ts` says why the call may be made at all; what this file holds
 * is the office layer's half of the contract: the id asked for reaches the
 * host unchanged, the selection is read back in the same batch, a host below
 * 1.5 is never asked, and a host that raises or goes silent comes back as
 * "did not answer" rather than as a jump.
 */
describe("selecting a slide and reading the selection back", () => {
  afterEach(() => {
    delete (globalThis as unknown as { PowerPoint?: unknown }).PowerPoint;
  });

  /** A host whose selection answers whatever the last write asked for, or a fixed list. */
  function selecting(options: { answers?: string[]; supports?: boolean; throws?: string }) {
    const written: string[][] = [];
    let current: string[] = ["256#1"];
    const chosen = {
      items: [] as { id: string }[],
      load() {
        return chosen;
      },
    };
    const context = {
      presentation: {
        setSelectedSlides: (ids: string[]) => {
          if (options.throws !== undefined) throw new Error(options.throws);
          written.push(ids);
          current = options.answers ?? ids;
        },
        getSelectedSlides: () => chosen,
      },
      sync: () => {
        // On the response, where a real host fills a collection.
        chosen.items = current.map((id) => ({ id }));
        return Promise.resolve();
      },
    };
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: () => options.supports !== false } },
    };
    (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
      run: async (cb: (c: typeof context) => unknown) => await cb(context),
    };
    return written;
  }

  it("asks for exactly the id it was given and answers what the host selected afterwards", async () => {
    vi.resetModules();
    const written = selecting({});
    const mod = await import("../src/office/powerpoint.js");
    expect(await mod.selectSlide("260#77")).toEqual({ supported: true, selected: ["260#77"] });
    expect(written).toEqual([["260#77"]]);
  });

  it("reports the host's answer even when it is not the slide asked for, and decides nothing", async () => {
    // The judgement lives in src/host/jump.ts; this only carries the answer.
    vi.resetModules();
    selecting({ answers: ["259#1"] });
    const mod = await import("../src/office/powerpoint.js");
    expect((await mod.selectSlide("260#77")).selected).toEqual(["259#1"]);
  });

  it("never asks a host below PowerPointApi 1.5", async () => {
    vi.resetModules();
    const written = selecting({ supports: false });
    const mod = await import("../src/office/powerpoint.js");
    expect(await mod.selectSlide("260#77")).toEqual({ supported: false, selected: null });
    expect(written).toEqual([]);
  });

  it("answers the reason, not a jump, when the host raises", async () => {
    // office-js#3552: desktop throws while the notes pane has focus.
    vi.resetModules();
    selecting({ throws: "GeneralException" });
    const mod = await import("../src/office/powerpoint.js");
    const seen = await mod.selectSlide("260#77");
    expect(seen.supported).toBe(true);
    expect(seen.selected).toBeNull();
    expect(seen.error).toContain("GeneralException");
  });

  it("answers null, inside the read budget, when the host goes silent", async () => {
    vi.useFakeTimers();
    try {
      vi.resetModules();
      (globalThis as unknown as { Office: unknown }).Office = {
        context: { requirements: { isSetSupported: () => true } },
      };
      (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = { run: () => new Promise(() => undefined) };
      const mod = await import("../src/office/powerpoint.js");
      const pending = mod.selectSlide("260#77");
      await vi.advanceTimersByTimeAsync(BUDGET.read + 1);
      const seen = await pending;
      expect(seen.selected).toBeNull();
      expect(seen.error).toContain("moving to a slide");
    } finally {
      vi.useRealTimers();
    }
  });
});
