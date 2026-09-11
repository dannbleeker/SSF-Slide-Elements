import { afterEach, describe, expect, it, vi } from "vitest";

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
