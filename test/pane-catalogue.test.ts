import { afterEach, describe, expect, it, vi } from "vitest";
import { Store, carriedTypes, dirOf, libraryFor, loadIndex, nameOfRatio, type Index } from "../src/pane/catalogue.js";

/**
 * The one file in the add-in that touches the network.
 *
 * `test/security.test.ts` holds what it may do — GET, from its own origin, with
 * no request options — and this holds what it DOES: which library a deck of a
 * given shape gets, what the pane says about a borrowed one, and that a part is
 * fetched once however many times it is asked for.
 *
 * `fetch` is stubbed rather than served, because a test that started a web
 * server would be testing the server. What matters here is the URL asked for
 * and what is done with the answer.
 */

const INDEX: Index = {
  version: "abc123",
  sizes: {
    "16:9": {
      size: "16:9",
      width: 12192000,
      height: 6858000,
      categories: [{ key: "boxes", name: "White boxes" }],
      elements: [],
      carried: { "ppt/media/image1.png": "image/png" },
    },
    "4:3": {
      size: "4:3",
      width: 9144000,
      height: 6858000,
      categories: [],
      elements: [],
      carried: {},
    },
  },
};

/** Every URL the module asked for, in order, so the paths themselves are asserted. */
let asked: string[] = [];

function serve(answers: Record<string, { body?: string | Uint8Array; status?: number }>): void {
  vi.stubGlobal("fetch", (url: string) => {
    asked.push(url);
    const found = answers[url];
    if (!found) return Promise.resolve({ ok: false, status: 404 });
    if ((found.status ?? 200) >= 400) return Promise.resolve({ ok: false, status: found.status });
    const body = found.body ?? "";
    return Promise.resolve({
      ok: true,
      status: 200,
      text: () => Promise.resolve(typeof body === "string" ? body : new TextDecoder().decode(body)),
      arrayBuffer: () =>
        Promise.resolve(
          typeof body === "string" ? new TextEncoder().encode(body).buffer : (body.buffer as ArrayBuffer),
        ),
    });
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
  asked = [];
});

describe("where the files are", () => {
  it("turns a slide size into a path segment, because a colon is not one", () => {
    expect(dirOf("16:9")).toBe("16x9");
    expect(dirOf("4:3")).toBe("4x3");
  });
});

describe("naming a slide shape in a sentence", () => {
  it("knows the two the design names, and falls back to the ratio itself", () => {
    expect(nameOfRatio(16 / 9)).toBe("16:9");
    expect(nameOfRatio(4 / 3)).toBe("4:3");
    expect(nameOfRatio(1.414)).toBe("A4");
    expect(nameOfRatio(1.6)).toBe("16:10");
    // A number is not pretty and it is honest: a user with a custom size
    // recognises their own deck in it, where "custom" tells them nothing.
    expect(nameOfRatio(1.25)).toBe("1.25:1");
  });
});

describe("which library a deck gets", () => {
  it("gives a 16:9 deck the 16:9 library, with nothing to say about it", () => {
    const library = libraryFor(INDEX, 12192000, 6858000);
    expect(library.size).toBe("16:9");
    expect(library.version).toBe("abc123");
    expect(library.borrowed).toBeUndefined();
  });

  it("gives a 4:3 deck the 4:3 library", () => {
    expect(libraryFor(INDEX, 9144000, 6858000).size).toBe("4:3");
  });

  it("borrows the NEAREST library by aspect ratio for a deck that is neither, and says so", () => {
    // A4 landscape is 1.414, which is nearer 4:3 (1.333) than 16:9 (1.778).
    const library = libraryFor(INDEX, 10692000, 7560000);
    expect(library.size).toBe("4:3");
    expect(library.borrowed).toBe("4:3 library, scaled to A4 slides.");
  });

  it("borrows 16:9 for a widescreen deck that is not quite 16:9", () => {
    const library = libraryFor(INDEX, 12192000, 7620000);
    expect(library.size).toBe("16:9");
    expect(library.borrowed).toBe("16:9 library, scaled to 16:10 slides.");
  });

  it("never refuses: a deck with no height at all still gets a library", () => {
    // An element that lands slightly scaled is worth more than a pane that will
    // not open on somebody's odd deck.
    expect(libraryFor(INDEX, 12192000, 0).size).toBe("16:9");
  });

  it("says so by name when the index carries no libraries at all", () => {
    expect(() => libraryFor({ version: "v", sizes: {} }, 100, 100)).toThrow(/carries no libraries/);
  });
});

describe("what each carried part is", () => {
  it("hands back the content types the harvest read off the library deck", () => {
    expect(carriedTypes(INDEX, "16:9")).toEqual({ "ppt/media/image1.png": "image/png" });
  });

  it("answers an empty map for a size the index does not have, rather than undefined", () => {
    // The splice spreads this into a lookup; undefined there would be a crash
    // at insert time instead of a part with no declared type.
    expect(carriedTypes(INDEX, "5:4")).toEqual({});
  });
});

describe("loading the index", () => {
  it("asks for it beside the pane, never at an absolute address", async () => {
    serve({ "./catalogue/catalogue.json": { body: JSON.stringify(INDEX) } });
    const index = await loadIndex();
    expect(index.version).toBe("abc123");
    expect(asked).toEqual(["./catalogue/catalogue.json"]);
  });

  it("says which file answered what, when one does not", async () => {
    serve({});
    await expect(loadIndex()).rejects.toThrow(/catalogue\.json answered 404/);
  });
});

describe("the store", () => {
  it("fetches an element's markup from its own size's directory, once", async () => {
    const markup = { xml: "<p:sp/>", rels: [], parts: [] };
    serve({ "./catalogue/16x9/elements/one-box.json": { body: JSON.stringify(markup) } });
    const store = new Store("16:9");
    const element = { id: "one-box" } as never;
    expect(await store.markup(element)).toEqual(markup);
    await store.markup(element);
    // Cached: a user who inserts the same element twice pays for it once.
    expect(asked).toEqual(["./catalogue/16x9/elements/one-box.json"]);
  });

  it("reads an XML part as text and a picture as bytes, which is what the splice needs", async () => {
    serve({
      "./catalogue/16x9/parts/ppt/tags/tag1.xml": { body: "<p:tagLst/>" },
      "./catalogue/16x9/parts/ppt/media/image1.png": { body: new Uint8Array([1, 2, 3]) },
    });
    const store = new Store("16:9");
    expect(await store.part("ppt/tags/tag1.xml")).toBe("<p:tagLst/>");
    const picture = await store.part("ppt/media/image1.png");
    expect(picture).toBeInstanceOf(Uint8Array);
    expect([...(picture as Uint8Array)]).toEqual([1, 2, 3]);
  });

  it("answers undefined for a part that is not there, so the splice can name it", async () => {
    serve({});
    expect(await new Store("16:9").part("ppt/media/gone.png")).toBeUndefined();
  });

  it("forgets a part it could not fetch, so a dropped network is retried", async () => {
    // Remembered as absent forever, one bad minute would break every later
    // insert of that element for the rest of the session.
    serve({});
    const store = new Store("16:9");
    expect(await store.part("ppt/media/x.png")).toBeUndefined();
    serve({ "./catalogue/16x9/parts/ppt/media/x.png": { body: new Uint8Array([9]) } });
    expect(await store.part("ppt/media/x.png")).toBeInstanceOf(Uint8Array);
  });

  it("escapes each path segment separately, so the slashes survive", async () => {
    serve({ "./catalogue/16x9/parts/ppt/media/a%20b.png": { body: new Uint8Array([1]) } });
    await new Store("16:9").part("ppt/media/a b.png");
    expect(asked[0]).toBe("./catalogue/16x9/parts/ppt/media/a%20b.png");
  });
});
