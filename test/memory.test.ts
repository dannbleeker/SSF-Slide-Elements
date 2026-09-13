import { describe, expect, it } from "vitest";
import { GLOBAL_KEY, deckKey } from "../src/host/memory.js";

/**
 * Which bucket the pane remembers itself in (`docs/DESIGN.md` section 4).
 *
 * Every case here is about a URL a real host hands over: a OneDrive link whose
 * query string changes between sessions, an unsaved deck with no URL at all, a
 * local Windows path. What the key IS does not matter to anything; that the
 * same deck answers the same key, and two decks do not, is the whole contract.
 */
describe("the bucket one deck remembers itself in", () => {
  const ONEDRIVE = "https://contoso-my.sharepoint.com/personal/someone/Documents/Q4%20review.pptx";

  it("falls back to the per-machine bucket when the host will not name a deck", () => {
    // An unsaved deck on the web has no URL, and the property can be an empty
    // string rather than absent. Forgetting outright would mean losing a search
    // on every close, which is the behaviour the feature exists to remove.
    expect(deckKey(undefined)).toBe(GLOBAL_KEY);
    expect(deckKey("")).toBe(GLOBAL_KEY);
    expect(deckKey("   ")).toBe(GLOBAL_KEY);
  });

  it("answers a bucket of its own for a deck the host does name", () => {
    const key = deckKey(ONEDRIVE);
    expect(key).not.toBe(GLOBAL_KEY);
    expect(key.startsWith(`${GLOBAL_KEY}:`)).toBe(true);
  });

  it("answers the same bucket for the same deck every time", () => {
    expect(deckKey(ONEDRIVE)).toBe(deckKey(ONEDRIVE));
  });

  it("ignores the query string and the fragment, which come and go between sessions", () => {
    // A OneDrive or SharePoint URL for one file is not one string: `?web=1`,
    // `?d=w…` and a `#` anchor appear and disappear, and a key that changed
    // with them would remember nothing across the closes it is meant to
    // survive.
    expect(deckKey(`${ONEDRIVE}?web=1`)).toBe(deckKey(ONEDRIVE));
    expect(deckKey(`${ONEDRIVE}#page=3`)).toBe(deckKey(ONEDRIVE));
    expect(deckKey(`${ONEDRIVE}?d=w1234&csf=1#anchor`)).toBe(deckKey(ONEDRIVE));
  });

  it("ignores the case, since a host can hand back either", () => {
    expect(deckKey(ONEDRIVE.toUpperCase())).toBe(deckKey(ONEDRIVE.toLowerCase()));
  });

  it("tells two decks apart, including two in the same folder", () => {
    const other = ONEDRIVE.replace("Q4%20review", "Q3%20review");
    expect(deckKey(other)).not.toBe(deckKey(ONEDRIVE));
    expect(deckKey("C:\\Users\\me\\Desktop\\a.pptx")).not.toBe(deckKey("C:\\Users\\me\\Desktop\\b.pptx"));
  });

  it("does not put the URL itself in the key", () => {
    // A SharePoint path can name a client, a project or a person, and only
    // equality is ever asked of this — so there is no reason for it to sit in
    // storage where anything else on the origin could read it back.
    const key = deckKey(ONEDRIVE);
    expect(key).not.toContain("sharepoint");
    expect(key).not.toContain("someone");
    expect(key).not.toContain("review");
    // Eight hex digits after the prefix, and nothing else.
    expect(key).toMatch(new RegExp(`^${GLOBAL_KEY}:[0-9a-f]{8}$`));
  });

  it("still answers a well-formed key for a very long path, and reads its tail", () => {
    // A deeply nested SharePoint path is a thousand characters, and the two
    // decks in it differ in the last one.
    const long = `${ONEDRIVE}/${"deep/".repeat(200)}`;
    expect(deckKey(`${long}a.pptx`)).not.toBe(deckKey(`${long}b.pptx`));
    expect(deckKey(`${long}a.pptx`)).toMatch(new RegExp(`^${GLOBAL_KEY}:[0-9a-f]{8}$`));
  });

  it("is FNV-1a over the normalized URL, and the digits are the published ones", () => {
    // The only case here that says what the hash IS, and it exists because
    // nothing else in this file can hold the loop that produces it. A hash
    // that reads one character too few, one character too many, or throws the
    // top bit away still answers eight hex digits, still answers the same
    // digits for the same deck, and still tells two decks apart — so every
    // case above stays green while the digest is wrong.
    //
    // `src/host/memory.ts` says FNV-1a, 32 bits, in hex (`docs/DESIGN.md`
    // section 4, "the deck is told apart by a HASH of its URL"), so the
    // numbers pinned are FNV-1a's own published 32-bit vectors: "a" is
    // 0xe40c292c and "foobar" is 0xbf9cf968. Both pass through `normalize`
    // unchanged, so what is hashed is exactly the vector. `bf9cf968` has its
    // top bit set, which is what makes the unsigned `>>> 0` visible: a signed
    // read or a shift of one would not answer these digits.
    //
    // Unlike every other case here, this one is sensitive to the multiply.
    // Measured 2026-09-12, by swapping `Math.imul` for a plain `*`: the
    // one-character vector below stays green — a double still holds that
    // product's low 32 bits — and "foobar" goes red at 0ee3c7f0, because by the
    // sixth character the exact-integer range is long gone. So the `imul` is
    // load-bearing from here, and the note in `src/host/memory.ts` saying this
    // file stays green with `*` is now out of date by one case.
    expect(deckKey("a")).toBe(`${GLOBAL_KEY}:e40c292c`);
    expect(deckKey("foobar")).toBe(`${GLOBAL_KEY}:bf9cf968`);
  });

  it("reads the URL from its first character, not from its second", () => {
    // The other side of the same loop: the same deck path on two drives is one
    // character apart, at index 0, and a hash that started at index 1 would
    // hand both decks the same bucket and each the other's search.
    expect(deckKey("C:\\Users\\me\\Desktop\\deck.pptx")).not.toBe(deckKey("D:\\Users\\me\\Desktop\\deck.pptx"));
  });
});
