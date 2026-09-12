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
    // decks in it differ in the last one. What this does NOT prove is that
    // `Math.imul` was needed: swapping it for a plain `*` leaves every case in
    // this file green, so the reason for `imul` is that it is the exact 32-bit
    // multiply FNV-1a specifies, not a measurement taken here.
    const long = `${ONEDRIVE}/${"deep/".repeat(200)}`;
    expect(deckKey(`${long}a.pptx`)).not.toBe(deckKey(`${long}b.pptx`));
    expect(deckKey(`${long}a.pptx`)).toMatch(new RegExp(`^${GLOBAL_KEY}:[0-9a-f]{8}$`));
  });
});
