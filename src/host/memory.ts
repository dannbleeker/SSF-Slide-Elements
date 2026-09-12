/**
 * Which bucket of the browser's storage the pane remembers itself in.
 *
 * `docs/DESIGN.md` section 4 splits what is remembered in two. Favourites and
 * the first-run flag are per MACHINE: a star is a statement about the library,
 * and the coach marks are "dismissed once", not once per deck. Everything about
 * how you were reading the library — the search, the tags, the open categories,
 * the size you picked in a stepper, Recent, and the settings — is per DECK,
 * because a deck is the unit of work and the way you were browsing for one is
 * rarely the way you want to browse for the next.
 *
 * Pure on purpose: `src/office/powerpoint.ts` asks the host for the deck's URL
 * and this decides what to do with it, so every guard below is one the suite
 * can hold without a PowerPoint.
 */

/** The per-machine bucket, and the fallback when the host will not name a deck. */
export const GLOBAL_KEY = "ssf-slide-elements";

/**
 * The bucket for one deck, from the URL the host gave for it.
 *
 * Three things this does, each because of something a real URL does:
 *
 * - **It answers `GLOBAL_KEY` when there is no usable URL.** An unsaved deck on
 *   the web has none, and `Office.context.document.url` can be an empty string
 *   rather than absent. Falling back means every unsaved deck shares one
 *   bucket, which is a compromise; forgetting outright would mean a pane that
 *   loses a search every time it is closed, which is the behaviour this feature
 *   exists to remove.
 * - **It drops the query and the fragment, and lowercases.** A OneDrive or
 *   SharePoint URL for one file is not one string: `?web=1`, `?d=w…` and a
 *   `#` anchor come and go between sessions, and a key that changed with them
 *   would remember nothing across the closes it is meant to survive.
 * - **It HASHES what is left.** Only equality is ever asked of this, so there
 *   is no reason to write somebody's SharePoint path — which can name a client,
 *   a project or a person — into storage where anything else on the origin
 *   could read it back. FNV-1a, 32 bits, in hex: not a security claim, and it
 *   is not being used as one; a collision costs two decks one shared memory of
 *   which categories were open.
 */
export function deckKey(url: string | undefined): string {
  const normal = normalize(url);
  return normal === "" ? GLOBAL_KEY : `${GLOBAL_KEY}:${fnv1a(normal)}`;
}

/** The part of a URL that names the same file every time it is asked for. */
function normalize(url: string | undefined): string {
  if (typeof url !== "string") return "";
  const cut = url.split("#")[0]?.split("?")[0] ?? "";
  return cut.trim().toLowerCase();
}

/** FNV-1a over the string's UTF-16 code units, as eight hex digits. */
function fnv1a(text: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    // The FNV prime. `Math.imul` because it is the exact 32-bit multiply the
    // algorithm specifies — a plain `*` leaves a double's exact-integer range
    // within a few characters and silently drops low bits. NOT a measured
    // difference here: `test/memory.test.ts` stays green with `*`, and says so.
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}
