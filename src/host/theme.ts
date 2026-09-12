/**
 * Which theme the pane should draw itself in, given the host's own colour.
 *
 * `docs/DESIGN.md` section 9: the pane follows POWERPOINT'S theme, not the
 * browser's. PowerPoint can be dark while the OS is light, so
 * `prefers-color-scheme` answers the wrong question; `Office.context.officeTheme`
 * answers the right one, and what it hands over is a colour rather than a name.
 * Turning that colour into "dark" or "light" is the decision, and it is the
 * whole of this file.
 *
 * Pure on purpose: `src/office/powerpoint.ts` asks the host for the colour and
 * this decides what it means, so every branch below is one the suite can hold
 * without a PowerPoint. Until 2026-09-12 the parse, the threshold and the read
 * were three lines inside `applyTheme` in `src/pane/main.ts`, which is the one
 * file the coverage floor exempts — so the arithmetic that decides what every
 * user sees was the arithmetic nothing measured.
 */

/** What the pane stamps on `<html>` as `data-theme`. */
export type PaneTheme = "dark" | "light";

/**
 * The theme for a host background colour, or undefined to leave it to the CSS.
 *
 * Undefined is a real answer and not a failure: outside a host there is no
 * `officeTheme` at all — which is every time this pane is opened in a browser
 * to look at it — and `taskpane.css` carries a `prefers-color-scheme` fallback
 * for exactly that case. Stamping a guess would override the fallback with a
 * worse answer, so an unreadable colour leaves the attribute unset rather than
 * defaulting to either side.
 */
export function paneTheme(background: string | undefined): PaneTheme | undefined {
  const value = luminance(background);
  if (value === undefined) return undefined;
  // Half of 255, and darker than that is dark.
  //
  // A mid-grey `#808080` comes out DARK rather than landing on the boundary,
  // and the reason is arithmetic rather than intent: the BT.601 coefficients
  // sum to 0.9999999999999999 in IEEE-754, so a grey of 128 computes to
  // 127.99999999999999. Measured 2026-09-12; `#818181` is the first grey this
  // calls light. Carried over unchanged from the inline version this replaced,
  // and pinned in `test/theme.test.ts` so a later tidy of the coefficients
  // cannot flip a host's theme without a case going red.
  return value < 128 ? "dark" : "light";
}

/**
 * Perceived brightness of an `#rrggbb` or `#rgb` colour, 0..255.
 *
 * ITU-R BT.601's coefficients, which is what a task pane wants here rather than
 * the sRGB-linear version: the question is "does this chrome read as dark to a
 * person", not "what is this colour's physical luminance". Both put green far
 * above blue, and the threshold above was chosen against these weights.
 *
 * Undefined when the string is not a colour this can read. The host is supposed
 * to hand over six hex digits, and does; the three-digit form and the missing
 * `#` are accepted because they cost one line each and a theme that silently
 * gave up would look exactly like a host that has no theme.
 */
function luminance(background: string | undefined): number | undefined {
  if (typeof background !== "string") return undefined;
  const hex = background.trim().replace(/^#/, "");
  // Only these two lengths. `parseInt` reads as far as it understands and
  // answers a number for "12345", which would be a brightness derived from
  // digits the caller never wrote.
  if (hex.length !== 3 && hex.length !== 6) return undefined;
  if (!/^[0-9a-f]+$/i.test(hex)) return undefined;
  const full = hex.length === 3 ? [...hex].map((c) => c + c).join("") : hex;
  // No `Number.isNaN` guard after this, deliberately. The inline version this
  // replaced needed one, because it handed `parseInt` whatever it had; here the
  // two checks above leave exactly six hex digits, so the parse cannot fail and
  // a guard against it would be a line nothing could ever make go red.
  const n = Number.parseInt(full, 16);
  return ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
}
