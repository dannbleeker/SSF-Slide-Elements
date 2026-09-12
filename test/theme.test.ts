import { describe, expect, it } from "vitest";
import { paneTheme } from "../src/host/theme.js";

/**
 * Which theme the pane draws itself in, from the colour PowerPoint gives.
 *
 * This arithmetic decided what every user of the add-in sees, and until
 * 2026-09-12 nothing measured it: it was five lines inside `applyTheme` in
 * `src/pane/main.ts`, the one file `scripts/coverage-scope.mjs` deliberately
 * exempts. The exemption is right about that file's job — Office.js calls and
 * DOM plumbing — and was wrong about its contents.
 *
 * The threshold, the three-digit form and the give-up-and-say-nothing case are
 * each held here, and so is the one behaviour that CHANGED in the move: a
 * string that is not a colour now answers undefined instead of a brightness
 * derived from digits nobody wrote.
 */

describe("light or dark, from the host's own chrome", () => {
  it("calls PowerPoint's dark chrome dark and its light chrome light", () => {
    // The two colours the host actually sends: near-black and near-white.
    expect(paneTheme("#1f1f1f")).toBe("dark");
    expect(paneTheme("#ffffff")).toBe("light");
  });

  it("splits at half of 255, and a mid-grey falls on the dark side", () => {
    // The boundary pair, MEASURED on 2026-09-12 rather than reasoned about —
    // and the first version of this case had it the other way round.
    //
    // A grey's luminance ought to be its own channel value, because the
    // coefficients are a weighted mean. They are not quite: 0.299 + 0.587 +
    // 0.114 is 0.9999999999999999 in IEEE-754, so a grey of 128 computes to
    // 127.99999999999999 and `< 128` calls it dark. 129 is the first grey this
    // calls light.
    //
    // Pinned because it is invisible. The obvious tidy — factoring the
    // coefficients, or rounding before comparing — flips the theme PowerPoint
    // gets for a mid-grey chrome and nothing else in the suite would notice.
    expect(paneTheme("#808080")).toBe("dark");
    expect(paneTheme("#818181")).toBe("light");
    // And well clear of the boundary on both sides, so the case still says
    // something if the float ever stops behaving this way.
    expect(paneTheme("#707070")).toBe("dark");
    expect(paneTheme("#909090")).toBe("light");
  });

  it("weighs green far above blue, which is the point of using luminance at all", () => {
    // Full green and full blue are the same distance from black by any
    // channel-blind measure, and nothing like it to a person. A mean of the
    // three channels would call both of these dark.
    expect(paneTheme("#00ff00")).toBe("light");
    expect(paneTheme("#0000ff")).toBe("dark");
  });

  it("reads the three-digit form the same as the six", () => {
    expect(paneTheme("#000")).toBe("dark");
    expect(paneTheme("#fff")).toBe("light");
    // The expansion is per digit, not a pad: `#abc` is `#aabbcc`, not `#abc000`.
    expect(paneTheme("#abc")).toBe(paneTheme("#aabbcc"));
  });

  it("does not mind the case, the whitespace, or a missing hash", () => {
    expect(paneTheme("#FFFFFF")).toBe("light");
    expect(paneTheme("  #1f1f1f  ")).toBe("dark");
    expect(paneTheme("ffffff")).toBe("light");
  });
});

describe("what it refuses to guess at", () => {
  /**
   * Undefined is a real answer here, not a failure.
   *
   * `taskpane.css` carries a `prefers-color-scheme` fallback, and outside a
   * host — which is every time this pane is opened in a browser to look at it —
   * that fallback is the RIGHT answer. Stamping `data-theme` anyway overrides
   * it with a worse one, so everything this cannot read leaves the attribute
   * unset.
   */

  it("says nothing when the host says nothing", () => {
    expect(paneTheme(undefined)).toBeUndefined();
    expect(paneTheme("")).toBeUndefined();
    expect(paneTheme("   ")).toBeUndefined();
  });

  it("says nothing for a string that is not a colour", () => {
    expect(paneTheme("nonsense")).toBeUndefined();
    expect(paneTheme("rgb(31,31,31)")).toBeUndefined();
    expect(paneTheme("#")).toBeUndefined();
  });

  it("says nothing for a hex string of a length no colour has", () => {
    // The behaviour that CHANGED on 2026-09-12. The old inline version handed
    // whatever it had to `parseInt`, which reads as far as it understands: five
    // digits parsed to 0x12345 and came back "dark", a theme derived from a
    // number the host never sent. Four and five digits are the shapes a
    // truncated or mistyped colour actually takes.
    expect(paneTheme("#12345")).toBeUndefined();
    expect(paneTheme("#abcd")).toBeUndefined();
    expect(paneTheme("#1234567")).toBeUndefined();
  });

  it("says nothing for six characters that are not all hex digits", () => {
    // Six long and still not a colour: the length check alone would let this
    // through to `parseInt`, which stops at the `g` and answers 0xff — black,
    // and so "dark", from a string that names nothing.
    expect(paneTheme("#ffgghh")).toBeUndefined();
  });

  it("says nothing for a value that is not a string at all", () => {
    // `Office.context.officeTheme.bodyBackgroundColor` is typed as a string and
    // the host sends one. The guard is here because the typings describe what
    // the host is supposed to do, and `CLAUDE.md` is a list of the times it did
    // not — and because a `.trim()` on a number is a raise, which reaches
    // `Office.onReady` and leaves the pane blank rather than merely untinted.
    // Cast, because the point is the value a typed caller cannot send.
    expect(paneTheme(42 as unknown as string)).toBeUndefined();
    expect(paneTheme(null as unknown as string)).toBeUndefined();
  });
});
