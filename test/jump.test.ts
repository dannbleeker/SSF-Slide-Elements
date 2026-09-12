import { describe, expect, it } from "vitest";
import { jumpOutcome, sameSlideId } from "../src/host/jump.js";

/**
 * The jump from "Used in this deck", decided over what the host answered.
 *
 * The call behind it is made on a sibling's measurement, not on this repo's
 * (`src/host/jump.ts` says which and when), so the one thing this file has to
 * hold is that the pane never CLAIMS a jump it did not see: only a read-back
 * naming the slide asked for may produce "Slide N".
 */
describe("whether two slide ids name the same slide", () => {
  it("matches the same id, and a selection id that lost its suffix", () => {
    // office-js#2474: a SlideRange id lacks the `#suffix` the deck's own list
    // carries, so the half before the `#` is the comparison when one side has none.
    expect(sameSlideId("257#1234", "257#1234")).toBe(true);
    expect(sameSlideId("257", "257#1234")).toBe(true);
    expect(sameSlideId("257#1234", "257")).toBe(true);
  });

  it("refuses different slides, suffixed or not", () => {
    expect(sameSlideId("257#1234", "258#1234")).toBe(false);
    expect(sameSlideId("257#1234", "257#9999")).toBe(false);
    expect(sameSlideId("257", "258")).toBe(false);
    expect(sameSlideId("", "")).toBe(true);
    expect(sameSlideId("", "#1")).toBe(false);
  });
});

describe("what a jump's read-back means", () => {
  const asked = { slide: 4, wanted: "260#77", supported: true };

  it("claims the jump only when the read-back names the slide", () => {
    expect(jumpOutcome({ ...asked, selected: ["260#77"] })).toEqual({ ok: true, detail: "Slide 4" });
    // A desktop host may answer the id without its suffix (office-js#2474).
    expect(jumpOutcome({ ...asked, selected: ["260", "261#3"] }).ok).toBe(true);
  });

  it("says the host did not move when the read-back names another slide, or nothing", () => {
    const other = jumpOutcome({ ...asked, selected: ["259#1"] });
    expect(other.ok).toBe(false);
    expect(other.detail).toBe("PowerPoint did not move to slide 4. Click slide 4 in the strip.");
    const none = jumpOutcome({ ...asked, selected: [] });
    expect(none.ok).toBe(false);
    expect(none.detail).toContain("did not move");
  });

  it("does not assume a move from a host that did not answer", () => {
    // The wedge the family measured for setSelectedShapes shows as the write
    // being taken and the read after it going silent. Silence is not a move.
    const silent = jumpOutcome({ ...asked, selected: null });
    expect(silent.ok).toBe(false);
    expect(silent.detail).toBe("PowerPoint did not say which slide it is on. Click slide 4 in the strip.");
  });

  it("names the reason when the call raised", () => {
    // office-js#3552: desktop throws while the notes pane has focus.
    const threw = jumpOutcome({ ...asked, selected: null, error: "GeneralException" });
    expect(threw.ok).toBe(false);
    expect(threw.detail).toBe("Could not go to slide 4: GeneralException. Click slide 4 in the strip.");
  });

  it("says what the host lacks below PowerPointApi 1.5", () => {
    const old = jumpOutcome({ ...asked, supported: false, selected: null });
    expect(old.ok).toBe(false);
    expect(old.detail).toContain("PowerPointApi 1.5");
    expect(old.detail).toContain("Click slide 4 in the strip.");
  });
});
