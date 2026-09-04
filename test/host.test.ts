import { describe, expect, it } from "vitest";
import { canInsert, canReadSelection, checkFloor, environmentLine, FLOOR } from "../src/host/capability.js";
import { choosePlacement, fitScale, overlapFraction } from "../src/host/placement.js";
import { removalProven, sweepPlan } from "../src/host/undo.js";
import { readable } from "../src/host/errors.js";
import { withTimeout } from "../src/host/timeout.js";

const supports = (have: string[]) => (v: string) => have.includes(v);

describe("capability", () => {
  it("passes a host at the floor and explains a host below it", () => {
    expect(checkFloor(supports([FLOOR])).ok).toBe(true);
    const refused = checkFloor(supports([]));
    expect(refused.ok).toBe(false);
    // The sentence has to name what is missing: a user on an old host sees
    // this and nothing else, and "not supported" is not something they can act
    // on or report.
    expect(refused.detail).toContain("PowerPointApi");
  });

  it("answers the selection question separately from the floor", () => {
    expect(canReadSelection(supports(["1.2"]))).toBe(false);
    expect(canReadSelection(supports(["1.5"]))).toBe(true);
    expect(canInsert(supports(["1.2"]))).toBe(true);
  });

  it("fills every field of the environment line, even when nothing answered", () => {
    const env = environmentLine({ supports: supports([]) });
    // A missing field reads as a field nobody asked about, which is a
    // different bug from a host that would not answer.
    expect(env).toEqual({
      build: "unknown",
      platform: "unknown",
      host: "unknown",
      officeVersion: "unknown",
      api: "none",
    });
  });

  it("reports the highest api the host has, not the lowest", () => {
    expect(environmentLine({ supports: supports(["1.1", "1.2", "1.5"]) }).api).toBe("1.5");
  });
});

describe("placement", () => {
  const canvas = { cx: 12192000, cy: 6858000 };
  const box = { x: 1000000, y: 1000000, cx: 3000000, cy: 2000000 };

  it("leaves an element where it was authored when nothing is in the way", () => {
    const at = choosePlacement(box, canvas, []);
    expect(at).toMatchObject({ x: box.x, y: box.y, scale: 1, reason: "as-authored" });
  });

  it("moves clear of something already there", () => {
    const at = choosePlacement(box, canvas, [box]);
    expect(at.reason).toBe("nudged-clear");
    expect(overlapFraction({ ...box, x: at.x, y: at.y }, box)).toBeLessThan(0.2);
  });

  it("centres rather than overlapping when the slide is full", () => {
    const full = [{ x: 0, y: 0, cx: canvas.cx, cy: canvas.cy }];
    expect(choosePlacement(box, canvas, full).reason).toBe("centred");
  });

  /**
   * Uniform, never per-axis. A stretched element is worse than a small one,
   * because it is wrong in a way that looks deliberate.
   */
  it("scales uniformly to fit a smaller canvas", () => {
    const wide = { x: 0, y: 0, cx: 12000000, cy: 3000000 };
    const small = { cx: 6000000, cy: 6000000 };
    const k = fitScale(wide, small, 0);
    expect(k).toBeCloseTo(0.5, 5);
  });

  it("does not scale something that already fits", () => {
    expect(fitScale(box, canvas, 0)).toBe(1);
  });

  it("reports no overlap for boxes that only touch", () => {
    expect(overlapFraction({ x: 0, y: 0, cx: 10, cy: 10 }, { x: 10, y: 0, cx: 10, cy: 10 })).toBe(0);
  });
});

describe("undo", () => {
  it("removes the slide the insert added", () => {
    expect(sweepPlan(10, 11, 3).positions).toEqual([4]);
  });

  /**
   * A deck that gained a different number of slides than the insert asked for
   * has had something else happen to it. Removing by position then removes
   * somebody's work, so this refuses and says why.
   */
  it("refuses when the deck gained something unexpected", () => {
    expect(sweepPlan(10, 13, 3).positions).toEqual([]);
    expect(sweepPlan(10, 13, 3).why).toContain("3 slides");
  });

  it("refuses when nothing was added", () => {
    expect(sweepPlan(10, 10, 3).positions).toEqual([]);
  });

  it("treats a delete that changed no count as not having happened", () => {
    expect(removalProven(10, 10, 1)).toBe(false);
    expect(removalProven(10, 9, 1)).toBe(true);
  });
});

describe("errors", () => {
  /**
   * Office echoes the failing argument back, and the argument is an entire
   * presentation as base64. Uncapped, a failed insert puts megabytes of the
   * user's own deck on screen as the failure sentence.
   */
  it("caps a host error that carries the whole deck back with it", () => {
    const huge = `refused: ${"A".repeat(200000)}`;
    const out = readable(new Error(huge));
    expect(out.length).toBeLessThan(420);
    expect(out.startsWith("refused:")).toBe(true);
  });

  it("says something rather than nothing for an error with no message", () => {
    expect(readable({})).toContain("PowerPoint");
    expect(readable(new Error("  "))).toContain("PowerPoint");
  });
});

describe("timeout", () => {
  it("names what was waiting when it gives up", async () => {
    await expect(withTimeout(new Promise(() => {}), 10, "reading the deck")).rejects.toThrow(/reading the deck/);
  });

  it("passes a value straight through", async () => {
    await expect(withTimeout(Promise.resolve(7), 1000, "x")).resolves.toBe(7);
  });
});
