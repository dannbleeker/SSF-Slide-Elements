import { describe, expect, it } from "vitest";
import { EMPTY, STEP_TITLE, STEPS, blockedReason, primary } from "../src/pane/steps.js";

/**
 * The pane's decisions, checked without a browser.
 *
 * Small on purpose: one step, one button, one reason. What matters is that
 * each is a VALUE the manual can be held against and the renderer cannot
 * invent — the docs guard reads these, and so does the screen.
 */
describe("the steps", () => {
  it("each have a title", () => {
    expect(STEPS.length).toBeGreaterThan(0);
    for (const id of STEPS) expect(STEP_TITLE[id].length, id).toBeGreaterThan(0);
  });

  it("start where the pane starts", () => {
    expect(STEPS[0]).toBe("start");
  });
});

describe("the one control", () => {
  it("is labelled, and cannot be pressed while there is nothing to insert", () => {
    const action = primary(EMPTY, "start");
    expect(action.label).toBe("Insert an element");
    expect(action.disabled).toBe(true);
  });

  it("comes with a reason, said rather than implied", () => {
    // A disabled button with no sentence beside it is a pane that looks broken.
    const why = blockedReason(EMPTY, "start");
    expect(why).toMatch(/nothing to insert/i);
    expect(why).toMatch(/later release/i);
  });
});
