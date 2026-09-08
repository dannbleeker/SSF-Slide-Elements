import { describe, expect, it } from "vitest";
import { API_FLOOR, checkFloor, type Supports } from "../src/host/capability.js";

/** A host that supports every set up to and including `top`. */
function upTo(top: string): Supports {
  const [major = 0, minor = 0] = top.split(".").map(Number);
  return (version) => {
    const [a = 0, b = 0] = version.split(".").map(Number);
    return a < major || (a === major && b <= minor);
  };
}

describe("the floor", () => {
  it("is PowerPointApi 1.2", () => {
    // Read off the calls the insert will make; see the table in capability.ts.
    expect(API_FLOOR).toBe("1.2");
  });

  it("passes a host at the floor", () => {
    const answer = checkFloor(upTo("1.2"));
    expect(answer.ok).toBe(true);
    expect(answer.detail).toContain(API_FLOOR);
  });

  it("refuses a host below it, and says what that costs", () => {
    const answer = checkFloor(upTo("1.1"));
    expect(answer.ok).toBe(false);
    expect(answer.detail).toContain(API_FLOOR);
    // Which version is missing, and what the user loses: the sentence a
    // declared requirement set would have replaced with silence.
    expect(answer.detail).toMatch(/insert|deck|read/i);
    expect(answer.detail).toContain("SSF Slide Elements");
  });

  it("passes every host above it", () => {
    for (const top of ["1.3", "1.5", "1.8", "1.10"]) expect(checkFloor(upTo(top)).ok, top).toBe(true);
  });

  it("treats a host that raises as one that cannot", () => {
    // `hostSupports` in src/office turns a raise into false; the decision
    // here only ever sees a boolean, and this pins that a false is a refusal
    // with a sentence rather than a blank.
    const answer = checkFloor(() => false);
    expect(answer.ok).toBe(false);
    expect(answer.detail.length).toBeGreaterThan(20);
  });
});
