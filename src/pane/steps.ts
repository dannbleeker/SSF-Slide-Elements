/**
 * What the pane decides, kept away from what it draws.
 *
 * `render.ts` builds DOM from the answers here and `main.ts` wires the host in;
 * neither decides anything of its own, and `test/architecture.test.ts` holds
 * that seam. Everything in this file is an ordinary function over plain values,
 * which is what lets the suite check the pane's behaviour with no PowerPoint
 * and no browser.
 *
 * One step today. The picker — browse, choose, insert — grows `StepId` and
 * `STEP_TITLE` in the change that builds it, and `test/docs.test.ts` fails
 * until the manual names every title here.
 */

export type StepId = "start";

export const STEPS: readonly StepId[] = ["start"];

/** The heading each step draws. The manual must quote every one of these. */
export const STEP_TITLE: Record<StepId, string> = {
  start: "Start here",
};

export interface PaneState {
  /** What the host or the pane has to say, shown under the reason. */
  notice?: string;
}

export const EMPTY: PaneState = {};

/**
 * Why the one button cannot be pressed.
 *
 * Stated rather than implied by a greyed-out control: a disabled button with
 * no sentence beside it is a pane that looks broken.
 */
export function blockedReason(_state: PaneState, _step: StepId): string {
  return (
    "There is nothing to insert yet. This build is the hosting, the build stamp and the version check; " +
    "the element library arrives in a later release."
  );
}

/** The one primary control per screen: what it says, and whether it can be pressed. */
export function primary(_state: PaneState, _step: StepId): { label: string; disabled: boolean } {
  return { label: "Insert an element", disabled: true };
}
