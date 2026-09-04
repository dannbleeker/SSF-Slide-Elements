/**
 * What this PowerPoint can do, decided as pure functions.
 *
 * Everything here is a decision the suite can check; `src/office` makes the
 * calls and imports every judgement from this file. The split is SSF-Merge's
 * and it is enforced by `test/architecture.test.ts` in both directions: an
 * Office.js import here makes a rule untestable, and a rule reimplemented
 * inline over there looks tidier and rots quietly.
 */

/** `Office.context.requirements.isSetSupported`, as a function that cannot throw. */
export type Supports = (version: string) => boolean;

/**
 * The floor, and why it is checked at runtime rather than declared.
 *
 * A `<Requirements>` element the host does not meet makes the add-in vanish
 * from the ribbon with no diagnostic at all — the user sees nothing and has
 * nothing to report. Checked here, they get a sentence naming what is missing.
 *
 * **1.2 is the floor and 1.5 is what the insert wants.** Reading the deck's
 * bytes goes through `getFileAsync`, which is a Common API and predates all of
 * this; putting slides back needs `insertSlidesFromBase64`, which the
 * documentation places at 1.2 in one page and 1.5 in another. The
 * disagreement is why this asks for the CAPABILITY at the point of use rather
 * than trusting either number.
 */
export const FLOOR = "1.2";

export function checkFloor(supports: Supports): { ok: boolean; detail: string } {
  if (supports(FLOOR)) return { ok: true, detail: `PowerPointApi ${FLOOR} or better` };
  return {
    ok: false,
    detail:
      `This PowerPoint is missing PowerPointApi ${FLOOR}, which SSF Slide Elements needs to read your deck and put slides back into it. ` +
      "PowerPoint on the web, Microsoft 365 on Windows and Mac, and PowerPoint 2021 or newer all have it.",
  };
}

/** Whether the host can take a package back. Asked at the point of use. */
export function canInsert(supports: Supports): boolean {
  return supports("1.2") || supports("1.5");
}

/** Whether the host can say which slide the user is looking at. */
export function canReadSelection(supports: Supports): boolean {
  return supports("1.5");
}

export interface Environment {
  build: string;
  platform: string;
  host: string;
  officeVersion: string;
  api: string;
}

/**
 * One line describing what this is running on.
 *
 * Every field falls back to "unknown" rather than being omitted. A round is
 * judged from this line, and a missing field reads as a field that was not
 * asked about — which is a different bug from a host that would not answer.
 */
export function environmentLine(input: {
  build?: string;
  platform?: string;
  host?: string;
  officeVersion?: string;
  supports: Supports;
}): Environment {
  // Highest first: the answer wanted is "what does this host have", and asking
  // upwards means the loop stops at the first yes.
  const versions = ["1.8", "1.7", "1.6", "1.5", "1.4", "1.3", "1.2", "1.1"];
  const api = versions.find((v) => input.supports(v)) ?? "none";
  return {
    build: input.build || "unknown",
    platform: input.platform || "unknown",
    host: input.host || "unknown",
    officeVersion: input.officeVersion || "unknown",
    api,
  };
}
