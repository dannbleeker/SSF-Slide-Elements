/**
 * What this host can do, decided away from the host.
 *
 * The question "is this version supported" is answered by Office.js, and
 * everything that follows from the answer is an ordinary function over strings
 * that the suite can check. A capability check written inline in a
 * `PowerPoint.run` callback is a check nobody can test and everybody has to
 * trust. This is the sibling projects' split, kept.
 */

/** Whether the host supports a given PowerPointApi version. */
export type Supports = (version: string) => boolean;

/**
 * The lowest PowerPointApi version this add-in can work on.
 *
 * Read off the calls the insert will make, against the office-js typings,
 * rather than picked:
 *
 * | call                                        | set |
 * | ------------------------------------------- | --- |
 * | `presentation.slides`                       | 1.2 |
 * | `slides.getCount` / `getItemAt`             | 1.2 |
 * | `slide.id`, `slide.delete`                  | 1.2 |
 * | `presentation.insertSlidesFromBase64`       | 1.2 |
 *
 * So the floor is **1.2**. `getFileAsync` is a Common API and is not gated by
 * PowerPointApi at all, which is what makes the deck readable on every host
 * that clears this bar.
 *
 * A second spelling lives in `scripts/manifest-source.mjs` as `FLOOR`, because
 * a build script cannot import TypeScript; `test/manifest.test.ts` holds the
 * two together. Change one and the test says so.
 */
export const API_FLOOR = "1.2";

export interface Readiness {
  ok: boolean;
  detail: string;
}

/**
 * Whether this host clears the floor.
 *
 * Checked at RUNTIME and never declared in the manifest. A declared requirement
 * set that the host does not meet makes the add-in vanish from the ribbon with
 * no diagnostic at all, so the user sees nothing and has nothing to report; a
 * runtime check can say which version is missing and what that costs them.
 */
export function checkFloor(supports: Supports): Readiness {
  if (supports(API_FLOOR)) {
    return { ok: true, detail: `this host supports PowerPointApi ${API_FLOOR}` };
  }
  return {
    ok: false,
    detail: `SSF Slide Elements needs PowerPointApi ${API_FLOOR} and this host does not have it. Reading the deck and inserting an element into it both need it; without it there is nothing the add-in can do.`,
  };
}
