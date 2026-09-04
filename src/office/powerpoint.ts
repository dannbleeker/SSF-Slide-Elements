/**
 * The Office.js calls, and nothing else.
 *
 * Every judgement is imported from `src/host`, where it is a pure function the
 * suite can check. Nothing here decides anything, because a decision inside a
 * `PowerPoint.run` callback is a decision nobody can test, and the host this
 * runs on is documented — by a sibling project, at length — to lie about ids,
 * to accept calls it does not perform, and to answer differently on two runs of
 * the same build.
 *
 * The insert is deliberately built out of the two calls that sibling has already
 * proven against a real PowerPoint:
 *
 * - `getFileAsync(Compressed)` to read the deck. A Common API, present on every
 *   host that clears the floor, and read back successfully on PowerPoint for
 *   the web.
 * - `insertSlidesFromBase64` WITH a `targetSlideId`. Without one the host
 *   inserts at the FRONT — a sibling put 37 generated slides ahead of somebody's
 *   title slide that way.
 *
 * What it does NOT use is `sourceSlideIds`, whose id format this project has
 * not measured. The package handed over is pruned to the one spliced slide
 * instead, which is work done in the file where it can be checked.
 */
import { BUDGET, withTimeout } from "../host/timeout.js";
import { readable } from "../host/errors.js";
import { canReadSelection, checkFloor, environmentLine, type Environment, type Supports } from "../host/capability.js";
import { removalProven, sweepPlan } from "../host/undo.js";
import type { Box } from "../core/catalogue/types.js";

/**
 * What the host says it supports.
 *
 * Answers false rather than throwing when it cannot ask: `Office.context` being
 * absent is not a supported host reporting an absent set, it is a host that
 * cannot answer — and every caller wants the same thing from both. The guard is
 * at the root because this is called before the pane has decided anything, and
 * a raise here leaves the user with a blank pane on precisely the host that
 * needed the explanation.
 */
export const hostSupports: Supports = (version) => {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
};

export function ready(): { ok: boolean; detail: string } {
  return checkFloor(hostSupports);
}

export function hostEnvironment(): Environment {
  const read = (f: () => string | undefined): string | undefined => {
    try {
      return f();
    } catch {
      return undefined;
    }
  };
  return environmentLine({
    ...(typeof __BUILD_STAMP__ === "string" ? { build: __BUILD_STAMP__ } : {}),
    ...((p) => (p ? { platform: p } : {}))(
      read(() => (Office.context.platform === undefined ? undefined : String(Office.context.platform))),
    ),
    // WHICH HOST, from the field that names it — stringified because the type
    // is an enum. This was filled from `diagnostics.version` in the sibling for
    // a while, so the one field answering "which application am I in" carried a
    // build number and the question went unasked.
    ...((h) => (h ? { host: h } : {}))(
      read(() => {
        const named = Office.context.diagnostics?.host;
        return named === undefined ? undefined : String(named);
      }),
    ),
    ...((v) => (v ? { officeVersion: v } : {}))(read(() => Office.context.diagnostics?.version)),
    supports: hostSupports,
  });
}

export async function slideCount(): Promise<number> {
  return withTimeout(
    PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const count = slides.getCount();
      await context.sync();
      return count.value;
    }),
    BUDGET.query,
    "counting the slides",
  );
}

/**
 * Which slide the user is looking at, as a ZERO-BASED position.
 *
 * A position, not an id, and that is the sibling's hardest-won rule: this
 * host's slide ids look like `256#3561048925`, both a position and an id are
 * strings, and nothing fails until PowerPoint does. The splice addresses a
 * slide by its place in `<p:sldIdLst>`, which is the same ordering the user
 * sees, so a position is what the engine actually wants.
 *
 * Falls back to the FIRST slide rather than throwing when the host will not
 * say. On a host below 1.5 there is no selection API at all, and refusing to
 * work is a worse answer than putting the element on slide one and saying so.
 */
export async function selectedSlideIndex(): Promise<{ index: number; certain: boolean }> {
  if (!canReadSelection(hostSupports)) return { index: 0, certain: false };
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const selected = context.presentation.getSelectedSlides();
        selected.load("items/id");
        const all = context.presentation.slides;
        all.load("items/id");
        await context.sync();
        const first = selected.items[0];
        if (!first) return { index: 0, certain: false };
        const at = all.items.findIndex((s) => s.id === first.id);
        return at < 0 ? { index: 0, certain: false } : { index: at, certain: true };
      }),
      BUDGET.query,
      "asking which slide is selected",
    );
  } catch {
    return { index: 0, certain: false };
  }
}

/**
 * The whole package, slice by slice.
 *
 * Each slice is turned into a string in chunks, because
 * `String.fromCharCode.apply` on a whole slice overruns the argument limit on a
 * deck of any size.
 */
export function readDeck(): Promise<string> {
  return withTimeout(
    new Promise<string>((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(res.error ? res.error.message : "PowerPoint would not open the file"));
          return;
        }
        const file = res.value;
        const chunks: string[] = [];
        const next = (i: number): void => {
          if (i >= file.sliceCount) {
            file.closeAsync();
            resolve(btoa(chunks.join("")));
            return;
          }
          file.getSliceAsync(i, (slice) => {
            if (slice.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(slice.error ? slice.error.message : `PowerPoint refused slice ${i}`));
              return;
            }
            const bytes = slice.value.data as number[];
            let s = "";
            for (let j = 0; j < bytes.length; j += 0x8000) {
              s += String.fromCharCode.apply(null, bytes.slice(j, j + 0x8000));
            }
            chunks.push(s);
            next(i + 1);
          });
        };
        next(0);
      });
    }),
    BUDGET.file,
    "reading the presentation",
  );
}

/** What is already on a slide, so the placement can avoid it. */
export async function occupiedOn(index: number): Promise<Box[]> {
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(index);
        const shapes = slide.shapes;
        shapes.load("items/left,items/top,items/width,items/height");
        await context.sync();
        // Office.js reports these in POINTS; everything in the engine is EMU.
        return shapes.items.map((s) => ({
          x: Math.round(s.left * 12700),
          y: Math.round(s.top * 12700),
          cx: Math.round(s.width * 12700),
          cy: Math.round(s.height * 12700),
        }));
      }),
      BUDGET.query,
      "reading what is already on the slide",
    );
  } catch {
    // A slide whose shapes cannot be read is treated as empty. The cost of
    // being wrong is an element that overlaps and gets dragged; the cost of
    // refusing is an add-in that does nothing on a host that would have worked.
    return [];
  }
}

export interface InsertOutcome {
  ok: boolean;
  before: number;
  after: number;
  /** Which position the new slide should be at, for a later removal. */
  addedAt: number;
  detail: string;
}

/**
 * Put the spliced slide into the deck, after the slide it replaces.
 *
 * The deck delta is the evidence, never the absence of an error: an insert has
 * timed out having landed everything it was asked for, and a run that treats
 * the raise as the answer reports a failure the user can see did not happen.
 */
export async function insertSpliced(base64: string, afterIndex: number): Promise<InsertOutcome> {
  const before = await slideCount();
  let error: string | undefined;
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        const target = slides.getItemAt(afterIndex);
        target.load("id");
        await context.sync();
        // Without a target the host inserts at the FRONT.
        context.presentation.insertSlidesFromBase64(base64, {
          formatting: "KeepSourceFormatting",
          targetSlideId: target.id,
        });
        await context.sync();
      }),
      BUDGET.insert,
      "inserting the slide",
    );
  } catch (e) {
    // `readable`, not `e.message`: the argument Office echoes back is the whole
    // presentation.
    error = readable(e);
  }
  const after = await slideCount();
  const gained = after - before;
  if (gained === 1) {
    return {
      ok: true,
      before,
      after,
      addedAt: afterIndex + 1,
      detail: error ? `landed despite: ${error}` : "inserted",
    };
  }
  return {
    ok: false,
    before,
    after,
    addedAt: -1,
    detail: error ?? `the deck gained ${gained} slides where 1 was expected`,
  };
}

/** Remove one slide by position, and prove it went. */
export async function removeSlideAt(index: number): Promise<{ ok: boolean; detail: string }> {
  const before = await slideCount();
  if (index < 0 || index >= before) return { ok: false, detail: "that slide is no longer where it was" };
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        context.presentation.slides.getItemAt(index).delete();
        await context.sync();
      }),
      BUDGET.remove,
      "removing the replaced slide",
    );
  } catch (e) {
    return { ok: false, detail: readable(e) };
  }
  // A queued delete that raised nothing has not necessarily happened.
  const after = await slideCount();
  return removalProven(before, after, 1)
    ? { ok: true, detail: "removed" }
    : { ok: false, detail: "PowerPoint accepted the removal but the deck is the same size" };
}

/** Take back an insert, by the plan `src/host/undo.ts` decides. */
export async function undoInsert(before: number, afterIndex: number): Promise<{ ok: boolean; detail: string }> {
  const now = await slideCount();
  const plan = sweepPlan(before, now, afterIndex);
  if (plan.positions.length === 0) return { ok: false, detail: plan.why };
  const at = plan.positions[0];
  if (at === undefined) return { ok: false, detail: plan.why };
  return removeSlideAt(at);
}
