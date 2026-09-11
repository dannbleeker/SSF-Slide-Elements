/**
 * The Office.js calls, and nothing else.
 *
 * Every judgement is imported from `src/host`, where it is a pure function the
 * suite can check. Nothing here decides anything, because a decision inside a
 * `PowerPoint.run` callback is a decision nobody can test — and the host this
 * will run on is documented, by both sibling projects and at length, to lie
 * about ids, to accept calls it does not perform, and to answer differently on
 * two runs of the same build. `CLAUDE.md` ("Host rules, learned the expensive way") is the ledger,
 * borrowed from the siblings.
 *
 * The insert is the two calls the siblings have proven against a real
 * PowerPoint, and that the probe measured again on the web on 2026-09-10:
 * `getFileAsync(Compressed)` to read the deck, and `insertSlidesFromBase64`
 * **with** a `targetSlideId`, because without one the host inserts at the FRONT
 * and a sibling put 37 generated slides ahead of somebody's title slide that
 * way.
 *
 * Nothing here retries. A retry inside the thing that measures whether an
 * insert happened is a second insert on a deck whose shape this code has
 * already misread, and the count it then reports is about neither attempt.
 */
import { checkFloor, type Readiness, type Supports } from "../host/capability.js";
import { readable } from "../host/errors.js";
import { BUDGET, withTimeout } from "../host/timeout.js";

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

/** Whether this host clears the floor, as a sentence the pane can show. */
export function ready(): Readiness {
  return checkFloor(hostSupports);
}

/** A rectangle in EMU, the unit the package uses. */
export interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/** Office measures a shape in POINTS; the file measures everything in EMU. */
const EMU_PER_POINT = 12700;

/** How many slides the deck holds, counted in its own batch. */
export function slideCount(): Promise<number> {
  return withTimeout(
    PowerPoint.run(async (context) => {
      const count = context.presentation.slides.getCount();
      await context.sync();
      return count.value;
    }),
    BUDGET.read,
    "counting the deck's slides",
  );
}

export interface DeckBytes {
  base64: string;
  bytes: number;
  ms: number;
}

/**
 * The user's whole presentation, as base64.
 *
 * `getFileAsync` is a Common API and is not gated by PowerPointApi at all,
 * which is what makes the deck readable on every host that clears the floor.
 * The other read, `exportAsBase64Presentation`, is deliberately NOT used: the
 * probe measured it dropping the deck's comment part and `ppt/authors.xml` on
 * the web on 2026-09-10, and this add-in sends back a package built from what
 * it read, so a dropped comment is a comment the user loses.
 *
 * The slices are chunked by hand because `apply()` over a whole slice blows the
 * argument limit on a large deck.
 */
export function readDeck(): Promise<DeckBytes> {
  const started = Date.now();
  return withTimeout(
    new Promise<DeckBytes>((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(res.error ? res.error.message : "getFileAsync would not open the file"));
          return;
        }
        const file = res.value;
        const chunks: string[] = [];
        let bytes = 0;
        const next = (i: number): void => {
          if (i >= file.sliceCount) {
            file.closeAsync();
            resolve({ base64: btoa(chunks.join("")), bytes, ms: Date.now() - started });
            return;
          }
          file.getSliceAsync(i, (slice) => {
            if (slice.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(slice.error ? slice.error.message : "getSliceAsync refused a slice"));
              return;
            }
            const data = slice.value.data as number[];
            bytes += data.length;
            let s = "";
            for (let j = 0; j < data.length; j += 0x8000) {
              s += String.fromCharCode.apply(null, data.slice(j, j + 0x8000));
            }
            chunks.push(s);
            next(i + 1);
          });
        };
        next(0);
      });
    }),
    BUDGET.deck,
    "reading this presentation",
  );
}

/** Which slide the user is on, and the id to insert after. */
export interface Current {
  /** Its position, counting from zero, which is what the splice takes. */
  index: number;
  /** Its id in the host's own `256#3561048925` spelling, for `targetSlideId`. */
  id: string;
}

/**
 * The slide the user is looking at.
 *
 * `getSelectedSlides` is PowerPointApi 1.5 and is READ-ONLY; nothing here ever
 * calls `setSelectedSlides` or `setSelectedShapes`, which wedge the web host's
 * selection subsystem (`CLAUDE.md`). The probe measured this on the web on
 * 2026-09-10: the selection named slide 2, and the API's order matched the
 * file's own `sldIdLst` order, which is what lets a selected id become an index
 * the splice can use.
 *
 * Answers undefined rather than guessing when the host has no selection read,
 * or names a slide this deck does not. The pane then asks the user which slide,
 * which is honest, where defaulting to the first would silently insert
 * somewhere they were not looking.
 */
export async function currentSlide(): Promise<Current | undefined> {
  if (!hostSupports("1.5")) return undefined;
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const selected = context.presentation.getSelectedSlides();
        // Named, never `load("items")` alone: that loads the collection and
        // none of the items' properties, and a read of it comes back empty.
        selected.load("items/id");
        const all = context.presentation.slides;
        all.load("items/id");
        await context.sync();
        const first = selected.items[0]?.id;
        if (first === undefined) return undefined;
        const index = all.items.findIndex((s) => s.id === first);
        return index < 0 ? undefined : { index, id: first };
      }),
      BUDGET.read,
      "asking which slide is selected",
    );
  } catch {
    // A selection read that fails is not something the user can act on, and the
    // pane has a fallback for not knowing. Swallowed here rather than raised.
    return undefined;
  }
}

/**
 * The selected shape's rectangle, for an element that lands "at the cursor".
 *
 * An add-in cannot see the mouse on the canvas; `docs/DESIGN.md` section 5 says
 * the selected shape stands in for it. Read-only, and undefined whenever the
 * host will not say — a marker then lands in the middle of the slide, which is
 * the documented fallback rather than a failure.
 */
export async function selectedShape(): Promise<Rect | undefined> {
  if (!hostSupports("1.5")) return undefined;
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/left,items/top,items/width,items/height");
        await context.sync();
        const shape = shapes.items[0];
        if (!shape) return undefined;
        return {
          x: Math.round(shape.left * EMU_PER_POINT),
          y: Math.round(shape.top * EMU_PER_POINT),
          cx: Math.round(shape.width * EMU_PER_POINT),
          cy: Math.round(shape.height * EMU_PER_POINT),
        };
      }),
      BUDGET.read,
      "asking which shape is selected",
    );
  } catch {
    return undefined;
  }
}

/**
 * The id of the slide at a POSITION, for an insert that has to aim at one.
 *
 * Read positionally and used immediately. Undo has to put a slide back next to
 * the rebuilt one rather than next to whatever the user has selected by then,
 * and `insertSlidesFromBase64` takes an id — so the index this code computed is
 * turned into an id here, at the last possible moment.
 *
 * A slide the run added IS accepted as a `targetSlideId`: the probe asked that
 * question directly on the web on 2026-09-10 and the answer is in
 * `docs/host-answers/`. What is never done is the other direction —
 * `slides.getItem(id)` on such a slide — and nothing here does it.
 */
export async function slideIdAt(index: number): Promise<string | undefined> {
  try {
    return await withTimeout(
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();
        return slides.items[index]?.id;
      }),
      BUDGET.read,
      "reading the id of the slide to insert against",
    );
  } catch {
    return undefined;
  }
}

/**
 * Hand the package to PowerPoint, after the slide the user is on.
 *
 * Returns the reason it raised, if it did, rather than throwing: a raise is not
 * the answer to whether the insert happened, and the caller settles that by
 * counting. `readable` bounds the reason, because Office echoes an argument
 * back through `debugInfo` and the argument here is an entire deck as base64 —
 * uncapped, a failed insert puts megabytes of the user's own presentation on
 * screen as the failure sentence.
 */
export async function insertPackage(base64: string, targetSlideId: string): Promise<string | undefined> {
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        context.presentation.insertSlidesFromBase64(base64, {
          formatting: "KeepSourceFormatting",
          targetSlideId,
        });
        await context.sync();
      }),
      BUDGET.insert,
      "inserting the rebuilt slide",
    );
    return undefined;
  } catch (e) {
    return readable(e);
  }
}

/**
 * Take one slide out, by POSITION.
 *
 * Never by id. A slide the run has just added does not round-trip through
 * `slides.getItem(id)` on the web — a sibling logged 46 refusals in one run —
 * and the slide being removed here sits next to one that was just added, so an
 * id read anywhere in this operation is suspect. The index is one the caller
 * computed and can defend, and the caller has already confirmed the deck is the
 * size that index was computed against.
 */
export async function removeSlideAt(index: number): Promise<string | undefined> {
  try {
    await withTimeout(
      PowerPoint.run(async (context) => {
        context.presentation.slides.getItemAt(index).delete();
        await context.sync();
      }),
      BUDGET.remove,
      "removing the slide the copy replaced",
    );
    return undefined;
  } catch (e) {
    return readable(e);
  }
}
