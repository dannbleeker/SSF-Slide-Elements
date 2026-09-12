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
import { BUDGET, CONFIRM, withTimeout } from "../host/timeout.js";

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

/**
 * The deck's size, asked again until it is the size it ought to be.
 *
 * `slideCount` above answers what the host says NOW, and the host can be
 * behind: `src/host/timeout.ts` carries the measurement and what it cost.
 * This is the one to use wherever a count is EVIDENCE that a call did
 * something — after an insert, after a removal — and `slideCount` the one to
 * use for a size nothing turns on.
 *
 * It stops at the first read that matches, so the ordinary case costs exactly
 * what it always did. It returns whatever it last saw rather than raising:
 * whether that number means success is a judgement, and judgements live in
 * `src/host`, not here. A deck that ends up the wrong size on purpose — two
 * slides where one was expected — comes back as that wrong size, and the
 * sentence for it is already written.
 */
export async function countReaching(want: number): Promise<number> {
  let count = await slideCount();
  for (const pause of CONFIRM) {
    if (count === want) return count;
    await new Promise((resolve) => setTimeout(resolve, pause));
    count = await slideCount();
  }
  return count;
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
 * Whether a collection read came back WHOLE, judged against the deck's own
 * scalar count.
 *
 * A collection load over about fifty items answers short on the web
 * (office-js#4272), and `docs/SIBLING.md` triaged that as relevant here months
 * before there was any code to be relevant to. The row promised the sibling's
 * defence — page `getItemAt` at twenty — and the code that shipped does three
 * plain collection loads and no paging at all. This is the defence that
 * replaces it, and the reason it is a check rather than a paging loop: paging
 * is more host calls on a host that also fails to load properties reliably
 * after a `context.sync()` (office-js#6363), while the count is one scalar in
 * the batch that is already being sent.
 *
 * What a short read costs here is not a shorter list. Both callers turn the
 * list into an INDEX — which slide the user is on, which slide to aim an insert
 * at — and an index off a list that dropped something in the middle names a
 * different slide. The removal that follows an insert is positional. So a list
 * that does not match the count is refused outright, and both callers already
 * have somewhere honest to go when the host will not say: the pane tells the
 * user it does not know which slide they are on, and the undo says PowerPoint
 * would not name the slide.
 *
 * Never the other way round: a count that disagrees is not evidence the deck
 * changed, and nothing here concludes anything about the deck from a read.
 * `CLAUDE.md`: an empty collection read is not an empty slide.
 */
function whole(read: number, counted: number): boolean {
  return read === counted;
}

/**
 * The slide the user is looking at.
 *
 * `getSelectedSlides` is PowerPointApi 1.5 and is READ-ONLY. Nothing here ever
 * calls `setSelectedShapes`, which wedges the web host's selection subsystem
 * (`CLAUDE.md`); `setSelectedSlides` is called by `selectSlide` alone, below,
 * on the sibling's dated evidence. The probe measured this read on the web on
 * 2026-09-10: the selection named slide 2, and the API's order matched the
 * file's own `sldIdLst` order, which is what lets a selected id become an index
 * the splice can use.
 *
 * Answers undefined rather than guessing when the host has no selection read,
 * or names a slide this deck does not. The pane then asks the user which slide,
 * which is honest, where defaulting to the first would silently insert
 * somewhere they were not looking.
 */
export async function currentSlide(within: number = BUDGET.read): Promise<Current | undefined | null> {
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
        // The scalar count, in the SAME batch, because the list cannot be
        // trusted to be whole. See `whole` above.
        const count = all.getCount();
        await context.sync();
        const first = selected.items[0]?.id;
        if (first === undefined) return undefined;
        if (!whole(all.items.length, count.value)) return undefined;
        const index = all.items.findIndex((s) => s.id === first);
        return index < 0 ? undefined : { index, id: first };
      }),
      within,
      "asking which slide is selected",
    );
  } catch {
    // `null`, and the difference from `undefined` is the whole point:
    // **undefined is the host SAYING there is no selection, null is the host
    // not answering.** They were one value, and the caller that only wants to
    // keep a line fresh has to tell them apart — replacing a good slide number
    // with "PowerPoint did not say" because a read ran out of time is telling
    // the user something the host never said. Swallowed rather than raised,
    // because a selection read that fails is not something a user can act on.
    return null;
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
 * Tell me when the user moves, so the pane can stop naming the wrong slide.
 *
 * The pane reads which slide you are on once, at boot. Without this the line
 * under the header keeps naming that slide however many others you click, and
 * on 2026-09-11 in PowerPoint for the web it did exactly that: the API said
 * slide 2 was selected and the pane still read "Slide 1.". The insert has
 * always read the selection again for itself, so nothing landed in the wrong
 * place — but the pane was telling the user something that was not true about
 * where it was about to land, which is worse than saying nothing.
 *
 * `DocumentSelectionChanged` is a Common API event rather than a
 * `PowerPointApi` one, so `hostSupports` cannot answer for it. `CLAUDE.md`:
 * **probe for the method, do not trust the requirement set.** This asks whether
 * the function is there, subscribes, and answers whether the host accepted —
 * false on a host without it, where the pane simply keeps the line it has.
 *
 * The handler is called on every selection change, shapes included, so what it
 * does with that has to be cheap. That is the caller's problem, not this
 * file's: nothing here decides anything.
 */
export async function onSlideChange(handler: () => void): Promise<boolean> {
  const doc: Office.Document | undefined = Office.context?.document;
  if (typeof doc?.addHandlerAsync !== "function") return false;
  try {
    return await new Promise<boolean>((resolve) => {
      doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, (result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded);
      });
    });
  } catch {
    return false;
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
        const count = slides.getCount();
        await context.sync();
        if (!whole(slides.items.length, count.value)) return undefined;
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

/**
 * What the host says it is, for a problem report.
 *
 * `Office.context.diagnostics` is the modern answer and carries the host, the
 * platform and the Office version; the two older properties are the fallback,
 * because `diagnostics` needs requirement set CommonApi 1.1 and this add-in's
 * floor is a PowerPointApi one, not a Common one. Everything is optional and
 * nothing raises: a report missing the platform is still a report, and a pane
 * that threw while building a support link would be a pane that cannot ask for
 * help.
 *
 * The VALUES are not judged here. `src/host/links.ts` decides which of them may
 * reach a URL, because that is the rule worth testing.
 */
export function hostStamp(): { host?: string; platform?: string } {
  try {
    const diagnostics = Office.context?.diagnostics;
    const host = diagnostics?.host ?? Office.context?.host;
    const platform = diagnostics?.platform ?? Office.context?.platform;
    return {
      ...(typeof host === "string" ? { host } : {}),
      ...(typeof platform === "string" ? { platform } : {}),
    };
  } catch {
    return {};
  }
}

/**
 * Open one of the site's own pages in the user's browser.
 *
 * `Office.context.ui.openBrowserWindow` is the supported route and the one
 * `docs/DESIGN.md` section 7 names — a plain `window.open` or a target="_blank"
 * link is unreliable in a desktop task pane, where the pane is a WebView rather
 * than a tab and there is no browser around it to open a tab in. It needs
 * requirement set OpenBrowserWindowApi 1.1, which is not in this add-in's
 * floor, so it is PROBED rather than assumed, and `window.open` is the fallback
 * for the host that lacks it — on the web that is a real tab, so the fallback
 * is not a worse experience there.
 *
 * Answers whether anything opened, so the pane can say so rather than leaving a
 * click that silently did nothing.
 */
export function openExternal(url: string): boolean {
  try {
    if (Office.context?.requirements?.isSetSupported("OpenBrowserWindowApi", "1.1")) {
      Office.context.ui.openBrowserWindow(url);
      return true;
    }
  } catch {
    // Fall through: a host that raises from the supported route still has the
    // fallback, and a report link is not worth a broken pane.
  }
  try {
    // `noopener` because the opened page must not get a handle on the pane:
    // it is our own page today, and a window handle into a task pane holding
    // somebody's presentation is not a thing to hand out on trust.
    return window.open(url, "_blank", "noopener,noreferrer") !== null;
  } catch {
    return false;
  }
}

/** What the host answered after being asked to select a slide. */
export interface Selected {
  /** False when the host has no `setSelectedSlides` (PowerPointApi 1.5); nothing was asked. */
  supported: boolean;
  /** `getSelectedSlides` after the call, first item the active slide; null when the host did not answer. */
  selected: string[] | null;
  /** Why the call raised, bounded, if it did. */
  error?: string;
}

/**
 * Go to a slide, and read back where the host says it is.
 *
 * `setSelectedSlides` is PowerPointApi 1.5. It is the one selection WRITE this
 * add-in makes, and the reason it may be made is in `src/host/jump.ts`: the
 * sibling that ships it has measured it in every archived round on the web
 * (2,429 rungs, 2026-08-13 to 2026-09-04, none silent) and the call the family
 * saw wedge the host is `setSelectedShapes`, which is still never called here.
 * No sibling measured this on Windows or Mac, so the call is never trusted on
 * its own: the same batch reads the selection back, and `jumpOutcome` claims
 * the jump only when that read names the slide. Office-js#3552 says desktop
 * THROWS while the notes pane has focus; that is the `error` branch.
 *
 * Bounded by the ordinary read budget rather than a glance: the user clicked
 * for this, and a silent host is reported as "did not say", never as moved.
 */
export async function selectSlide(id: string): Promise<Selected> {
  if (!hostSupports("1.5")) return { supported: false, selected: null };
  try {
    const selected = await withTimeout(
      PowerPoint.run(async (context) => {
        context.presentation.setSelectedSlides([id]);
        await context.sync();
        const after = context.presentation.getSelectedSlides();
        after.load("items/id");
        await context.sync();
        return after.items.map((s) => s.id);
      }),
      BUDGET.read,
      "moving to a slide",
    );
    return { supported: true, selected };
  } catch (e) {
    return { supported: true, selected: null, error: readable(e) };
  }
}

/**
 * The URL the host gives for the open deck, or nothing when it will not say.
 *
 * `docs/DESIGN.md` section 4 keeps the way you were browsing per DECK, and this
 * is what tells one deck from another. `Office.context.document.url` is Common
 * API and needs no requirement set; on an unsaved deck it can be an empty
 * string rather than absent, and on some hosts the property is missing
 * altogether, so both answer `undefined` here and `deckKey` in `src/host` is
 * where the fallback is decided.
 *
 * Nothing is judged here, and nothing is stored: the URL can name a client, a
 * project or a person, and `deckKey` hashes it before it reaches storage.
 */
export function deckUrl(): string | undefined {
  try {
    const url = Office.context?.document?.url;
    return typeof url === "string" && url !== "" ? url : undefined;
  } catch {
    return undefined;
  }
}

/**
 * The background colour of PowerPoint's own chrome, or nothing when it will not say.
 *
 * `Office.context.officeTheme` is Common API and needs no requirement set. It
 * is absent outside a host — every time this pane is opened in a browser to
 * look at it — and the property access itself can throw on a host that
 * provides `Office.context` without a theme, so both answer `undefined` and
 * `paneTheme` in `src/host` is where the colour becomes a theme.
 *
 * Read once, on ready. There is no theme-change event for a PowerPoint pane:
 * the typings put `OfficeThemeChanged` on Outlook's `Mailbox` and nowhere else,
 * so switching PowerPoint's theme mid-session needs the pane reopened.
 */
export function themeBackground(): string | undefined {
  try {
    const body = Office.context?.officeTheme?.bodyBackgroundColor;
    return typeof body === "string" && body !== "" ? body : undefined;
  } catch {
    return undefined;
  }
}
