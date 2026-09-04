/**
 * The pane's entry point, and the only file here allowed to touch Office.js.
 *
 * It wires three things that are each tested somewhere else: the catalogue
 * (built by the harvest, checked by the suite), the picker's DOM (`render.ts`,
 * checked in jsdom), and the insert (`src/core` and `src/office`, checked by
 * the splice sweep). This file has no logic of its own worth testing and is
 * excluded from coverage for that reason — pooling it with the engine would
 * produce one number that hides both.
 */
import { Pkg } from "../core/pptx/pkg.js";
import { splice } from "../core/element/splice.js";
import { choosePlacement } from "../host/placement.js";
import type { CatalogueIndex, ElementPayload } from "../core/catalogue/types.js";
import {
  hostEnvironment,
  insertSpliced,
  occupiedOn,
  readDeck,
  ready,
  removeSlideAt,
  selectedSlideIndex,
} from "../office/powerpoint.js";
import { renderPicker, statusLine } from "./render.js";

const app = document.getElementById("app");
let index: CatalogueIndex | undefined;
let query = "";
let busy = false;

function say(text: string, warn = false): void {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = text;
  el.className = warn ? "notice warn" : "status";
}

function draw(): void {
  if (!app || !index) return;
  const focused = document.activeElement?.id === "q";
  const caret = (document.getElementById("q") as HTMLInputElement | null)?.selectionStart ?? null;
  app.innerHTML = renderPicker(index, query);
  const box = document.getElementById("q") as HTMLInputElement | null;
  if (box && focused) {
    box.focus();
    if (caret !== null) box.setSelectionRange(caret, caret);
  }
}

/**
 * Insert one element.
 *
 * The whole flow, in the order the file is touched: read the deck, splice the
 * element into the slide the user is on, reduce the package to that one slide,
 * hand it back, and remove the slide it replaced.
 *
 * The removal is LAST and is conditional on the insert having landed. Removing
 * first would be faster to write and would lose the user's slide on any host
 * that accepts an insert it does not perform — which is the documented
 * behaviour of the host this runs on.
 */
async function insert(id: string): Promise<void> {
  if (busy || !index) return;
  const entry = index.elements.find((e) => e.id === id);
  if (!entry) return;
  busy = true;
  try {
    say(statusLine({ phase: "reading" }));
    const [base64, where] = await Promise.all([readDeck(), selectedSlideIndex()]);

    say(statusLine({ phase: "splicing", name: entry.name }));
    const payload = (await fetch(`./catalogue/elements/${encodeURIComponent(id)}.json`).then((r) => {
      if (!r.ok) throw new Error(`the library file for “${entry.name}” could not be loaded`);
      return r.json();
    })) as ElementPayload;

    const pkg = await Pkg.open(base64);
    const slides = await pkg.slidePaths();
    const target = slides[where.index];
    if (target === undefined) throw new Error("the slide you are on is no longer in the presentation");

    const canvas = await pkg.slideSize();
    const at = choosePlacement(payload.bounds, canvas, await occupiedOn(where.index));
    const report = await splice(pkg, target, payload, at);
    await pkg.keepOnlySlide(target);

    say(statusLine({ phase: "inserting" }));
    const outcome = await insertSpliced(await pkg.toBase64(), where.index);
    if (!outcome.ok) {
      say(`That did not work: ${outcome.detail}`, true);
      return;
    }
    // Only now is the original expendable — the replacement is in the deck.
    const removed = await removeSlideAt(where.index);
    if (!removed.ok) {
      say(
        `“${entry.name}” was added, but the slide it replaces is still there — you now have both, and can delete slide ${where.index + 1}.`,
        true,
      );
      return;
    }
    const nudged = at.reason !== "as-authored";
    say(
      statusLine({ phase: "done", name: entry.name }) +
        (nudged ? ` Moved clear of what was already on the slide.` : "") +
        (report.droppedRefs.length > 0 ? ` ${report.droppedRefs.length} reference(s) could not be carried.` : ""),
    );
  } catch (e) {
    say(`That did not work: ${e instanceof Error ? e.message : String(e)}`, true);
  } finally {
    busy = false;
  }
}

/** Draw the picker and wire it up. Everything that does not need a host. */
async function start(): Promise<void> {
  try {
    index = (await fetch("./catalogue/index.json").then((r) => r.json())) as CatalogueIndex;
    if (index.version !== 1) throw new Error("this library was built by a different version of the add-in");
  } catch (e) {
    if (app) {
      app.innerHTML = `<p class="notice warn">The element library could not be loaded: ${e instanceof Error ? e.message : String(e)}</p>`;
    }
    return;
  }

  draw();

  app?.addEventListener("input", (ev) => {
    const box = ev.target as HTMLInputElement;
    if (box.id !== "q") return;
    query = box.value;
    draw();
  });

  app?.addEventListener("click", (ev) => {
    const card = (ev.target as HTMLElement).closest(".card");
    const id = card?.getAttribute("data-id");
    if (id) void insert(id);
  });
}

/**
 * Whether Office.js is here at all.
 *
 * Outside PowerPoint the global does not merely fail to initialise — in this
 * repo's own preview it is UNDEFINED, because the pane is opened from a plain
 * web server and the script tag's CDN may not be reachable. A pane that throws
 * on the first line is a pane nobody can look at, and looking at it is the only
 * check the suite cannot perform: jsdom has no layout and no colour.
 *
 * So the picker runs either way and only the INSERT needs a host. Clicking a
 * card without one fails at `readDeck` and says so, which is the honest
 * outcome — the alternative, disabling the cards, would hide the very state
 * a screenshot pass is trying to photograph.
 */
function inHost(): boolean {
  return typeof Office !== "undefined" && typeof Office.onReady === "function";
}

if (inHost()) {
  /**
   * Two `void`s, and both are load-bearing.
   *
   * `Office.onReady` RETURNS a promise, and it takes a callback whose return
   * value it ignores — so an async callback hands it a second promise nothing
   * is waiting on. A rejection in either is unhandled, and on some hosts that
   * is a silent dead pane rather than an error anybody sees. Every path inside
   * catches its own failure and renders a sentence, so there is nothing left to
   * propagate.
   */
  void Office.onReady(
    () =>
      void (async () => {
        const stamp = document.getElementById("build");
        const env = hostEnvironment();
        if (stamp) stamp.textContent = env.build;

        // Office's own theme, not the browser's: PowerPoint can be dark while
        // the OS is light, and the pane is inside PowerPoint. The stylesheet's
        // `prefers-color-scheme` block is the fallback for the case Office
        // never answers, which is exactly how the pane is inspected here.
        try {
          const body = Office.context.officeTheme?.bodyBackgroundColor;
          if (body) {
            const hex = body.replace("#", "");
            const lum = parseInt(hex.slice(0, 2), 16) + parseInt(hex.slice(2, 4), 16) + parseInt(hex.slice(4, 6), 16);
            document.documentElement.dataset.theme = lum < 384 ? "dark" : "light";
          }
        } catch {
          // No theme is not an error.
        }

        const floor = ready();
        if (!floor.ok) {
          if (app) app.innerHTML = `<p class="notice warn">${floor.detail}</p>`;
          return;
        }
        await start();
      })(),
  );
} else {
  const stamp = document.getElementById("build");
  if (stamp) stamp.textContent = typeof __BUILD_STAMP__ === "string" ? `${__BUILD_STAMP__} · no host` : "no host";
  void start();
}
