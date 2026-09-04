/**
 * The pane's DOM, as pure functions over the catalogue.
 *
 * No Office.js here — `main.ts` is the only file in this directory allowed to
 * touch it. That is what lets the picker's behaviour be checked in jsdom: which
 * elements a search matches, what an empty result says, whether a section with
 * nothing in it is drawn at all.
 *
 * The pane cannot be judged by the suite alone: jsdom has no layout and no
 * colour. `scripts/pane-shots.mjs` renders every state at 320 and 512 in both
 * themes and measures the three things a screenshot cannot show — horizontal
 * overflow, text contrast, and where the keyboard focus ring is.
 */
import type { CatalogueIndex, ElementEntry } from "../core/catalogue/types.js";

/**
 * Which elements a query matches.
 *
 * Matched against the element's name AND its section's name, because the
 * library's names are written assuming the section is visible: "2 vertikale" is
 * three separate elements in three sections, and searching for "streger" should
 * find the one under "Kun streger" even though its own name never says so.
 *
 * Diacritic-insensitive, so "grå" is found by typing "gra". Everything in this
 * library is Danish and a task-pane search box is not where somebody wants to
 * be precise about their vowels.
 */
export function matches(index: CatalogueIndex, query: string): ElementEntry[] {
  const q = fold(query);
  if (!q) return index.elements;
  const sectionName = new Map(index.sections.map((s) => [s.id, fold(s.name)]));
  const terms = q.split(/\s+/).filter(Boolean);
  return index.elements.filter((e) => {
    const hay = `${fold(e.name)} ${sectionName.get(e.section) ?? ""} ${e.kinds.join(" ")}`;
    // Every term must appear. A search of "hvid 3" should narrow, not widen.
    return terms.every((t) => hay.includes(t));
  });
}

/** Lowercase, and stripped of the accents a Danish keyboard puts on. */
export function fold(text: string): string {
  return text
    .toLowerCase()
    .replace(/æ/g, "ae")
    .replace(/ø/g, "oe")
    .replace(/å/g, "aa")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

/** The elements of each section that survived a search, sections with none dropped. */
export function grouped(index: CatalogueIndex, shown: ElementEntry[]): { name: string; items: ElementEntry[] }[] {
  const byId = new Map(shown.map((e) => [e.id, e]));
  const out: { name: string; items: ElementEntry[] }[] = [];
  for (const section of index.sections) {
    const items = index.elements.filter((e) => e.section === section.id && byId.has(e.id));
    // A section header with nothing under it is a promise the pane does not
    // keep — it reads as a section that failed to load rather than one that
    // simply has no match.
    if (items.length > 0) out.push({ name: section.name, items });
  }
  return out;
}

/** How many, said in words rather than as a bare number. */
export function countLine(shown: number, total: number): string {
  if (shown === total) return `${total} element${total === 1 ? "" : "s"}`;
  if (shown === 0) return "nothing matches";
  return `${shown} of ${total}`;
}

/**
 * A label short enough for a two-line card, truncated in the STRING.
 *
 * Not in CSS, and that is a decision rather than a shortcut. `-webkit-line-clamp`
 * needs `display: -webkit-box`, which is blockified to `flow-root` in several
 * situations; the standard `line-clamp` landed recently and behaves differently
 * again. Measured here, the label computed to `flow-root` with the clamp inert
 * and the text hard-clipped mid-line — no ellipsis, half a third line showing.
 *
 * The pane runs in PowerPoint's embedded WebView, whose engine version is not
 * this repo's to choose and cannot be tested from here. A character budget
 * behaves identically on every one of them, and it is checkable in jsdom, where
 * a CSS clamp is invisible.
 *
 * The full name stays in the card's `title`, so nothing is actually lost — and
 * the budget is generous enough that only the longest quarter of the library is
 * touched at all.
 */
export function clampName(name: string, budget = 46): string {
  if (name.length <= budget) return name;
  // Cut at a space so a truncation never ends mid-word, unless the first
  // `budget` characters contain no space at all — which for this library would
  // be a name nobody could read anyway.
  const cut = name.slice(0, budget);
  const space = cut.lastIndexOf(" ");
  return `${(space > budget * 0.6 ? cut.slice(0, space) : cut).trimEnd()}…`;
}

/** Escape text that is going into markup. Every name here came out of a file. */
export function esc(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * The picker.
 *
 * The preview SVG is inserted as markup rather than as an `<img src="data:">`,
 * because a data URI per card is ninety-eight extra fetches the browser has to
 * decode, and because inline SVG inherits the pane's own colours if it ever
 * needs to. The SVG is generated by this repo's own harvest from a deck in this
 * repo — it is not user input — but it is escaped-by-construction anyway: only
 * `previewSvg` writes it, and it emits nothing but rects.
 */
export function renderPicker(index: CatalogueIndex, query: string): string {
  const shown = matches(index, query);
  const sections = grouped(index, shown);
  const head =
    `<label class="visually-hidden" for="q">Search the library</label>` +
    `<input class="search" id="q" type="search" placeholder="Search ${index.elements.length} elements…" value="${esc(query)}" />` +
    `<p class="count">${esc(countLine(shown.length, index.elements.length))}</p>`;

  if (sections.length === 0) {
    return `${head}<p class="empty">Nothing matches “${esc(query)}”.</p>`;
  }

  const body = sections
    .map(
      (s) =>
        `<section class="section"><h2>${esc(s.name)}</h2><div class="grid">` +
        s.items
          .map(
            (e) =>
              `<button class="card" data-id="${esc(e.id)}" title="${esc(e.name)}">` +
              `<span class="thumb">${e.preview}</span>` +
              `<span class="label">${esc(clampName(e.name))}</span>` +
              `</button>`,
          )
          .join("") +
        `</div></section>`,
    )
    .join("");

  return head + body + `<p class="status" id="status"></p>`;
}

/** What the pane says while an insert is running, and after it. */
export function statusLine(state: {
  phase: "idle" | "reading" | "splicing" | "inserting" | "done" | "failed";
  name?: string;
  detail?: string;
}): string {
  switch (state.phase) {
    case "reading":
      return `Reading your presentation…`;
    case "splicing":
      return `Building “${state.name ?? "the element"}” into your slide…`;
    case "inserting":
      return `Handing the slide back to PowerPoint…`;
    case "done":
      return `Inserted “${state.name ?? "the element"}”.`;
    case "failed":
      return state.detail ?? "That did not work.";
    default:
      return "";
  }
}
