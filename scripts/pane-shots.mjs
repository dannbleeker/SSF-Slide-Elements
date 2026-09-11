#!/usr/bin/env node
/**
 * Render every pane state to a PNG, at both ends of the width it must hold.
 *
 * The pane is the one surface the suite cannot judge. `pane-render.test.ts`
 * pins its behaviour in jsdom, which has no layout and no colour, so a rule
 * like the ORANGE BUDGET — one orange element per view — is invisible there
 * unless somebody thought to assert it.
 *
 *   npx vite --port 5199 --strictPort &
 *   node scripts/pane-shots.mjs
 *
 * 320 and 512 are the ends of the range a task pane is dragged between. A pane
 * judged at one width is a layout that breaks at the other. Both THEMES, too:
 * PowerPoint can be dark while the OS is light, and the dark palette is a
 * separate set of tokens that nothing else in this repo exercises.
 *
 * It also MEASURES, because the things a screenshot shows are numbers a reader
 * cannot take off a PNG, and every one of them has produced a real defect on a
 * sibling project: horizontal overflow from a spaceless host error, text
 * contrast on the dark palette, a near-invisible focus ring, a 19px hit area.
 *
 * Findings are printed and the process exits 1, so this can be read by a person
 * or wired to something. Disabled controls are exempt from the contrast rule,
 * which is what WCAG 1.4.3 says and not a convenience: a greyed-out button is
 * meant to read as unavailable.
 */
import { existsSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { chromium } from "playwright";

/**
 * axe-core, read off disk and injected per page.
 *
 * It answers the half of accessibility a person cannot eyeball — names, roles,
 * labels, duplicate ids, a control with nothing to call it — over the same
 * state list, both widths and both themes. What it does NOT answer is whether
 * the pane SAYS anything when it changes: nothing static can see a missing
 * live region. It is a floor, not a verdict.
 */
const AXE = readFileSync("node_modules/axe-core/axe.min.js", "utf8");

const PORT = process.env.PANE_PORT ?? "5199";
const OUT = process.env.PANE_SHOTS ?? "/tmp/pane-shots";

/**
 * A notice with nothing in it a line can break at.
 *
 * `overflow-wrap` is what keeps this inside a 320px frame, and nothing else in
 * these fixtures has a word long enough to need it. A host error with no
 * spaces is what `readable` caps at 400 characters; a cap is not a break.
 */
const SPACELESS = `PowerPoint refused the insert: ${"GeneralException".repeat(20)}`;

/** A long sentence WITH spaces, so ordinary wrapping is measured too. */
const LONG =
  "PowerPoint answered, and what it said was longer than this pane is wide, which is the ordinary case for a " +
  "host that echoes its argument back into the message it raises and the reason the notice wraps at all.";

/**
 * The states the pane can be in today, each with what it claims to draw.
 *
 * `shows` and `hides` are asserted before anything is measured: a state named
 * for a control it does not render is a shot of the wrong screen, and every
 * measurement taken from it is about something else. A sibling had exactly
 * that for an unknown number of runs.
 */
/**
 * A library small enough to read in a shot and varied enough to exercise the
 * tile: a plain element, a sized run that draws a stepper, and a second
 * category so a closed header is measured beside an open one.
 *
 * Built here rather than fetched: the audit calls `render` directly, so it
 * never loads a catalogue, and a fixture that had to be fetched would make the
 * shots depend on the network.
 */
const el = (id, over = {}) => ({
  id,
  key: over.key ?? id,
  name: over.name ?? id,
  category: over.category ?? { key: "boxes", name: "White boxes" },
  slide: 1,
  kind: "slide",
  box: over.box ?? { x: 0.06, y: 0.23, w: 0.88, h: 0.65 },
  landing: "layout",
  shapes: 4,
  tags: over.tags ?? ["boxes", "white"],
  markup: { xml: "", rels: [], parts: [] },
  ...over,
});

const LIBRARY = {
  size: "16:9",
  width: 12192000,
  height: 6858000,
  version: "shots",
  categories: [
    { key: "boxes", name: "White boxes" },
    { key: "stamps", name: "Stamps and labels" },
  ],
  elements: [
    el("one-box", { name: "One box" }),
    el("two-boxes", { name: "Two boxes" }),
    el("flow-1", {
      name: "Process flow, 1 box",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 1 },
    }),
    el("flow-2", {
      name: "Process flow, 2 boxes",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 2 },
    }),
    el("flow-3", {
      name: "Process flow, 3 boxes",
      tags: ["flow"],
      run: { key: "Process flow, N boxes", noun: "boxes", count: 3 },
    }),
    el("approved", {
      name: "Approved stamp",
      key: "Godkendt",
      category: { key: "stamps", name: "Stamps and labels" },
      landing: "top-right",
      kind: "part",
      tags: ["stamp"],
    }),
  ],
};

/** The shape every browse state starts from. */
const BROWSING = {
  // Every other shot is of a pane somebody has used before; the first-run
  // note has its own state below.
  coached: true,
  query: "",
  tags: [],
  open: ["boxes"],
  settings: { target: "onto", group: true, colours: "deck" },
  undo: 0,
  favourites: [],
  recent: [],
  library: LIBRARY,
  slide: 2,
};

/**
 * The states the pane can be in today, each with what it claims to draw.
 *
 * `shows` and `hides` are asserted before anything is measured: a state named
 * for a control it does not render is a shot of the wrong screen, and every
 * measurement taken from it is about something else. A sibling had exactly
 * that for an unknown number of runs.
 */
const STATES = [
  { name: "loading", step: "loading", state: { ...BROWSING, library: undefined }, shows: ["insert"] },
  {
    name: "problem",
    step: "problem",
    state: { ...BROWSING, library: undefined, problem: "The library did not load: the site did not answer." },
    shows: ["retry"],
    hides: ["search", "tile"],
  },
  { name: "browse", step: "browse", state: BROWSING, shows: ["search", "gear", "tile", "star", "category", "step"] },
  {
    // The first open ever (section 4): three lines and one way out.
    name: "browse-first-run",
    step: "browse",
    state: { ...BROWSING, coached: false },
    shows: ["coached", "search", "tile"],
  },
  {
    // The stepper greying the sizes the query did not ask for (section 8).
    // Here for the CONTRAST: a greyed control still has to read as a control,
    // and this fixture cannot produce a query that spans two categories, so
    // the chips beside it are covered by test/pane-render.test.ts instead.
    name: "browse-narrowed",
    step: "browse",
    state: { ...BROWSING, query: "1 box" },
    shows: ["search", "tile"],
  },
  {
    // A search that found nothing, WITH a way out of it. The existing
    // nothing-found state has a query that is near nothing, so it shows only
    // the clear button; this one is a typo of a real name.
    name: "browse-did-you-mean",
    step: "browse",
    state: { ...BROWSING, query: "bax" },
    shows: ["search", "guess", "clear"],
    hides: ["tile"],
  },
  {
    // The preview card, which at 512 docks beside the list and at 320 pins over
    // the top of it. Both widths are shot, so the audit measures the card in the
    // narrow case where it overlaps and the wide one where it must not.
    name: "browse-previewing",
    step: "browse",
    state: { ...BROWSING, previewing: "one-box" },
    shows: ["search", "tile"],
  },
  {
    name: "browse-closed",
    step: "browse",
    state: { ...BROWSING, open: [] },
    shows: ["search", "category"],
    hides: ["tile"],
  },
  {
    name: "browse-chosen",
    step: "browse",
    state: { ...BROWSING, chosen: "one-box" },
    shows: ["insert", "tile"],
  },
  {
    name: "browse-gear",
    step: "browse",
    state: { ...BROWSING, gear: true },
    shows: ["target", "group", "colours", "report", "catalogue"],
  },
  {
    // A part already in the deck, with the confirm open. The only thing in this
    // pane that takes something OUT of a deck, so it is the state worth seeing:
    // the question, what it would touch, and that the pane cannot undo it.
    name: "browse-removing",
    step: "browse",
    state: {
      ...BROWSING,
      open: ["stamps"],
      used: [{ element: "approved", slides: [2, 5, 9] }],
      removing: { id: "approved", slides: [2, 5, 9], done: 0, where: "stamps" },
    },
    shows: ["remove-go", "remove-cancel"],
  },
  {
    // The preview card over a slide the pane has read: the element's landing in
    // blue, and in grey what the slide already holds.
    name: "browse-card-on-slide",
    step: "browse",
    state: {
      ...BROWSING,
      open: ["boxes"],
      previewing: "one-box",
      slide: 2,
      onSlide: {
        slide: 2,
        boxes: [
          { x: 0.06, y: 0.08, w: 0.55, h: 0.14 },
          { x: 0.06, y: 0.3, w: 0.42, h: 0.5 },
          { x: 0.54, y: 0.3, w: 0.4, h: 0.24 },
        ],
      },
    },
    shows: ["tile"],
  },
  {
    // The tile's right-click menu: the other insert target, over the tile it
    // belongs to. Drawn on an OPEN category, because a menu on a collapsed one
    // would be a menu over nothing.
    name: "browse-tile-menu",
    step: "browse",
    state: { ...BROWSING, open: ["boxes"], menuFor: "boxes:one-box" },
    shows: ["other-target", "tile"],
  },
  {
    // "Used in this deck", read: the state nothing else in this list covers,
    // and the one where a long element name meets a long list of slide numbers.
    name: "browse-used",
    step: "browse",
    state: {
      ...BROWSING,
      used: [
        { element: "one-box", slides: [2] },
        { element: "two-boxes", slides: [3, 5, 11] },
        { element: "gone-from-the-library", slides: [7] },
      ],
    },
    shows: ["used"],
  },
  {
    name: "browse-searched",
    step: "browse",
    state: { ...BROWSING, query: "flow" },
    shows: ["search", "tile"],
  },
  {
    name: "browse-nothing-found",
    step: "browse",
    state: { ...BROWSING, query: "nothing at all matches this" },
    shows: ["clear"],
    hides: ["tile"],
  },
  {
    name: "browse-busy",
    step: "browse",
    state: { ...BROWSING, chosen: "one-box", busy: true },
    shows: ["tile"],
  },
  {
    name: "browse-after-insert",
    step: "browse",
    state: {
      ...BROWSING,
      chosen: "one-box",
      favourites: ["one-box"],
      recent: ["one-box"],
      undo: 1,
      outcome: { ok: true, byHand: false, name: "One box", detail: "12 → 13 → 12 slides, slide 2 replaced." },
    },
    shows: ["again", "undo", "star"],
  },
  {
    name: "browse-by-hand",
    step: "browse",
    state: {
      ...BROWSING,
      outcome: {
        ok: false,
        byHand: true,
        name: "One box",
        detail: "The deck grew by one but the copy could not be removed: delete slide 2 by hand.",
      },
    },
    shows: ["gear"],
  },
  {
    name: "browse-borrowed",
    step: "browse",
    state: { ...BROWSING, library: { ...LIBRARY, borrowed: "4:3 library, scaled to A4 slides." } },
    shows: ["search"],
  },
  { name: "browse-spaceless-notice", step: "browse", state: { ...BROWSING, notice: SPACELESS }, shows: ["search"] },
  { name: "browse-long-notice", step: "browse", state: { ...BROWSING, notice: LONG }, shows: ["search"] },
];

// The bundled browser and the installed playwright can disagree on build
// number in THIS environment, so the binary is named rather than discovered.
//
// Only when it is actually there. On a CI runner playwright installs its own
// and knows where it put it, and naming a path that does not exist fails the
// launch with an error about a missing executable rather than about anything
// this script is for. `CHROMIUM` overrides both.
const CONTAINER_CHROMIUM = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
const EXECUTABLE = process.env.CHROMIUM || (existsSync(CONTAINER_CHROMIUM) ? CONTAINER_CHROMIUM : undefined);

/**
 * What the page can say about itself that a PNG cannot.
 *
 * Runs INSIDE the page, so it reads computed styles and real geometry rather
 * than anything this file believes about the stylesheet.
 *
 * Contrast is measured against the nearest ancestor with a non-transparent
 * background, which is what the eye sees; a colour with no opaque ground
 * behind it falls back to white rather than being skipped, because skipping is
 * how a measurement quietly stops measuring.
 */
function audit() {
  const numbers = (/** @type {string} */ s) => (s.match(/[\d.]+/g) ?? []).map(Number);
  const luminance = (/** @type {number[]} */ [r, g, b]) => {
    const channel = (v) => {
      const x = v / 255;
      return x <= 0.03928 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4;
    };
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
  };
  const groundOf = (/** @type {Element} */ el) => {
    for (let node = el; node; node = node.parentElement) {
      const colour = numbers(getComputedStyle(node).backgroundColor);
      if (colour.length >= 3 && (colour[3] === undefined || colour[3] > 0)) return colour;
    }
    return [255, 255, 255];
  };

  const findings = [];
  const root = document.documentElement;
  // The DOCUMENT, not the pane: a child can overhang its parent harmlessly,
  // and what matters is whether the frame has to scroll sideways.
  if (root.scrollWidth > root.clientWidth) {
    const past = [];
    for (const el of document.querySelectorAll("#pane *")) {
      const box = el.getBoundingClientRect();
      if (box.right > root.clientWidth + 0.5 || box.left < -0.5) {
        past.push(`${el.tagName.toLowerCase()}.${el.className || "-"} "${(el.textContent ?? "").trim().slice(0, 40)}"`);
      }
    }
    // A block-level element's BOX is clamped to its container, so the one that
    // overflows is usually a paragraph whose text is wider than it is. Name
    // the widest by `scrollWidth` when no box is past the edge, or the finding
    // says only that something is wrong.
    if (past.length === 0) {
      let widest = null;
      for (const el of document.querySelectorAll("#pane *")) {
        // An element that scrolls or clips its own content cannot push the
        // document — a textarea's `scrollWidth` is its text, and naming it
        // sends the reader at the one box on the screen that is fine.
        if (getComputedStyle(el).overflowX !== "visible") continue;
        if (el.scrollWidth > root.clientWidth && (!widest || el.scrollWidth > widest.scrollWidth)) widest = el;
      }
      if (widest) {
        past.push(
          `${widest.tagName.toLowerCase()}.${widest.className || "-"} is ${widest.scrollWidth}px wide "${(widest.textContent ?? "").trim().slice(0, 40)}"`,
        );
      }
    }
    findings.push(`overflows sideways: ${root.scrollWidth}px in ${root.clientWidth} — ${past[0] ?? "nothing named"}`);
  }

  const walk = (el) => {
    for (const node of el.childNodes) {
      if (node.nodeType === 3 && (node.textContent ?? "").trim() !== "") {
        // WCAG 1.4.3 exempts an inactive control: greyed out is meant to read
        // as unavailable, and holding it to 4.5:1 would remove the signal.
        if (el.disabled === true || el.closest("[disabled]")) continue;
        const style = getComputedStyle(el);
        const ratio =
          (Math.max(luminance(numbers(style.color)), luminance(groundOf(el))) + 0.05) /
          (Math.min(luminance(numbers(style.color)), luminance(groundOf(el))) + 0.05);
        const size = Number.parseFloat(style.fontSize);
        const large = size >= 24 || (size >= 18.66 && Number.parseInt(style.fontWeight, 10) >= 700);
        const need = large ? 3 : 4.5;
        if (ratio < need) {
          findings.push(
            `contrast ${ratio.toFixed(2)}:1 (needs ${need}) on ${el.tagName.toLowerCase()}.${el.className || "-"} "${(node.textContent ?? "").trim().slice(0, 40)}"`,
          );
        }
      } else if (node.nodeType === 1) walk(node);
    }
  };
  walk(document.body);

  // WHERE THE KEYBOARD IS, which is the third thing a PNG cannot show and the
  // one a screenshot cannot show at all: a focus ring is only on screen while
  // something is focused.
  //
  // On a sibling's pane it found the sharpest of the three. Buttons were not in the stylesheet's
  // focus rule, so they fell back to Chrome's own ring — `rgb(16, 16, 16)`,
  // near-black, drawn on the dark theme's near-black pane at 1.03:1. Every
  // chip, both disclosures, the back links and the undo button, with no
  // visible focus anywhere in that theme except inside a text box.
  //
  // The caller presses Tab once before this runs. Chrome matches
  // `:focus-visible` on a programmatic `focus()` only once the last
  // interaction was a keyboard one, so without that every button here reports
  // no outline at all — a measurement that answers "clean" because it never
  // looked.
  const focusable = document.querySelectorAll(
    "#pane button:not([disabled]), #pane input:not([disabled]), #pane select:not([disabled]), #pane textarea:not([disabled])",
  );
  const already = new Set();
  for (const el of focusable) {
    el.focus();
    if (document.activeElement !== el) continue; // not actually focusable
    const style = getComputedStyle(el);
    const key = `${el.tagName}.${el.className}`;
    if (already.has(key)) continue;
    already.add(key);
    const width = Number.parseFloat(style.outlineWidth);
    if (!(width > 0) || style.outlineStyle === "none") {
      findings.push(`no focus ring on ${el.tagName.toLowerCase()}.${el.className || "-"}`);
      continue;
    }
    const ratio =
      (Math.max(luminance(numbers(style.outlineColor)), luminance(groundOf(el.parentElement ?? el))) + 0.05) /
      (Math.min(luminance(numbers(style.outlineColor)), luminance(groundOf(el.parentElement ?? el))) + 0.05);
    // 3:1, which is what WCAG 1.4.11 asks of a non-text indicator.
    if (ratio < 3) {
      findings.push(`focus ring ${ratio.toFixed(2)}:1 (needs 3) on ${el.tagName.toLowerCase()}.${el.className || "-"}`);
    }
  }

  /*
   * The fifth thing: whether a FINGER can hit what a mouse can.
   *
   * A task pane runs on touch-only devices, which AppSource asks about, and
   * nothing here had ever measured a hit area. The incident this was written
   * for is a SIBLING's — sixteen controls under 24 CSS px, every "Back to…",
   * every option link and both run-log disclosures, each at the line box of its
   * own text with `padding: 0`. This pane has none of those three kinds of
   * control, which is why the measure is worth keeping rather than dropping: it
   * is the check that would catch the first one to arrive.
   *
   * **24px BY SIZE, and the spacing exception is deliberately not taken.**
   * WCAG 2.5.8 lets an undersized target pass when a 24px circle centred on it
   * clears its neighbours, and by that reading this pane was already compliant:
   * a 19px link with 14px of air around it has 33px between centres. The first
   * version of this check said exactly that, and it was worthless — reverting
   * the padding left it green, because the criterion it quoted had never been
   * the thing being broken.
   *
   * So this is a house floor rather than the standard, and it is written down
   * as one. The pane is dragged to 320px, where controls stack tightly and the
   * spacing that excuses a small target is one layout change from being false;
   * a rule that depends on the gaps is a rule that stops holding the day
   * somebody tightens them.
   *
   * 24 and not the 44 the platforms recommend: 44 is guidance, and on a 320px
   * pane it costs the vertical space the overflow half of this audit exists to
   * protect. A gate that fires on a preference is a gate somebody switches off.
   */
  const targets = [
    ...document.querySelectorAll(
      "#pane button:not([disabled]), #pane a[href], #pane input:not([disabled]), #pane select:not([disabled]), #pane summary",
    ),
  ].filter((el) => {
    const r = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return Boolean(r.width || r.height) && style.display !== "none" && style.visibility !== "hidden";
  });
  const smallSeen = new Set();
  for (const el of targets) {
    const r = el.getBoundingClientRect();
    let short = Math.min(r.width, r.height);
    // A control inside a LABEL is hit by tapping the label, so the box a finger
    // meets is the label's. A sibling's row checkboxes were 13px and their label
    // 290x30; calling those a failure would be measuring the wrong element.
    const label = el.closest("label");
    if (label) {
      const lr = label.getBoundingClientRect();
      short = Math.max(short, Math.min(lr.width, lr.height));
    }
    if (short >= 24) continue;
    const key = `${el.tagName}.${el.className}`;
    if (smallSeen.has(key)) continue;
    smallSeen.add(key);
    findings.push(
      `hit area ${Math.round(r.width)}x${Math.round(r.height)} on ${el.tagName.toLowerCase()}.${el.className || "-"} — a finger wants 24 on the short side`,
    );
  }
  /*
   * The sixth thing: whether a label that WRAPS still starts where its
   * neighbours do.
   *
   * Inherited from a sibling, whose option links were `<button>` elements
   * written to read as links — and a button centres its label. This pane has no
   * link-shaped control yet, so `button.back` matches nothing until one is
   * added; the rule waits here rather than being rediscovered. While a label fits on one line
   * the button shrinks to it and the centring is invisible, so this survived
   * every screenshot and every rule above: nothing overflowed, contrast was
   * fine, the hit area was fine. At 320px "A blank cell leaves a blank — change
   * what happens" wraps, and its second line sat centred underneath two
   * siblings that were flush left.
   *
   * Measured off the LINE BOXES rather than off `text-align`, deliberately. A
   * computed-style check on a class is a restatement of the stylesheet and
   * passes for any value the stylesheet happens to hold; the line rects are
   * where the reader's eye actually lands, and they catch a centring that
   * arrives from a parent, from `direction`, or from a padding that only
   * applies to one side.
   *
   * Only the link-shaped controls. The primary button is centred on purpose and
   * a rule that fired on it would be switched off within a week.
   */
  for (const el of document.querySelectorAll("#pane button.back")) {
    const range = document.createRange();
    range.selectNodeContents(el);
    const lines = [...range.getClientRects()].filter((r) => r.width > 0);
    if (lines.length < 2) continue;
    const lefts = lines.map((r) => r.left);
    const ragged = Math.max(...lefts) - Math.min(...lefts);
    if (ragged <= 0.5) continue;
    findings.push(
      `wrapped label is not flush left: ${ragged.toFixed(1)}px between line starts on ` +
        `${el.tagName.toLowerCase()}.${el.className || "-"} "${(el.textContent ?? "").trim().slice(0, 40)}"`,
    );
  }

  return findings;
}

// EMPTIED first, not merely created. This directory is read by a person
// comparing renders, and a PNG left by an earlier run is indistinguishable from
// one this run produced — same name, same place, days older. It cost a false
// defect report on 2026-09-01: a shot from an earlier build showed the undo
// card, the current build's shot did not, and the difference read as a theme
// bug for twenty minutes. A stale artifact is not clutter, it is a wrong
// answer that looks like a right one.
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
const browser = await chromium.launch(EXECUTABLE ? { executablePath: EXECUTABLE } : {});
let taken = 0;
// Deduplicated across states: the same control is on every screen, and twelve
// copies of one finding is a report nobody reads to the end.
const found = new Map();
// Fixture claims that failed, kept apart from `found`: a broken fixture is a
// fact about this file, and reporting it beside a contrast failure would read
// as a defect in the pane.
/** @type {string[]} */
const claims = [];
for (const width of [320, 512]) {
  for (const theme of ["light", "dark"]) {
    for (const { name, step, state, shows, hides } of STATES) {
      const page = await browser.newPage({ viewport: { width, height: 620 } });
      // Office.js is fetched from Microsoft by `taskpane.html` and is not what
      // is being measured here — `render` is called directly. Refused rather
      // than waited on: on a machine that cannot reach it, every one of these
      // pages otherwise spends its connect timeout before drawing anything.
      await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
      await page.goto(`http://localhost:${PORT}/taskpane.html`);
      await page.evaluate(
        async ({ state, step, theme }) => {
          // The stamp `main.ts` writes from `Office.context.officeTheme`. Set
          // before the render so the first paint is the one measured.
          document.documentElement.setAttribute("data-theme", theme);
          const { render } = await import("/render.ts");
          render(document.getElementById("pane"), state, step);
        },
        { state, step, theme },
      );
      // WHAT THE FIXTURE CLAIMS TO BE, asserted before anything is measured.
      //
      // A state named for a control it does not render is a shot of the wrong
      // screen, and every measurement taken from it is about something else.
      // A sibling's `5-merge-done-undo` was exactly that for an unknown number of
      // runs: its undo card needed a `deckAtStart` no fixture set, so the audit
      // reported clean on a pane with no undo card in it, under a name that
      // said there was. Names are documentation; this makes them a check.
      //
      // `hides` is the same rule for a fixture whose subject is a control being
      // WITHHELD — without it, a negative case passes when the control is
      // missing for a reason the fixture was not testing.
      const drawn = new Set(
        await page.evaluate(() =>
          [...document.querySelectorAll("[data-action]")].map((e) => e.getAttribute("data-action")),
        ),
      );
      for (const action of shows ?? []) {
        if (!drawn.has(action)) claims.push(`${name}: claims to show "${action}" and does not`);
      }
      for (const action of hides ?? []) {
        if (drawn.has(action)) claims.push(`${name}: claims to withhold "${action}" and shows it`);
      }

      await page.screenshot({ path: `${OUT}/${width}-${theme}-${name}.png` });
      // BEFORE the Tab below: a focus ring in the accessibility tree is not
      // what axe is being asked about, and the shot above is the clean state.
      await page.addScriptTag({ content: AXE });
      // The strings are built INSIDE the page and only strings come back —
      // axe's result object is a large `any` across the boundary, and mapping
      // it here rather than there is how this file would start carrying an
      // untyped shape it never uses.
      /** @type {string[]} */
      const violations = await page.evaluate(async () => {
        /** @type {{ violations: { impact: string; id: string; help: string }[] }} */
        const run = await globalThis.axe.run(document, { resultTypes: ["violations"] });
        return run.violations.map((x) => `${x.impact}: ${x.id} — ${x.help}`);
      });
      for (const v of violations) {
        if (!found.has(v)) found.set(v, `${width} ${theme} ${name}`);
      }
      // AFTER the shot, and before the audit. It tells Chrome the last
      // interaction was a keyboard one, which is what makes `:focus-visible`
      // match the programmatic `focus()` the focus sweep uses — and it would
      // put a ring in the picture if it ran first.
      await page.keyboard.press("Tab");
      for (const finding of await page.evaluate(audit)) {
        if (!found.has(finding)) found.set(finding, `${width} ${theme} ${name}`);
      }
      await page.close();
      taken++;
    }
  }
}
await browser.close();
console.log(`${taken} shots in ${OUT}`);
// BEFORE the audit's own verdict. A fixture that did not render its subject
// makes every measurement taken from it meaningless, so it is not a finding
// among findings — it is a reason to distrust the run.
if (claims.length) {
  console.log("these fixtures did not render what their names claim:");
  for (const claim of [...new Set(claims)]) console.log(`  ${claim}`);
  process.exit(1);
}
if (found.size === 0) {
  console.log(
    "audit: nothing overflows, every live label clears its contrast floor, every control shows where the keyboard is, " +
      "every hit area is at least 24px on its short side, every wrapped link label starts flush left, " +
      "and axe finds no violation",
  );
} else {
  console.log(`audit: ${found.size} finding(s)`);
  for (const [finding, where] of found) console.log(`  ${where}: ${finding}`);
  process.exitCode = 1;
}
