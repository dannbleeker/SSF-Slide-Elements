/**
 * The pane's entry point, and the only file here that touches Office.js.
 *
 * Everything it shows comes from `render.ts` and everything it decides comes
 * from `steps.ts`, both checked by the suite without a PowerPoint anywhere.
 * `test/architecture.test.ts` holds that seam: a decision that migrates into
 * this file becomes untestable the moment it arrives.
 *
 * Excluded from coverage for the same reason — pooling it with the engine
 * would produce one number that hides both.
 */
import { ready as hostReady } from "../office/powerpoint.js";
import { render } from "./render.js";
import { EMPTY, type PaneState, type StepId } from "./steps.js";

const state: PaneState = EMPTY;
const step: StepId = "start";

function root(): HTMLElement {
  const node = document.getElementById("pane");
  if (!node) throw new Error("the pane's root element is missing");
  return node;
}

/**
 * The live region, made once and never rebuilt.
 *
 * `render` empties `#pane` and builds fresh elements on every draw, and a live
 * region CREATED with its content in it does not announce — the region has to
 * exist first and have text put into it. So this one lives outside the pane, is
 * made on the first draw, and is only ever written to.
 *
 * Made here rather than in `taskpane.html` so there is one definition and the
 * jsdom wiring tests get it without keeping a copy of the page's markup in step
 * with the real one. It is off-screen rather than `display: none`, which would
 * take it out of the accessibility tree along with everything in it.
 */
function liveRegion(): HTMLElement {
  const existing = document.getElementById("announcer");
  if (existing) return existing;
  const node = document.createElement("p");
  node.id = "announcer";
  node.className = "visually-hidden";
  // `polite`, never `assertive`: nothing here is urgent enough to cut across
  // what the user is already being told.
  node.setAttribute("role", "status");
  node.setAttribute("aria-live", "polite");
  document.body.append(node);
  return node;
}

/** The last thing announced, so the same sentence is not said twice. */
let announced = "";

function announce(text: string): void {
  // Only on a CHANGE: writing the same string back into a live region makes
  // some screen readers say it again.
  if (text === announced) return;
  announced = text;
  liveRegion().textContent = text;
}

function draw(): void {
  render(root(), state, step);
  // The region exists from the first draw, so the first sentence that ever
  // reaches it is announced rather than merely present.
  liveRegion();
  announce(state.notice ?? "");
}

/**
 * Follow PowerPoint's theme, not the browser's.
 *
 * The pane lives inside PowerPoint, which can be dark while the OS is light, so
 * `prefers-color-scheme` is the wrong question. `officeTheme` answers the right
 * one. Outside a host it is undefined — which is the case every time this pane
 * is opened in a browser to look at it — and the stylesheet's media query
 * carries that fallback.
 *
 * Read ONCE, on ready. There is no theme-change event for a PowerPoint pane:
 * the typings put `OfficeThemeChanged` on Outlook's `Mailbox` and nowhere else,
 * so switching PowerPoint's theme mid-session needs the pane reopened.
 */
function applyTheme(): void {
  const body = Office.context?.officeTheme?.bodyBackgroundColor;
  if (!body) return;
  const hex = body.replace("#", "");
  const n = Number.parseInt(hex.length === 3 ? [...hex].map((c) => c + c).join("") : hex, 16);
  if (Number.isNaN(n)) return;
  const luminance = ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
  document.documentElement.setAttribute("data-theme", luminance < 128 ? "dark" : "light");
}

/**
 * The build this pane was served from, in the header, before anything is run.
 *
 * PowerPoint caches the pane's HTML for about ten minutes, so opening it too
 * soon after a deploy tests code the host never fetched, and the result reads
 * as a clean run of the wrong build. The sibling projects record whole rounds
 * lost to it. Seven characters in the header is how the two are told apart.
 *
 * In the HEADER rather than in the pane, because the layout rule the pane was
 * approved on is that the primary button is the last element in the view.
 */
function showBuild(): void {
  const build = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "unknown";
  const header = document.querySelector("header");
  if (!header || build === "unknown") return;
  const span = document.createElement("span");
  span.className = "build";
  span.textContent = build;
  // Named, because seven hex characters in a header is a mystery otherwise.
  span.title = `SSF Slide Elements was built from commit ${build}`;
  header.append(span);
}

void Office.onReady(() => {
  applyTheme();
  // Before the floor check: a host that cannot run the add-in is exactly the
  // case where somebody needs to say which build refused them.
  showBuild();
  const check = hostReady();
  if (!check.ok) {
    // Said out loud rather than swallowed: a pane that renders a dead UI on an
    // unsupported host gives the user nothing to report.
    const node = root();
    node.textContent = "";
    const p = document.createElement("p");
    p.className = "blocked";
    p.textContent = check.detail;
    node.append(p);
    return;
  }
  draw();
});
