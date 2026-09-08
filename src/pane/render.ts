/**
 * The pane's DOM, built from the decisions in `steps.ts`.
 *
 * Every string a user sees reaches the page through `textContent`; nothing here
 * assigns markup, so a notice the host wrote cannot become an element. The
 * layout rules the SSF pane was approved on — one column, one orange tick, one
 * primary control drawn last — are what this file holds, and
 * `scripts/pane-shots.mjs` is what measures them.
 */
import { STEP_TITLE, blockedReason, primary, type PaneState, type StepId } from "./steps.js";

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

/**
 * Draw the pane into `root`, replacing whatever was there.
 *
 * Replaced, never appended: a redraw that appended would show two panes, and
 * a redraw is the ordinary way state reaches the screen.
 */
export function render(root: HTMLElement, state: PaneState, step: StepId): void {
  root.textContent = "";
  const main = el("main");
  main.append(el("span", "tick"));
  main.append(el("h1", undefined, STEP_TITLE[step]));
  main.append(el("p", "blocked", blockedReason(state, step)));
  if (state.notice) main.append(el("p", "notice", state.notice));

  const action = primary(state, step);
  const button = el("button", "primary", action.label);
  button.type = "button";
  button.dataset["action"] = "insert";
  button.disabled = action.disabled;
  // Last, always: one thing to press per screen, and it is at the bottom.
  main.append(button);
  root.append(main);
}
