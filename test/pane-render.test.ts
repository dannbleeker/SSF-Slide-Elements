/**
 * @vitest-environment jsdom
 */
import { beforeEach, describe, expect, it } from "vitest";
import { render } from "../src/pane/render.js";
import { EMPTY, STEP_TITLE, blockedReason, primary } from "../src/pane/steps.js";

/**
 * The pane's DOM, in the one environment the suite has for it.
 *
 * jsdom has no layout and no colour, so nothing here can say the pane LOOKS
 * right — that is what `scripts/pane-shots.mjs` is for. What it can pin is
 * what is drawn, in what order, and that a string from outside reaches the
 * page as text.
 */
let root: HTMLElement;

beforeEach(() => {
  document.body.innerHTML = '<header><b>SSF</b><span>Slide Elements</span></header><div id="pane"></div>';
  root = document.getElementById("pane") as HTMLElement;
});

describe("the start step", () => {
  it("draws the heading, the reason, and exactly one primary button, last", () => {
    render(root, EMPTY, "start");
    expect(root.querySelector("h1")?.textContent).toBe(STEP_TITLE.start);
    expect(root.querySelector("p.blocked")?.textContent).toBe(blockedReason(EMPTY, "start"));
    const buttons = root.querySelectorAll("button");
    expect(buttons.length).toBe(1);
    const button = buttons[0] as HTMLButtonElement;
    expect(button.className).toBe("primary");
    expect(button.textContent).toBe(primary(EMPTY, "start").label);
    expect(button.disabled).toBe(true);
    expect(button.dataset["action"]).toBe("insert");
    // The last element in the view, which is the layout rule.
    const main = root.querySelector("main") as HTMLElement;
    expect(main.lastElementChild).toBe(button);
  });

  it("spends the orange budget once", () => {
    render(root, EMPTY, "start");
    expect(root.querySelectorAll(".tick").length).toBe(1);
  });

  it("draws no notice when there is none", () => {
    render(root, EMPTY, "start");
    expect(root.querySelector(".notice")).toBeNull();
  });

  it("puts a notice on screen as TEXT, never as markup", () => {
    // A notice comes from the host, and a host echoes its argument back. If
    // that argument ever contains markup it must not become an element.
    render(root, { notice: '<img src=x onerror="alert(1)">' }, "start");
    expect(root.querySelector("img")).toBeNull();
    expect(root.querySelector(".notice")?.textContent).toBe('<img src=x onerror="alert(1)">');
  });

  it("replaces rather than appends on a redraw", () => {
    render(root, EMPTY, "start");
    render(root, { notice: "again" }, "start");
    expect(root.querySelectorAll("main").length).toBe(1);
    expect(root.querySelectorAll("button").length).toBe(1);
  });
});
