/**
 * The pane, rendered and MEASURED at both ends of a task pane's width.
 *
 * The suite pins the picker's behaviour in jsdom, which has no layout and no
 * colour — so a rule about how the pane LOOKS is invisible there unless
 * somebody turns it into something countable. This renders every state at 320
 * and 512, in both themes, and measures the two things a screenshot cannot
 * show by itself.
 *
 *   npm run build
 *   npx vite preview --port 4178 --strictPort &
 *   node scripts/pane-shots.mjs
 *
 * **Horizontal overflow.** Every long string on this screen comes from outside
 * it — an element's own eighty-character title, an error PowerPoint wrote. An
 * unbroken one takes the pane past its own width and puts content off the side.
 *
 * **Labels that spill their box.** The first pass caught a real one: the card
 * label's `-webkit-line-clamp` computed to an inert `flow-root` and the text
 * was hard-clipped mid-line, with no ellipsis, in the BUILT pane. It looked
 * like the minifier dropping a property and it was not. Nothing in jsdom could
 * have seen it, and it is the reason `clampName` truncates in the string.
 *
 * Exits non-zero on a finding, and writes the PNGs either way.
 */
import { mkdirSync } from "node:fs";
import { chromium } from "playwright";

mkdirSync("pane-shots", { recursive: true });
const findings = [];
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
/**
 * The two ends of the width a task pane is dragged between, in both themes.
 * 320 is the narrowest PowerPoint will make it and the width every layout bug
 * shows up at first; 512 is where the grid gains its third column.
 */
const VIEWS = [
  { name: "light-320", viewport: { width: 320, height: 1100 }, colorScheme: "light" },
  { name: "dark-512", viewport: { width: 512, height: 1100 }, colorScheme: "dark" },
];

for (const view of VIEWS) {
  const name = view.name;
  const opts = { viewport: view.viewport, colorScheme: view.colorScheme };
  const ctx = await b.newContext(opts);
  const p = await ctx.newPage();
  const errs = [];
  p.on("pageerror", (e) => errs.push(String(e)));
  await p.goto("http://localhost:4178/taskpane.html", { waitUntil: "networkidle" });
  await p.waitForSelector(".card", { timeout: 15000 });
  const state = await p.evaluate(() => ({
    cards: document.querySelectorAll(".card").length,
    sections: document.querySelectorAll(".section h2").length,
    svgs: document.querySelectorAll(".card .thumb svg").length,
    overflow: document.documentElement.scrollWidth - document.documentElement.clientWidth,
    firstSection: document.querySelector(".section h2")?.textContent,
    spilled: [...document.querySelectorAll(".card .label")].filter(
      (l) => l.scrollHeight > Math.ceil(l.getBoundingClientRect().height) + 1,
    ).length,
  }));
  console.log(name, JSON.stringify(state), errs.length ? errs : "");
  if (state.overflow > 0) findings.push(`${name}: the pane is ${state.overflow}px wider than its frame`);
  if (state.spilled > 0) findings.push(`${name}: ${state.spilled} card label(s) spill their box`);
  if (errs.length > 0) findings.push(`${name}: ${errs.join("; ")}`);
  await p.screenshot({ path: `pane-shots/pane-${name}.png` });
  await p.fill("#q", "trekant");
  await p.waitForTimeout(250);
  console.log(
    "   search 'trekant':",
    JSON.stringify(
      await p.evaluate(() => ({
        cards: document.querySelectorAll(".card").length,
        count: document.querySelector(".count")?.textContent,
      })),
    ),
  );
  await p.screenshot({ path: `pane-shots/pane-${name}-search.png` });
  await ctx.close();
}
await b.close();
if (findings.length > 0) {
  console.error(`\n${findings.length} finding(s):`);
  for (const f of findings) console.error(`  ${f}`);
  process.exit(1);
}
console.log("\nno overflow, no spilled labels");
