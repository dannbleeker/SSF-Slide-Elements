#!/usr/bin/env node
/**
 * Read an answer sheet the probe produced and say what it means.
 *
 * The snippet that runs inside PowerPoint collects raw observations and makes
 * no judgements. Every reading happens here, through `src/host/probe.ts`,
 * which is covered by tests. A probe that reasons inside the host is a probe
 * whose conclusions nobody can check.
 *
 *   npm run build:lib
 *   node scripts/read-answers.mjs sheet.json
 *   node scripts/read-answers.mjs first.json second.json --save
 *
 * Two sheets, in the order they were taken, answer the Ctrl+Z question: the
 * first leaves a tagged slide behind, the second looks for it. With `--save`
 * every sheet given is filed under `docs/host-answers/`, stamped with when it
 * was taken.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import {
  exportPartsVerdict,
  floorLine,
  insertVerdict,
  insertionBlame,
  leftBehind,
  masterVerdict,
  notAsked,
  orderVerdict,
  pruningReading,
  selectionVerdict,
  targetAddedVerdict,
  timingLine,
  undoVerdict,
} from "../dist-lib/host/probe.js";

const files = process.argv.slice(2).filter((a) => !a.startsWith("--"));
if (files.length === 0) {
  console.error("usage: node scripts/read-answers.mjs <sheet.json> [<second-sheet.json>] [--save]");
  process.exit(1);
}

/** The snippet prints markers around the JSON so it can be copied out of a console. Tolerate a paste that still has them. */
function sheetOf(file) {
  const raw = readFileSync(file, "utf8");
  return JSON.parse(raw.slice(raw.indexOf("{"), raw.lastIndexOf("}") + 1));
}

const sheets = files.map(sheetOf);
const sheet = sheets[sheets.length - 1];
const previous = sheets.length > 1 ? sheets[sheets.length - 2] : undefined;

const line = (label, value) => console.log(`  ${label.padEnd(24)} ${value}`);

console.log(`\nSSF Slide Elements answer sheet — ${sheet.takenAt ?? "no timestamp"}`);
line("platform", sheet.platform ?? "unknown");
line("host version", sheet.host ?? "unknown");
line("PowerPointApi", (sheet.requirementSets ?? []).join(", ") || "none reported");
line("deck", `${sheet.deckAtStart} slides at start, ${sheet.deckAtEnd} at end, ${sheet.totalMs ?? "?"} ms in all`);
if (sheet.undoAtStart?.removed) line("previous run's slide", sheet.undoAtStart.removed);

// The own-deck arms insert the whole package, so what they should land is the
// number of slides that package holds — read from its part list, not guessed.
const ownSlides = sheet.partsAtStart?.slides ?? sheet.deckAtStart;
const insert = (arm, expected, what) =>
  arm && arm.before !== undefined ? insertVerdict({ ...arm, expected }) : notAsked(what);

const own = insert(sheet.insertOwn, ownSlides, "own-deck control arm");
const ownDest = insert(sheet.insertOwnDestTheme, ownSlides, "own-deck UseDestinationTheme arm");
const listed = insert(sheet.insertListed, 2, "listed two-slide arm");
const pruned = insert(sheet.insertPruned, 1, "pruned arm");
const unlisted = insert(sheet.insertUnlisted, 1, "unlisted arm");

console.log("\n1. Does the host accept a package pruned to one slide with the other's parts still present?");
line("control: own deck", `${own.verdict} — ${own.detail} (${sheet.insertOwn?.ms ?? "?"} ms)`);
line("own, dest theme", `${ownDest.verdict} — ${ownDest.detail} (${sheet.insertOwnDestTheme?.ms ?? "?"} ms)`);
line("listed, two slides", `${listed.verdict} — ${listed.detail}`);
line("pruned (id + rel)", `${pruned.verdict} — ${pruned.detail}`);
line("unlisted (id only)", `${unlisted.verdict} — ${unlisted.detail}`);
console.log(`\n  => ${insertionBlame(listed.verdict, own.verdict)}`);
console.log(`  => ${pruningReading(listed, pruned, unlisted)}`);
{
  const keep = masterVerdict({ ...(sheet.insertOwn ?? {}), landed: own.landed, formatting: "KeepSourceFormatting" });
  const dest = masterVerdict({
    ...(sheet.insertOwnDestTheme ?? {}),
    landed: ownDest.landed,
    formatting: "UseDestinationTheme",
  });
  line("masters, keep source", `${keep.verdict} — ${keep.detail}`);
  line("masters, dest theme", `${dest.verdict} — ${dest.detail}`);
  line("swept after own", sheet.sweepOwn ?? "not reported");
  line("swept after dest", sheet.sweepOwnDestTheme ?? "not reported");
  line("swept after pruning", sheet.sweepPruning ?? "not reported");
}

console.log("\n2. Does insert-after-current, then a positional delete of current, keep the order?");
{
  const order = sheet.order ?? {};
  const v = orderVerdict(order);
  line("verdict", `${v.verdict} — ${v.detail}`);
  const t = targetAddedVerdict(order);
  line("just-added as target", `${t.verdict} — ${t.detail}`);
  if (order.afterFirst) line("after first insert", order.afterFirst.join(", "));
  if (order.afterSecond) line("after second insert", order.afterSecond.join(", "));
  if (order.afterDelete) line("after the delete", order.afterDelete.join(", "));
  line("swept", order.sweep ?? "not reported");
}

console.log("\n3. Does getSelectedSlides name the slide you were on, at the file's position?");
{
  const v = selectionVerdict(sheet.selection ?? {});
  line("verdict", `${v.verdict} — ${v.detail}`);
  const s = sheet.selection ?? {};
  if (s.apiIdsHead) line("API ids, head", s.apiIdsHead.join(", "));
  if (s.sldIdLstHead) line("file sldIdLst, head", s.sldIdLstHead.join(", "));
  if (s.apiIdsHaveHash === false) line("note", "the API's ids carry no '#', so the file comparison used the whole id");
}

console.log("\n4. Which read of the deck, and what does each drop?");
{
  const v = exportPartsVerdict(sheet.exportParts ?? {});
  line("verdict", `${v.verdict} — ${v.detail}`);
  const missing = sheet.exportParts?.missing ?? [];
  if (missing.length > 0) line("not carried over", missing.join(", "));
  const added = sheet.exportParts?.addedByExport ?? [];
  if (added.length > 0) line("added by the export", added.join(", "));
}

console.log("\n5. Does PowerPoint's own Ctrl+Z revert the insert?");
{
  const v = undoVerdict({
    foundAtStart: sheet.undoAtStart?.found,
    leftBehind: sheet.undo?.leftBehind,
    deckAtStart: sheet.deckAtStart,
    previousDeckAtEnd: previous?.deckAtEnd,
    error: sheet.undoAtStart?.error,
  });
  line("verdict", `${v.verdict} — ${v.detail}`);
  if (sheet.undo?.insert) {
    const u = insertVerdict({ ...sheet.undo.insert, expected: 1 });
    line("slide left behind", `${u.verdict} — ${u.detail}`);
  }
}

console.log("\n6. How long does a read take, and is the floor met?");
line("getFileAsync", timingLine("getFileAsync", sheet.ownRead));
line("getFileAsync, at end", timingLine("getFileAsync", sheet.ownReadAtEnd));
line("export", timingLine("exportAsBase64Presentation", sheet.exportParts));
line("floor", floorLine(sheet.requirementSets ?? [], sheet.platform, sheet.floor ?? "1.2"));

console.log(`\nclean-up: ${sheet.sweep ?? "not reported"}`);
line(
  "package delta",
  leftBehind(sheet.partsAtStart, sheet.partsAtEnd) +
    (sheet.undo?.leftBehind ? " One of the slides is the one left for Ctrl+Z." : ""),
);
const expectedAtEnd = sheet.deckAtStart + (sheet.undo?.leftBehind ? 1 : 0);
if (sheet.deckAtEnd !== expectedAtEnd) {
  console.log(
    `  WARNING: the deck ended at ${sheet.deckAtEnd} and should be ${expectedAtEnd} (${sheet.deckAtStart} at start${sheet.undo?.leftBehind ? " plus the slide left for Ctrl+Z" : ""}). Check what was left behind.`,
  );
}

if (process.argv.includes("--save")) {
  mkdirSync("docs/host-answers", { recursive: true });
  for (const s of sheets) {
    const stamp = (s.takenAt ?? "unknown").replace(/[:.]/g, "-");
    const out = `docs/host-answers/${stamp}.json`;
    writeFileSync(out, `${JSON.stringify(s, null, 2)}\n`);
    console.log(`\nsaved ${out}`);
  }
}
console.log();
