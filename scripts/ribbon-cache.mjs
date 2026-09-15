/**
 * PowerPoint's record of which add-ins put buttons on its ribbon.
 *
 * Needed for one job: the **AppSource screenshot**. The listing shot has to
 * show this add-in, and the machine that takes it also carries the two sibling
 * projects, whose ribbon groups would otherwise sit in a picture on a public
 * store page. `docs/LISTING.md` says what the shot must and must not contain.
 *
 * ## The file
 *
 * `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\AppCommands\18.0\PowerPoint.RibbonCache.<locale>`
 * — a few hundred bytes of records separated by 0x1E. Read on 2026-09-14:
 *
 *     0   "1:911d2b9d99de05ee_LiveId"                 a header
 *     1   "4:wa104380862+en-US"                        COUNT, then the first add-in
 *     2   "43ebbbac-…+\\AITEST\OfficeAddins"           SSF Merge
 *     3   "5eb9457b-…+\\AITEST\OfficeAddins"           SSF Slide Elements
 *     4   "b7f6d3a2-…+\\AITEST\OfficeAddins"           SSF Charts
 *     5+  per-add-in blocks, then one for the store add-in
 *
 * The count on record 1 INCLUDES the add-in on record 1, so the entries occupy
 * records 1 through `count`. The blocks after them are left alone: an add-in
 * the list does not name is not loaded whether or not a block survives, and
 * rewriting counts whose meaning is not established is how a cache gets
 * rejected whole.
 *
 * ## Why not the Office dialog
 *
 * Because it will not open. Measured 2026-09-14, four attempts across two
 * routes: Home → Add-ins → More Add-ins never produced the Office Add-ins
 * dialog under a driven click, and File → Options never opened either. The
 * add-ins are sideloaded from a shared-folder catalog, so they do not appear in
 * the ribbon flyout's own short list at all.
 *
 * ## What NOT to do, measured
 *
 * Deleting this file removes **every** add-in's ribbon entry, including this
 * project's, and PowerPoint did not rebuild the entry from the install records
 * that were still on disk — two restarts, nothing. The way back was a
 * byte-for-byte copy of the cache. So: **edit it, never delete it, and keep a
 * backup**. Removing the manifests from the catalog on their own does nothing;
 * the cache is what the ribbon is built from.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { isMain } from "./is-main.mjs";

/** The separator between records. */
const SEP = "\x1e";

/**
 * The add-ins this cache says are installed, and where the list sits.
 *
 * @param {string} text
 * @returns {{ records: string[], count: number, entries: string[] }}
 */
export function readCache(text) {
  const records = text.split(SEP);
  const head = records[1] ?? "";
  const colon = head.indexOf(":");
  if (colon < 1) throw new Error("this is not a ribbon cache: record 1 carries no count");
  const count = Number(head.slice(0, colon));
  if (!Number.isInteger(count) || count < 1) throw new Error(`the count is not a count: ${JSON.stringify(head)}`);
  if (records.length < count + 1) throw new Error(`the cache claims ${count} add-ins and holds ${records.length - 1}`);
  const entries = [head.slice(colon + 1), ...records.slice(2, count + 1)];
  return { records, count, entries };
}

/**
 * The add-in id an entry names — a GUID, or a store id like `wa104380862`.
 *
 * @param {string} entry
 * @returns {string}
 */
export function idOf(entry) {
  const plus = entry.indexOf("+");
  return plus === -1 ? entry : entry.slice(0, plus);
}

/**
 * The same cache with the named add-ins taken out of the installed list.
 *
 * The count moves with them, and the `N:` prefix travels to whichever entry is
 * first afterwards — dropping the entry that carried the prefix is the case
 * this gets wrong if written carelessly.
 *
 * @param {string} text
 * @param {readonly string[]} ids add-in ids to remove
 * @returns {string}
 */
export function withoutAddins(text, ids) {
  const { records, count, entries } = readCache(text);
  const kept = entries.filter((entry) => !ids.includes(idOf(entry)));
  if (kept.length === 0) {
    // Measured 2026-09-14: with no add-ins listed, PowerPoint drew no add-in
    // ribbon entries at all and did not rebuild them on the next two starts.
    // Refusing is better than writing the file that caused that.
    throw new Error("that would leave no add-ins at all, which is the state PowerPoint did not recover from");
  }
  const first = kept[0] ?? "";
  return [records[0] ?? "", `${kept.length}:${first}`, ...kept.slice(1), ...records.slice(count + 1)].join(SEP);
}

/**
 * `node scripts/ribbon-cache.mjs <cacheFile> [id,id,…]`
 *
 * With one argument it LISTS what the ribbon is built from; with two it writes
 * the file back without those add-ins. **Copy the file first** — the header
 * says why.
 */
function main(argv) {
  const [file, drop] = argv;
  if (!file) {
    console.error("usage: node scripts/ribbon-cache.mjs <PowerPoint.RibbonCache.xx-XX> [id,id,…]");
    return 2;
  }
  const text = readFileSync(file, "latin1");
  const { entries } = readCache(text);
  if (!drop) {
    console.log(`${entries.length} add-in(s) on the ribbon:`);
    for (const entry of entries) console.log(`  ${idOf(entry)}`);
    return 0;
  }
  const ids = drop.split(",").filter(Boolean);
  const out = withoutAddins(text, ids);
  writeFileSync(file, out, "latin1");
  const left = readCache(out).entries.map(idOf);
  console.log(`removed ${entries.length - left.length}; ${left.length} left: ${left.join(", ")}`);
  return 0;
}

if (isMain(import.meta.url)) {
  process.exitCode = main(process.argv.slice(2));
}
