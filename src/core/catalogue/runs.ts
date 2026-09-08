/**
 * Elements that come in several sizes.
 *
 * A run is a set of elements whose English names differ only by one count:
 * "Process flow, horizontal, 3 boxes with table" for 1 to 6. The pane shows a
 * run as one tile with a stepper (`docs/DESIGN.md` section 2). The rule is the
 * name pattern and nothing else: a number that stands alone before a word, never
 * inside "2×2" or "1-2-3", "box" the one irregular plural, and the key-figure
 * flows one run by an explicit pattern because their names carry two counts.
 */
import type { SizeRun } from "./types.js";

const WORDS: Record<string, number> = { one: 1, two: 2, three: 3, four: 4, five: 5, six: 6 };
const KEY_FIGURES = /^Process flow, (\d) box(?:es)?(?:,| and) .*key figures$/;
const COUNTED = /\b(box|row|area|column|level)\b/g;
const NOUNS = ["boxes", "rows", "lines", "levels", "columns", "areas"];

/** Every way to read one counting number out of a name: the key with `N` in its place, and the count. */
export function countKeys(name: string): { key: string; n: number }[] {
  const norm = name.replace(/\b(one|two|three|four|five|six)\b/gi, (w) => String(WORDS[w.toLowerCase()]));
  const figures = KEY_FIGURES.exec(norm);
  if (figures) return [{ key: "Process flow, N boxes with key figures", n: Number(figures[1]) }];
  const out: { key: string; n: number }[] = [];
  const re = /(?<![×\-\d])(\d)(?=\s[a-z])/g;
  let hit: RegExpExecArray | null;
  while ((hit = re.exec(norm)) !== null) {
    const key = (norm.slice(0, hit.index) + "N" + norm.slice(hit.index + 1)).replace(COUNTED, (w) =>
      w === "box" ? "boxes" : `${w}s`,
    );
    out.push({ key, n: Number(hit[1]) });
  }
  return out;
}

/** What a run's stepper counts, read off its key: the noun after `N`, else the first counted noun in the name. */
export function countedNoun(key: string): string {
  const after = /N (\w+)/.exec(key);
  if (after && NOUNS.includes(after[1] ?? "")) return after[1] ?? "items";
  return NOUNS.find((n) => new RegExp(`\\b${n}\\b`, "i").test(key)) ?? "items";
}

/**
 * The runs among a set of names, as a map from each member's name to its run.
 *
 * A run needs at least two members with distinct counts, and no two members
 * with the same count. A name that could belong to two runs joins the larger.
 */
export function sizeRuns(names: string[]): Map<string, SizeRun> {
  const groups = new Map<string, { name: string; n: number }[]>();
  for (const name of names) {
    for (const c of countKeys(name)) {
      const list = groups.get(c.key) ?? [];
      list.push({ name, n: c.n });
      groups.set(c.key, list);
    }
  }
  const best = new Map<string, { run: SizeRun; size: number }>();
  for (const [key, members] of groups) {
    const counts = new Set(members.map((m) => m.n));
    if (counts.size < 2 || counts.size !== members.length) continue;
    const noun = countedNoun(key);
    for (const m of members) {
      const current = best.get(m.name);
      if (!current || current.size < members.length)
        best.set(m.name, { run: { key, noun, count: m.n }, size: members.length });
    }
  }
  return new Map([...best].map(([name, b]) => [name, b.run]));
}
