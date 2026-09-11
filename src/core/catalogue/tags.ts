/**
 * Tags derived from the deck's names and categories.
 *
 * Derived, not authored: the owner turned an authored vocabulary down
 * (`docs/DESIGN.md` section 16). A tag is a word people filter by, so the list
 * is the words the library's own names use, mapped to English.
 *
 * The words are mostly Danish because the keys mostly are: a whole-slide
 * element is keyed by its slide title, and titles stay Danish
 * (`docs/DESIGN.md` section 2). A PART is keyed by its own text on the slide,
 * and that text is English since 2026-09-11 — so any rule a part can match
 * needs BOTH spellings. `dokument|document` is the one that actually bites;
 * `proces` already matches "process" as a substring, and `decision` was
 * already there beside `beslutning`. The rest are reached through the
 * category, which is a heading slide's title and so still Danish.
 */

const TAGS: [RegExp, string][] = [
  [/kasse/, "boxes"],
  [/streg/, "lines"],
  [/tabel/, "table"],
  [/procesflow|proces/, "process"],
  [/hierarki/, "hierarchy"],
  [/matrix/, "matrix"],
  [/trekant/, "triangle"],
  [/kausalitet/, "causality"],
  [/pil|peger/, "arrows"],
  [/vægt/, "weighing"],
  [/nøgletal|kpi/, "figures"],
  [/checkliste/, "checklist"],
  [/citat/, "quotes"],
  [/billed/, "pictures"],
  [/kommentering/, "commentary"],
  [/stempl|markering|etiket/, "stamps"],
  [/flowchart/, "flowchart"],
  [/ikon/, "icons"],
  [/agenda|ændringslog|decision|beslutning|dokument|document/, "meeting"],
  [/vertikal/, "columns"],
  [/horisontal|rækker/, "rows"],
  [/grå/, "grey"],
  [/hvid/, "white"],
  [/sort/, "black"],
  [/one-page|problem statement|årsplan|bridging/, "one-pager"],
];

/** The tags for an element, from its Danish key and category; a part also gets `small`. */
export function tagsFor(key: string, category: string, part = false): string[] {
  const text = `${key} ${category}`.toLowerCase();
  const out: string[] = [];
  for (const [pattern, tag] of TAGS) if (pattern.test(text) && !out.includes(tag)) out.push(tag);
  if (part && !out.includes("small")) out.push("small");
  return out;
}
