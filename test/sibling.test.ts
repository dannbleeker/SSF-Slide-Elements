import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

// @ts-expect-error — plain .mjs with no types, and deliberately the SAME module
// the weekly sweep runs. Two copies of the triage table is how a claim quietly
// stops matching its check.
import { SOURCES, TRIAGED, VERDICTS, idOf, tableDeclared, tableKeys } from "../scripts/sibling-watch.mjs";
// @ts-expect-error — the same module; a second line only because one line of nine names is too long.
import { tablesFrom, untriaged, verdictOf } from "../scripts/sibling-watch.mjs";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { withoutTsProse } from "../scripts/without-prose.mjs";

/**
 * The borrowed-fact guard.
 *
 * Everything this repo knew about the PowerPoint host when this guard was
 * written had been learned by a sibling project and hand-copied here. Some of
 * it is first-hand now — the probe's four sheets of 2026-09-10 and the product
 * round of 2026-09-11, both on the web — and none of it is first-hand for
 * Windows, Mac or iPad. `docs/SIBLING.md` is the ledger and says which is
 * which; this is the half of it a machine can hold.
 *
 * What rots is narrow. A single run's observation ("a by-id clean-up reported
 * 45 deletes and removed nothing") is true forever and needs nothing. A COUNT
 * OF ROUNDS is a live counter by definition: the sibling runs more of them.
 * SSF-Merge had four source comments saying "174 consecutive archived rounds",
 * all correct on the morning they were written, all wrong the moment round 175
 * ran, with nothing anywhere to say so.
 *
 * So: a round count carries the date it was taken, or it is not a round count.
 *
 * **Scoped to source and scripts, never to docs, and that is deliberate.**
 * `docs/SIBLING.md` and `CLAUDE.md` state the rule, which means they quote the
 * exact sentences that break it.
 */

/** This file, which is the one place in `test/` whose SUBJECT is the rule. */
const SELF = join("test", "sibling.test.ts");

/** Every .ts and .mjs file under the directories that hold justifying comments. */
function sourceFiles(): string[] {
  const roots = ["src", "scripts", "test"];
  const out: string[] = [];
  const walk = (dir: string): void => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) walk(path);
      else if (/\.(ts|mjs)$/.test(path) && path !== SELF) out.push(path);
    }
  };
  for (const root of roots) walk(root);
  return out;
}

/**
 * A file as one string, with comment prefixes stripped, plus where each line
 * began. A line-by-line scan cannot see a claim that wraps across two comment
 * lines, which is exactly the shape that rotted on the sibling.
 */
function joined(text: string): { body: string; starts: number[] } {
  const lines = text.split("\n").map((l) => l.replace(/^\s*(?:\*|\/\/)\s?/, ""));
  const starts: number[] = [];
  let at = 0;
  for (const line of lines) {
    starts.push(at);
    at += line.length + 1;
  }
  return { body: lines.join(" "), starts };
}

function lineAt(starts: number[], offset: number): number {
  let lo = 0;
  for (let i = 0; i < starts.length; i++) if ((starts[i] ?? 0) <= offset) lo = i;
  return lo + 1;
}

/**
 * A claim about how many ROUNDS a sibling has run. Deliberately not "any
 * number near the word sibling": `45 successful deletes` and `37 generated
 * slides` are single observations that will read correctly in ten years.
 */
const ROUND_COUNT =
  /\b(\d+|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|seventeen|twenty)\s+(?:consecutive\s+)?(?:archived\s+)?(?:real-host\s+)?rounds\b/gi;

/** An ISO date, which is how this repo stamps a measurement. */
const DATED = /\b20\d\d-\d\d-\d\d\b/;

/** Lines either side of a claim that may carry its date. */
const WINDOW = 6;

describe("a borrowed round count carries its date", () => {
  it("finds no undated round count in any source comment", () => {
    const offenders: string[] = [];
    for (const file of sourceFiles()) {
      const text = readFileSync(file, "utf8");
      const lines = text.split("\n");
      const { body, starts } = joined(text);
      ROUND_COUNT.lastIndex = 0;
      for (let m = ROUND_COUNT.exec(body); m; m = ROUND_COUNT.exec(body)) {
        const line = lineAt(starts, m.index);
        const near = lines.slice(Math.max(0, line - 1 - WINDOW), line + WINDOW).join("\n");
        if (DATED.test(near)) continue;
        offenders.push(`${file}:${line}  …${m[0]}…`);
      }
    }
    expect(offenders, `undated round counts (see docs/SIBLING.md):\n${offenders.join("\n")}`).toEqual([]);
  });

  it("catches the shape that actually rotted", () => {
    // The comment that was in a sibling's powerpoint.ts until 2026-08-27,
    // WRAPPED AS IT WAS: the claim lives only across the join.
    const rotted = [
      " * in every round of its self-test battery, and across **174 consecutive",
      " * archived rounds every rung answered**, in 550-710ms, with zero refusals",
    ].join("\n");
    ROUND_COUNT.lastIndex = 0;
    expect(ROUND_COUNT.test(joined(rotted).body)).toBe(true);
    expect(DATED.test(rotted)).toBe(false);
    const fixed = rotted + "\n * and zero silences. Measured 2026-08-27.";
    expect(DATED.test(fixed)).toBe(true);
  });

  it("leaves a single run's observation alone", () => {
    for (const durable of [
      "a sibling project's by-id clean-up once reported 45 successful deletes",
      "a sibling add-in logged 46 `InvalidParam passed to GetItem(id)` failures in one run",
      "a sibling project put 37 generated slides ahead of somebody's title slide",
      "a sibling add-in pages every collection read at 20 for that reason",
    ]) {
      ROUND_COUNT.lastIndex = 0;
      expect(ROUND_COUNT.test(durable), durable).toBe(false);
    }
  });
});

describe("the ledger is wired in", () => {
  const ledger = readFileSync("docs/SIBLING.md", "utf8");

  it("is reachable from the places somebody starts reading", () => {
    for (const [file, why] of [
      ["CLAUDE.md", "the memory file's host rules"],
      ["README.md", "the docs table"],
      ["CONTRIBUTING.md", "the note before touching PowerPoint"],
    ] as const) {
      expect(readFileSync(file, "utf8"), why).toContain("docs/SIBLING.md");
    }
  });

  it("answers every row, including the ones that are no exposure", () => {
    expect(ledger).toMatch(/no exposure/i);
  });

  it("says out loud that reading it by hand needs all three repositories", () => {
    // The precondition that made this transferable at all, and the one a
    // session holding only this repo cannot discover for itself.
    expect(ledger).toMatch(/both sibling repositories checked out in the same session/i);
  });
});

describe("the sibling sweep", () => {
  const table = readFileSync("scripts/sibling-watch.mjs", "utf8");

  it("pulls the keys out of a table, quoted or bare", () => {
    const src = [
      "export const T = {",
      '  bare: "a",',
      '  "quoted": "b",',
      '  "with-dashes": "c",',
      '  under_scored: "d",',
      '  "dotted.id": "e",',
      "};",
      "",
    ].join("\n");
    expect(tableKeys(src, "T")).toEqual(["bare", "quoted", "with-dashes", "under_scored", "dotted.id"]);
  });

  it("reads a quoted key that carries a colon, which is how SSF-Merge spells its rows", () => {
    // The sibling's own reader answered `issue` for `"issue:1650":` and
    // stopped, so watching its table needed this.
    const src = [
      "export const TRIAGED = {",
      '  "issue:1650":',
      '    "ADOPTED — a slide add.",',
      '  "question:getitemat-past-end": "RELEVANT: both.",',
      "};",
      "",
    ].join("\n");
    expect(tableKeys(src, "TRIAGED")).toEqual(["issue:1650", "question:getitemat-past-end"]);
  });

  it("says a missing table is MISSING, never empty", () => {
    expect(tableKeys("export const OTHER = {\n  a: 1,\n};\n", "GONE")).toBeNull();
    expect(() => {
      tablesFrom((): string => "export const NOTHING = {\n};\n");
    }).toThrow(/renamed, moved, emptied or reformatted/);
  });

  it("says a table it can no longer READ is missing too", () => {
    // A reindent upstream parses to nothing, which must not read as quiet.
    const reindented = ["export const T = {", '    bare: "a",', "};", ""].join("\n");
    expect(tableKeys(reindented, "T"), "four spaces is a table this cannot read").toBeNull();
    expect(() => {
      tablesFrom((): string => reindented.replace("T", "FAKE_BASELINE"));
    }).toThrow(/emptied or reformatted/);
  });

  it("reads a register that is declared and empty as empty, when the source says it may be", () => {
    // SSF-Charts' PENDING_QUESTIONS is empty by design between a probe's
    // commit and its answering round; the sibling's watcher has been red on it
    // since 2026-09-02. An ABSENT table is still broken, for every source.
    const pending = (SOURCES as { table: string; mayBeEmpty?: boolean }[]).find((s) => s.table === "PENDING_QUESTIONS");
    expect(pending?.mayBeEmpty).toBe(true);
    const others = (SOURCES as { table: string; mayBeEmpty?: boolean }[]).filter(
      (s) => s.table !== "PENDING_QUESTIONS",
    );
    expect(
      others.every((s) => !s.mayBeEmpty),
      "only the register may be empty",
    ).toBe(true);

    const full = (table: string) => `export const ${table} = {\n  "some-key": "x",\n};\n`;
    const declaredEmpty = "export const PENDING_QUESTIONS = {\n};\n";
    const read = (_repo: string, path: string): string =>
      path.endsWith("host-baseline.mjs")
        ? [full("FAKE_BASELINE"), full("KNOWN_DIVERGENCES"), full("UNSTABLE_ANSWERS"), declaredEmpty].join("\n")
        : path.endsWith("office-js-watch.mjs")
          ? full("KNOWN_ISSUES")
          : full("TRIAGED").replace('"some-key"', '"issue:1"');
    const tables = tablesFrom(read) as { source: { table: string }; keys: string[] }[];
    expect(tables.find((t) => t.source.table === "PENDING_QUESTIONS")?.keys).toEqual([]);
    expect(() => {
      tablesFrom((repo: string, path: string): string =>
        read(repo, path).replace("export const PENDING_QUESTIONS", "export const GONE"),
      );
    }).toThrow(/PENDING_QUESTIONS/);
    expect(tableDeclared(declaredEmpty, "PENDING_QUESTIONS")).toBe(true);
  });

  it("names a finding by whose it is, never by where it was read", () => {
    // A Charts finding read from Merge's table is the same finding.
    expect(idOf({ space: "charts", kind: "issue" }, "1650")).toBe("charts:issue:1650");
    expect(idOf({ space: "charts", kind: "self" }, "issue:1650")).toBe("charts:issue:1650");
    const found = untriaged(
      [
        { source: { repo: "charts", space: "charts", table: "KNOWN_ISSUES", kind: "issue" }, keys: ["1650"] },
        {
          source: { repo: "merge", space: "charts", table: "TRIAGED", kind: "self" },
          keys: ["issue:1650", "question:fresh"],
        },
      ],
      {},
    ) as { id: string; tables: string[] }[];
    expect(found.map((f) => f.id)).toEqual(["charts:issue:1650", "charts:question:fresh"]);
    expect(found[0]?.tables).toEqual(["charts/KNOWN_ISSUES", "merge/TRIAGED"]);
  });

  it("reports a finding with no row, and keeps which tables it was in", () => {
    const tables = [
      {
        source: { repo: "charts", space: "charts", table: "FAKE_BASELINE", kind: "question" },
        keys: ["known", "fresh"],
      },
      { source: { repo: "charts", space: "charts", table: "UNSTABLE_ANSWERS", kind: "question" }, keys: ["fresh"] },
    ];
    const found = untriaged(tables, { "charts:question:known": "NO EXPOSURE — nothing here." });
    expect(found).toEqual([
      {
        id: "charts:question:fresh",
        space: "charts",
        kind: "question",
        key: "fresh",
        tables: ["charts/FAKE_BASELINE", "charts/UNSTABLE_ANSWERS"],
      },
    ]);
  });

  it("cannot be fooled by a key inherited from Object.prototype", () => {
    const found = untriaged(
      [
        {
          source: { repo: "charts", space: "charts", table: "T", kind: "question" },
          keys: ["constructor", "__proto__"],
        },
      ],
      {},
    ) as { key: string }[];
    expect(found.map((f) => f.key)).toEqual(["constructor", "__proto__"]);
  });

  it("opens every reason with a verdict from the closed vocabulary", () => {
    const unparsed = Object.entries(TRIAGED).filter(([, why]) => verdictOf(why) === undefined);
    expect(
      unparsed.map(([k]) => k),
      "a reason must open with one of " + VERDICTS.join(", "),
    ).toEqual([]);
  });

  it("gives every finding a reason, not just a verdict", () => {
    const bare = Object.entries(TRIAGED as Record<string, string>).filter(
      ([, why]) => why.replace(verdictOf(why) ?? "", "").replace(/^[\s—-]+/, "").length < 8,
    );
    expect(bare.map(([k]) => k)).toEqual([]);
  });

  it("keys everything by sibling, then issue: or question:", () => {
    const odd = Object.keys(TRIAGED).filter((k) => !/^(charts|merge):(issue:\d+|question:[a-z0-9-]+)$/.test(k));
    expect(odd).toEqual([]);
    expect(Object.keys(TRIAGED).length, "the table has been emptied").toBeGreaterThan(50);
  });

  it("keeps the ledger from falling behind the table", () => {
    // Every finding we ACTED on has to appear in the prose a human reads; the
    // ones that are no exposure do not, or the ledger becomes a list of shrugs.
    const ledger = readFileSync("docs/SIBLING.md", "utf8");
    const missing = Object.entries(TRIAGED as Record<string, string>)
      .filter(([, why]) => verdictOf(why) !== "NO EXPOSURE")
      .map(([key]) => key)
      .filter((key) => !ledger.includes(key.slice(key.lastIndexOf(":") + 1)));
    expect(missing, "acted on but absent from docs/SIBLING.md").toEqual([]);
  });

  it("reads raw files and never runs the siblings' code", () => {
    expect(table).toContain("raw.githubusercontent.com");
    expect(table).toContain("dannbleeker/SSF-Charts");
    expect(table).toContain("dannbleeker/SSF-Merge");
    expect(withoutTsProse(table)).not.toMatch(/\bimport\s*\(/);
  });
});
