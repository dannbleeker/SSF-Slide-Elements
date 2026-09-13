import { describe, expect, it } from "vitest";
import { countKeys, sizeRuns } from "../src/core/catalogue/runs.js";

/**
 * Size runs: the numbers in the name pattern, and the tie-break between two
 * runs a name could join.
 *
 * `test/catalogue.test.ts` already holds what a run IS — which names group,
 * which do not, and what the stepper counts. What is here is the part a
 * mutation sweep found unheld on 2026-09-13: the spelled-out number table, and
 * the comparison that picks between two runs of the SAME size.
 */

describe("the spelled-out counts", () => {
  // `WORDS` is a table of six literals, and the sweep killed only the middle
  // three: nothing pinned `one`, `five` or `six`, so `one: 2` went unnoticed.
  // Each word is asserted against the digit it stands for, one row per literal,
  // because a table is only held by reading every row of it.
  it.each([
    ["one", 1],
    ["two", 2],
    ["three", 3],
    ["four", 4],
    ["five", 5],
    ["six", 6],
  ])("reads %s as %i", (word, n) => {
    expect(countKeys(`Matrix, ${word} row, boxes with arrows`)).toEqual([
      { key: "Matrix, N rows, boxes with arrows", n },
    ]);
  });

  it("reads a spelled-out count whatever its case", () => {
    expect(countKeys("Matrix, Six rows")).toEqual([{ key: "Matrix, N rows", n: 6 }]);
  });
});

describe("a name that fits two runs of the same size", () => {
  // The larger run wins (`test/catalogue.test.ts` holds that). When the two are
  // the SAME size there is nothing to be larger, and the code keeps whichever
  // run it saw FIRST — `current.size < members.length` is a strict `<`, so an
  // equal-sized later run does not displace it. Group order is the order the
  // keys were first seen, which is the order of `countKeys` on the first name
  // that produced them, so "rows" beats "columns" here for a reason a reader
  // can follow rather than by luck of a hash.
  //
  // This is a recording of what the code does, not an argument that first-seen
  // is the better run. Both are two-member runs offering the same stepper
  // depth; what matters is that the answer is DETERMINISTIC, and this pins it.
  const runs = sizeRuns(["Grid, 2 rows and 2 columns", "Grid, 3 rows and 2 columns", "Grid, 2 rows and 3 columns"]);

  it("keeps the run it found first", () => {
    expect(runs.get("Grid, 2 rows and 2 columns")).toEqual({
      key: "Grid, N rows and 2 columns",
      noun: "rows",
      count: 2,
    });
  });

  it("still gives the run it did not join its own other member", () => {
    expect(runs.get("Grid, 2 rows and 3 columns")).toEqual({
      key: "Grid, 2 rows and N columns",
      noun: "columns",
      count: 3,
    });
  });
});
