import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, shared with the scripts.
import { idOf, readCache, withoutAddins } from "../scripts/ribbon-cache.mjs";

/**
 * The listing screenshot's one piece of surgery, held to its own rules.
 *
 * `scripts/ribbon-cache.mjs` edits a file belonging to PowerPoint so the
 * AppSource screenshot carries this add-in's ribbon group and not the two
 * sibling projects'. Every claim about the format below was read off the real
 * file on 2026-09-14; this file is where they are checked, because the tool
 * cannot be run in CI and a mistake in it is measured in the only way that
 * matters — by PowerPoint refusing to draw any add-in at all.
 */
const SEP = "\x1e";

/** A cache shaped exactly like the one measured: header, count+first, three more, then blocks. */
const real = [
  "1:911d2b9d99de05ee_LiveId",
  "4:wa104380862+en-US",
  "43ebbbac-44ad-42b2-a582-0ef079093e6c+\\\\AITEST\\OfficeAddins",
  "5eb9457b-eeb7-43e1-b12b-9bf3c01a22f1+\\\\AITEST\\OfficeAddins",
  "b7f6d3a2-4c1e-4e8a-9f2b-7d5c0a1e6f43+\\\\AITEST\\OfficeAddins",
  "2:4",
  "\\\\AITEST\\OfficeAddins",
  "43ebbbac-44ad-42b2-a582-0ef079093e6c",
  "",
  "1:0:1:0en-GB",
].join(SEP);

const MERGE = "43ebbbac-44ad-42b2-a582-0ef079093e6c";
const OURS = "5eb9457b-eeb7-43e1-b12b-9bf3c01a22f1";
const CHARTS = "b7f6d3a2-4c1e-4e8a-9f2b-7d5c0a1e6f43";
const SCRIPT_LAB = "wa104380862";

describe("what the ribbon cache says is installed", () => {
  it("reads the count off the record that also carries the first add-in", () => {
    const { count, entries } = readCache(real) as { count: number; entries: string[] };
    // The count INCLUDES the add-in sharing its record, so four means records
    // 1 through 4. Reading it as "four MORE" is the off-by-one this case exists
    // for: it would drop Script Lab's entry out of the list silently.
    expect(count).toBe(4);
    expect(entries.map(idOf)).toEqual([SCRIPT_LAB, MERGE, OURS, CHARTS]);
  });

  it("names a store add-in and a catalogue add-in the same way", () => {
    expect(idOf("wa104380862+en-US")).toBe(SCRIPT_LAB);
    expect(idOf(`${OURS}+\\\\AITEST\\OfficeAddins`)).toBe(OURS);
    expect(idOf("no-plus-here")).toBe("no-plus-here");
  });

  it("refuses a file that is not a ribbon cache rather than rewriting it", () => {
    expect(() => {
      readCache("just some text");
    }).toThrow(/carries no count/);
    expect(() => {
      readCache(["header", "notanumber:x"].join(SEP));
    }).toThrow(/carries no count|not a count/);
    expect(() => {
      readCache(["header", "9:only+one"].join(SEP));
    }).toThrow(/claims 9 add-ins and holds 1/);
  });
});

describe("taking the sibling add-ins off the ribbon", () => {
  it("removes the two named and lowers the count", () => {
    const out = withoutAddins(real, [MERGE, CHARTS]) as string;
    const { count, entries } = readCache(out) as { count: number; entries: string[] };
    expect(count).toBe(2);
    expect(entries.map(idOf)).toEqual([SCRIPT_LAB, OURS]);
  });

  it("leaves the header and every block after the list untouched", () => {
    const out = (withoutAddins(real, [MERGE, CHARTS]) as string).split(SEP);
    expect(out[0]).toBe("1:911d2b9d99de05ee_LiveId");
    // The blocks are the tail of the file. They still mention the removed
    // add-ins, deliberately: an add-in the LIST does not name is not loaded,
    // and rewriting counts whose meaning is not established is what gets a
    // cache rejected whole.
    expect(out.slice(-5)).toEqual(["2:4", "\\\\AITEST\\OfficeAddins", MERGE, "", "1:0:1:0en-GB"]);
  });

  it("moves the count prefix when the entry carrying it is the one removed", () => {
    // Record 1 is "4:wa104380862+en-US" — the count and an add-in in one
    // string. Drop that add-in and the prefix has to travel to whatever is
    // first afterwards. Written carelessly this produces a cache whose first
    // entry is a bare guid with no count, which is the unreadable case above.
    const out = withoutAddins(real, [SCRIPT_LAB]) as string;
    const { count, entries } = readCache(out) as { count: number; entries: string[] };
    expect(count).toBe(3);
    expect(entries.map(idOf)).toEqual([MERGE, OURS, CHARTS]);
    expect(out.split(SEP)[1]).toBe(`3:${MERGE}+\\\\AITEST\\OfficeAddins`);
  });

  it("does nothing to a cache that does not carry the add-in asked about", () => {
    const out = withoutAddins(real, ["nothing-like-this"]) as string;
    expect(out).toBe(real);
  });

  it("refuses to empty the list, which is the state PowerPoint did not recover from", () => {
    // Measured 2026-09-14: with no add-ins listed, every add-in's ribbon entry
    // went — this project's included — and two restarts did not rebuild them
    // from the install records still on disk. Only a byte-for-byte restore of
    // the cache brought them back. So the tool refuses rather than writing the
    // file that caused it.
    expect(() => {
      withoutAddins(real, [SCRIPT_LAB, MERGE, OURS, CHARTS]);
    }).toThrow(/no add-ins at all/);
  });
});
