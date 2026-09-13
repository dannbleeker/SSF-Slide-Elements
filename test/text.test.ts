import { describe, expect, it } from "vitest";
import { paragraphsOf, partName } from "../src/core/catalogue/text.js";
import { A_NS, P_NS, parseXml } from "../src/core/pptx/xml.js";

/**
 * Text read out of DrawingML: which paragraphs count, and how long a part's
 * name is allowed to be.
 *
 * `test/catalogue.test.ts` holds the naming RULES — brackets, the colon, the
 * fall back to the shape name, the numbering of duplicates. What is here is
 * what a mutation sweep found unheld on 2026-09-13: the empty-paragraph filter
 * and the length limit, both of them boundaries that no case sat either side
 * of.
 */

/** A `<p:sp>` whose text body carries these paragraphs, given as their inner XML. */
function shape(...paragraphs: string[]): Element {
  const doc = parseXml(
    `<p:sp xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:txBody>${paragraphs
      .map((inner) => `<a:p>${inner}</a:p>`)
      .join("")}</p:txBody></p:sp>`,
  );
  return doc.documentElement;
}

/** One paragraph holding one run of this text. */
const run = (text: string) => `<a:r><a:t>${text}</a:t></a:r>`;

describe("paragraphsOf", () => {
  it("drops the empty paragraphs and keeps a one-character one", () => {
    // Two boundaries on one line. A paragraph with NO text and one whose text
    // is only whitespace both measure 0 and go; a paragraph of a single
    // character measures 1 and stays — a lone "3" or "%" is a real label in
    // the owner's decks, and dropping it would silently rename its part.
    expect(paragraphsOf(shape("", run("   "), run("3"), run("Kicker box")))).toEqual(["3", "Kicker box"]);
  });
});

describe("partName's length limit", () => {
  // 40 characters, a literal in `partName` with no comment and no line in
  // `docs/DESIGN.md` behind it — the design record fixes the naming RULES
  // (section "Names") and is silent on the length. So this pins the number
  // that exists rather than arguing for a better one: 40 through, 41 cut.
  const named = (n: number) => "Indholdselementet".padEnd(n, "e");
  const nameOf = (text: string) => partName(shape(run(text)), "S", new Set<string>(), { count: 0 });

  it("passes a name of exactly 40 characters through whole", () => {
    const forty = named(40);
    expect(forty).toHaveLength(40);
    expect(nameOf(forty)).toBe(forty);
  });

  it("cuts a name of 41 characters and marks the cut", () => {
    const fortyOne = named(41);
    expect(fortyOne).toHaveLength(41);
    expect(nameOf(fortyOne)).toBe(`${named(40)}…`);
  });
});
