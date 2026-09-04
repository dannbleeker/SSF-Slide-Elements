import { describe, expect, it } from "vitest";
import "./setup-base64.js";
import { boundsOf, elementId, kindOf, slug, unionBounds } from "../src/core/catalogue/harvest.js";
import { A_NS, P_NS, element, parseXml } from "../src/core/pptx/xml.js";

describe("slug", () => {
  /**
   * Danish letters are transliterated, not dropped. "Hvid kasse på vægt"
   * losing its vowels is `hvid-kasse-p-vgt`, which nobody can match back to a
   * slide when they are reading a bug report — and every name in the shipped
   * library is Danish, so this is the common case.
   */
  it("transliterates Danish letters rather than dropping them", () => {
    expect(slug("Grå kasser")).toBe("graa-kasser");
    expect(slug("Hvid kasse på vægt")).toBe("hvid-kasse-paa-vaegt");
    expect(slug("Ændringslog")).toBe("aendringslog");
  });

  it("never answers an empty string", () => {
    expect(slug("— —")).toBe("unnamed");
  });
});

describe("elementId", () => {
  it("is stable for a name that appears once", () => {
    expect(elementId("procesflow", "Procesflow med 1 kasse", new Map())).toEqual({
      id: "procesflow--procesflow-med-1-kasse",
      repeated: false,
    });
  });

  /**
   * A name ending in a digit is not a duplicate, and a regex over the id said
   * it was: half the library's names end in one ("…med checkliste på 4"), so
   * the first version of this warning fired on them and buried the real
   * duplicate the deck actually contains.
   */
  it("does not call a name ending in a digit a repeat", () => {
    const seen = new Map<string, number>();
    expect(elementId("hvide-kasser", "Hvid kasse med checkliste på 4", seen).repeated).toBe(false);
  });

  it("reports a genuine repeat and suffixes it", () => {
    const seen = new Map<string, number>();
    elementId("procesflow", "Samme navn", seen);
    expect(elementId("procesflow", "Samme navn", seen)).toEqual({
      id: "procesflow--samme-navn-2",
      repeated: true,
    });
  });
});

describe("bounds", () => {
  const sp = (x: number, y: number, cx: number, cy: number): Element =>
    parseXml(
      `<p:sp xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr></p:sp>`,
    ).documentElement;

  it("reads a shape's own transform", () => {
    expect(boundsOf(sp(10, 20, 30, 40))).toEqual({ x: 10, y: 20, cx: 30, cy: 40 });
  });

  it("answers undefined for a shape that states none", () => {
    expect(boundsOf(parseXml(`<p:sp xmlns:p="${P_NS}"/>`).documentElement)).toBeUndefined();
  });

  /**
   * Read from the shape's OWN properties, never by descending. A group's
   * children each carry a transform in the group's child coordinate space, and
   * the first `<a:off>` a descendant search finds belongs to the group only by
   * luck of document order.
   */
  it("does not take a group child's transform for the group's own", () => {
    const grp = parseXml(
      `<p:grpSp xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:grpSpPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="900" cy="900"/><a:chOff x="0" y="0"/><a:chExt cx="900" cy="900"/></a:xfrm></p:grpSpPr><p:sp><p:spPr><a:xfrm><a:off x="500" y="500"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp></p:grpSp>`,
    ).documentElement;
    expect(boundsOf(grp)).toEqual({ x: 100, y: 100, cx: 900, cy: 900 });
  });

  it("unions boxes into the smallest one containing them", () => {
    expect(
      unionBounds([
        { x: 10, y: 10, cx: 10, cy: 10 },
        { x: 50, y: 5, cx: 10, cy: 10 },
      ]),
    ).toEqual({ x: 10, y: 5, cx: 50, cy: 15 });
  });

  it("answers an empty box for no boxes rather than Infinity", () => {
    expect(unionBounds([])).toEqual({ x: 0, y: 0, cx: 0, cy: 0 });
  });
});

describe("kindOf", () => {
  const frame = (uri: string): Element =>
    parseXml(
      `<p:graphicFrame xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:graphic><a:graphicData uri="${uri}"/></a:graphic></p:graphicFrame>`,
    ).documentElement;

  /**
   * Table, chart and SmartArt all arrive as a `<p:graphicFrame>`; which it is
   * shows only in the `uri`. A picker badge that called every one of them
   * "graphic" would be true and useless.
   */
  it("tells a table from a chart from a diagram", () => {
    expect(kindOf(frame("http://schemas.openxmlformats.org/drawingml/2006/table"))).toBe("table");
    expect(kindOf(frame("http://schemas.openxmlformats.org/drawingml/2006/chart"))).toBe("chart");
    expect(kindOf(frame("http://schemas.openxmlformats.org/drawingml/2006/diagram"))).toBe("diagram");
  });

  it("calls a shape with text text, and one without a shape", () => {
    const withText = parseXml(
      `<p:sp xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:txBody><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>`,
    ).documentElement;
    expect(kindOf(withText)).toBe("text");
    expect(kindOf(parseXml(`<p:sp xmlns:p="${P_NS}"/>`).documentElement)).toBe("shape");
  });
});

describe("xml helpers", () => {
  it("finds direct children only, never descendants", async () => {
    const { child, children } = await import("../src/core/pptx/xml.js");
    const doc = parseXml(`<p:spTree xmlns:p="${P_NS}"><p:sp><p:sp/></p:sp><p:sp/></p:spTree>`);
    const tree = element(doc, P_NS, "spTree")!;
    // Three `<p:sp>` in the document, two of them children of the tree.
    expect(children(tree, P_NS, "sp")).toHaveLength(2);
    expect(child(tree, P_NS, "sp")).toBeDefined();
  });
});
