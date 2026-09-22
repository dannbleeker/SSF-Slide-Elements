import { describe, expect, it } from "vitest";
import { P_NS, parseXml } from "../src/core/pptx/xml.js";
import {
  boxOf,
  contentCount,
  occupiedBoxes,
  offSlide,
  rotationOf,
  topLevelShapes,
  union,
} from "../src/core/catalogue/boxes.js";

/**
 * The geometry: where a shape sits on its slide, how far round it is turned,
 * and which of a slide's shapes count as already there.
 *
 * `boxes.ts` had no test file of its own. It was reached only through
 * `catalogue.test.ts`, which harvests the owner's real decks and compares the
 * result against the committed index, and through `splice-malformed.test.ts`,
 * which feeds it shapes it cannot measure. Both are worth having and neither
 * pins a NUMBER: the committed index is rounded to four decimals, so an error
 * of one EMU on a 12192000 EMU slide — 8.2e-8 of the width — rounds away
 * before anything compares it, and a malformed shape's answer is `undefined`
 * whatever the arithmetic above it does.
 *
 * The mutation sweep of 2026-09-13 said so plainly: nineteen survivors in this
 * one file, including the `180` that decides whether a shape is rotated at all,
 * the `360` that decides whether it is rotated ENOUGH to record, and four of
 * the `2`s that halve a side to find a centre. A rotation wrong by a degree or
 * a centre off by half a box moves an element on a real slide, and nothing
 * noticed any of them moving.
 *
 * So everything here is called with round numbers and asserted exactly. The
 * slide is 16:9 in EMU, which is what both library decks and most destination
 * decks are.
 */

const W = 12192000;
const H = 6858000;

const NS = [
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

/** The `<p:nvGrpSpPr>` and `<p:grpSpPr>` every `<p:spTree>` opens with. */
const TREE_HEAD =
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;

/** A slide document whose shape tree holds exactly `body`, with the two property children in front. */
const slide = (body: string): Document =>
  parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld ${NS}><p:cSld><p:spTree>${TREE_HEAD}${body}</p:spTree></p:cSld></p:sld>`,
  );

/** The first top-level shape of such a slide. */
const only = (body: string): globalThis.Element => topLevelShapes(slide(body))[0] as globalThis.Element;

/** A rectangle with the given `<a:xfrm>` contents. */
const sp = (xfrm: string, extra = ""): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm ${extra}>${xfrm}</a:xfrm></p:spPr></p:sp>`;

/** A rectangle at a plain position, optionally turned `rot` sixtythousandths of a degree. */
const box = (x: number, y: number, w: number, h: number, rot?: number): string =>
  sp(`<a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/>`, rot === undefined ? "" : `rot="${rot}"`);

describe("reading a shape's frame", () => {
  it("refuses a coordinate that is present but empty", () => {
    // `Number("")` is 0, not NaN, so an empty attribute is the one malformed
    // value that reads as a perfectly good position at the origin. A generator
    // that writes `x=""` for "not laid out yet" would put the shape at the top
    // left of the preview card and count it as content there.
    expect(boxOf(only(sp(`<a:off x="" y="0"/><a:ext cx="10" cy="10"/>`)), W, H), "empty x").toBeUndefined();
    expect(boxOf(only(sp(`<a:off x="0" y=""/><a:ext cx="10" cy="10"/>`)), W, H), "empty y").toBeUndefined();
    expect(boxOf(only(sp(`<a:off x="0" y="0"/><a:ext cx="" cy="10"/>`)), W, H), "empty cx").toBeUndefined();
    expect(boxOf(only(sp(`<a:off x="0" y="0"/><a:ext cx="10" cy=""/>`)), W, H), "empty cy").toBeUndefined();
    // A written zero is a position, not an absence, and must still measure.
    expect(boxOf(only(box(0, 0, 10, 10)), W, H)).toEqual({ x: 0, y: 0, w: 10 / W, h: 10 / H });
  });

  it("measures an unrotated shape as the frame the XML gives, to the EMU", () => {
    expect(boxOf(only(box(1219200, 685800, 6096000, 3429000)), W, H)).toEqual({
      x: 0.1,
      y: 0.1,
      w: 0.5,
      h: 0.5,
    });
  });
});

describe("a table, whose frame lies about its size", () => {
  /**
   * `docs/DESIGN.md` section 3: "A table's box is the sum of its columns and
   * rows, not its frame's `ext`, which PowerPoint ignores when it draws the
   * table." The sums are asserted to the EMU here, because the committed
   * catalogue rounds to four decimals and cannot see a column a unit wide.
   */
  const table = (grid: string, rows: string, w = 10, h = 10): string =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>` +
    `<a:graphic><a:graphicData><a:tbl>${grid}${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

  it("adds its columns and its rows up exactly", () => {
    const grid = `<a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="1000000"/></a:tblGrid>`;
    const rows = `<a:tr h="900000"/><a:tr h="100000"/>`;
    const measured = boxOf(only(table(grid, rows)), W, H);
    expect(measured?.w, "3000000 + 1000000").toBe(4000000 / W);
    expect(measured?.h, "900000 + 100000").toBe(1000000 / H);
  });

  it("adds nothing for a column or a row that does not say how big it is", () => {
    // A `<a:gridCol>` with no `w` is what a table still being laid out looks
    // like. It contributes NOTHING, not a unit: the totals below are the same
    // two totals as above, with an unmeasured column and row alongside.
    const grid = `<a:tblGrid><a:gridCol w="3000000"/><a:gridCol/><a:gridCol w="1000000"/></a:tblGrid>`;
    const rows = `<a:tr h="900000"/><a:tr/><a:tr h="100000"/>`;
    const measured = boxOf(only(table(grid, rows)), W, H);
    expect(measured?.w).toBe(4000000 / W);
    expect(measured?.h).toBe(1000000 / H);
  });

  it("gives a table with no column grid no width of its own", () => {
    // The frame is what is left when there is no `<a:tblGrid>` to add up, and
    // the frame here is empty — the only shape of table that can show what the
    // missing grid contributes, since `Math.max` hides it behind any real
    // frame. The rows still add up, so the case is not vacuous in both halves.
    const measured = boxOf(only(table("", `<a:tr h="500000"/>`, 0, 0)), W, H);
    expect(measured?.w, "no columns, and no frame either").toBe(0);
    expect(measured?.h, "the one row").toBe(500000 / H);
  });

  it("measures a GROUP by its own frame, not by a table nested inside it", () => {
    // A table grouped with its caption — one gesture, and ordinary in a user's
    // deck. `element` walks descendants, so asking any top-level shape for a
    // table found this one and handed back its column widths. Those are in the
    // group's own CHILD coordinate space: `chExt` here is ten times `ext`, the
    // kind of scale a group routinely carries, so the sum below came out ten
    // times too large and `Math.max` took it for the group's width. The preview
    // card then drew a grey box wider than the slide it sits on.
    const grouped =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="Table and caption"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="1000000" y="500000"/><a:ext cx="4000000" cy="1000000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="40000000" cy="10000000"/></a:xfrm></p:grpSpPr>` +
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="31" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/>` +
      `</p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="40000000" cy="10000000"/></p:xfrm>` +
      `<a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="40000000"/></a:tblGrid>` +
      `<a:tr h="10000000"/></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:grpSp>`;
    const measured = boxOf(only(grouped), W, H);
    expect(measured?.w, "the group's own ext, not the table's columns").toBe(4000000 / W);
    expect(measured?.h, "the group's own ext, not the table's rows").toBe(1000000 / H);
  });
});

describe("a rotated shape's extent", () => {
  /**
   * `docs/DESIGN.md` section 3: "The box of a rotated shape is its rotated
   * extent, not the unrotated frame the XML gives." The angle is stored in
   * sixtythousandths of a degree, so `rot="5400000"` is a quarter turn.
   */
  it("turns a 4 by 2 frame a quarter turn into a 2 by 4 extent about the same centre", () => {
    // Round numbers on purpose: about the centre (3000000, 1500000), a
    // 4000000 x 2000000 rectangle turned 90° is a 2000000 x 4000000 one, so
    // every number below can be checked by hand. Its top edge leaves the slide,
    // which is what happens to a wide shape stood on end.
    const measured = boxOf(only(box(1000000, 500000, 4000000, 2000000, 90 * 60000)), W, H);
    expect(measured?.w, "the frame's height becomes the extent's width").toBeCloseTo(2000000 / W, 9);
    expect(measured?.h, "and its width the extent's height").toBeCloseTo(4000000 / H, 9);
    expect(measured?.x, "centre 3000000 less half of 2000000").toBeCloseTo(2000000 / W, 9);
    expect(measured?.y, "centre 1500000 less half of 4000000").toBeCloseTo(-500000 / H, 9);
  });

  it("treats a half turn as no rotation at all, and one degree past it as rotation", () => {
    // 180° is the boundary: a half turn gives back the frame it started from,
    // so the arithmetic is skipped. One degree further and it is a real
    // rotation — the extent grows on both axes, by cos(1°) and sin(1°) of the
    // sides. The pair is what pins the `180`; either alone is satisfied by a
    // check against any number at all.
    const half = boxOf(only(box(1000000, 500000, 4000000, 2000000, 180 * 60000)), W, H);
    expect(half, "a half turn is the frame, exactly").toEqual({
      x: 1000000 / W,
      y: 500000 / H,
      w: 4000000 / W,
      h: 2000000 / H,
    });
    const past = boxOf(only(box(1000000, 500000, 4000000, 2000000, 181 * 60000)), W, H);
    expect(past?.w, "4000000·cos1° + 2000000·sin1°").toBeCloseTo(4034295.59 / W, 6);
    expect(past?.h, "4000000·sin1° + 2000000·cos1°").toBeCloseTo(2069505.02 / H, 6);
  });
});

describe("the unrotated frame the preview cut masks to", () => {
  /**
   * `rotationOf` answers the other half of the same question (`docs/DESIGN.md`
   * section 3): the extent cannot be un-rotated back into the frame it came
   * from, so the frame is recorded beside it. The owner's two stamps are the
   * only elements that have one, at −29.06° and 35.02°.
   */
  it("gives the angle and the frame it was measured from", () => {
    const spun = rotationOf(only(box(1219200, 685800, 6096000, 3429000, 29.06 * 60000)), W, H);
    expect(spun?.deg).toBeCloseTo(29.06, 10);
    expect(spun?.frame, "the frame, not the extent").toEqual({ x: 0.1, y: 0.1, w: 0.5, h: 0.5 });
  });

  it("says nothing about a shape that is not turned, or has no frame to turn", () => {
    // Three ways of not being rotated, and they are not the same line: no
    // `rot` at all, a `rot` of zero, and a FULL turn, which is a shape somebody
    // dragged all the way round and is the same picture as no turn at all.
    expect(rotationOf(only(box(0, 0, 10, 10)), W, H), "no rot attribute").toBeUndefined();
    expect(rotationOf(only(box(0, 0, 10, 10, 0)), W, H), "rot of zero").toBeUndefined();
    expect(rotationOf(only(box(0, 0, 10, 10, 360 * 60000)), W, H), "a full turn").toBeUndefined();
    const bare = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Bare"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>`;
    expect(rotationOf(only(bare), W, H), "no frame at all").toBeUndefined();
  });

  it("records the smallest turn either side of those", () => {
    // One degree, and one short of a full turn: the boundary on both sides of
    // the two zeroes above. A degree is visible on a slide — it is about 5 mm
    // of drop across the width of a 33 cm one.
    expect(rotationOf(only(box(0, 0, 10, 10, 1 * 60000)), W, H)?.deg).toBeCloseTo(1, 10);
    expect(rotationOf(only(box(0, 0, 10, 10, 359 * 60000)), W, H)?.deg).toBeCloseTo(359, 10);
  });
});

describe("the smallest box around several", () => {
  it("is the one box itself when there is only one", () => {
    // Eighths and quarters, so the one box comes back bit for bit rather than
    // through a subtraction that rounds.
    const one = { x: 0.25, y: 0.5, w: 0.125, h: 0.0625 };
    expect(union([one])).toEqual(one);
  });

  it("is the whole slide when there are none", () => {
    // Nothing to enclose, so the caller gets a rectangle it can draw rather
    // than an undefined it has to test for.
    expect(union([])).toEqual({ x: 0, y: 0, w: 1, h: 1 });
  });

  it("reaches the leftmost, topmost, rightmost and bottommost edges", () => {
    const boxes = [
      { x: 0.5, y: 0.5, w: 0.2, h: 0.2 },
      { x: 0.1, y: 0.6, w: 0.2, h: 0.3 },
      { x: 0.4, y: 0.2, w: 0.5, h: 0.1 },
    ];
    expect(union(boxes)).toEqual({ x: 0.1, y: 0.2, w: 0.9 - 0.1, h: 0.9 - 0.2 });
  });
});

describe("a shape that is off the slide", () => {
  /**
   * The 4:3 deck's instruction boxes sit past the right edge, and a
   * destination deck can carry the same. `offSlide` is what keeps them out of
   * the preview card and out of `contentCount`, and its edge is the edge of the
   * slide: a shape whose left edge is exactly at the right edge is off.
   */
  it("is off at the edge and on a hair inside it", () => {
    expect(offSlide({ x: 1, y: 0, w: 0.1, h: 0.1 }), "left edge on the right edge").toBe(true);
    expect(offSlide({ x: 0, y: 1, w: 0.1, h: 0.1 }), "top edge on the bottom edge").toBe(true);
    expect(offSlide({ x: 0.9999, y: 0, w: 0.1, h: 0.1 }), "a hair inside the right edge").toBe(false);
    expect(offSlide({ x: 0, y: 0.9999, w: 0.1, h: 0.1 }), "a hair above the bottom edge").toBe(false);
    // Hanging over an edge is not being off it: half a shape is still visible.
    expect(offSlide({ x: 0.95, y: 0.95, w: 0.5, h: 0.5 })).toBe(false);
  });
});

describe("walking a slide's top-level shapes", () => {
  it("starts at the tree's first child, wherever the properties are", () => {
    // The walk is over `childNodes` by index and the two property children are
    // skipped by NAME, not by position — so it must begin at the beginning. A
    // tree whose first child is a shape loses its bottom-most one otherwise,
    // and z-order is the order the preview card paints in.
    const two =
      `<p:sld ${NS}><p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Under"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Over"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
      `</p:spTree></p:cSld></p:sld>`;
    const shapes = topLevelShapes(parseXml(two));
    const named = shapes.map((s) => s.getElementsByTagNameNS(P_NS, "cNvPr")[0]?.getAttribute("name"));
    expect(named, "both of them, in z-order").toEqual(["Under", "Over"]);
  });

  it("leaves the tree's own two property children out", () => {
    expect(topLevelShapes(slide(box(0, 0, 10, 10)))).toHaveLength(1);
  });
});

describe("what a slide already holds", () => {
  /**
   * The ghost rule (`docs/DESIGN.md` section 6): the insert removes the
   * "Click to add text" placeholders it lands over, so an EMPTY one is not
   * something the element would land on top of. A placeholder with anything at
   * all in it is content — and "anything at all" is one character, one table or
   * one picture, which is what the three cases below are.
   */
  const ph = (inner: string, type = ' type="body"'): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="6" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph${type} idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="6096000" cy="3429000"/></a:xfrm></p:spPr>` +
    `${inner}</p:sp>`;

  it("does not count an empty placeholder, which the insert takes away", () => {
    const empty = ph(`<p:txBody><a:bodyPr/><a:p/></p:txBody>`);
    expect(contentCount(slide(empty), W, H)).toBe(0);
    expect(occupiedBoxes(slide(empty), W, H)).toEqual([]);
  });

  it("counts the running furniture as furniture, not as content", () => {
    /**
     * A footer, a slide number and a date are on nearly every corporate slide,
     * and none of them is something the user put there for this deck. They
     * carry text — a company name, "2", today's date — so the empty-placeholder
     * rule does not reach them, and they were counted.
     *
     * What that decides is the "Move to a new slide" offer: `held` is what
     * `moveableAfter` reads, so an otherwise EMPTY slide with a footer and a
     * slide number on it reported two things already there, and the pane
     * offered to move the element off a slide that has nothing on it but its
     * own furniture.
     *
     * `src/core/pptx/layout.ts` already knew this — it keeps `TITLES` and
     * `CHROME` apart and excludes both — and `harvest.ts` has its own copy of
     * the same set. This file had neither.
     */
    const furniture =
      ph(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>ACME A/S</a:t></a:r></a:p></p:txBody>`, ' type="ftr"') +
      ph(
        `<p:txBody><a:bodyPr/><a:p><a:fld id="{1}" type="slidenum"><a:t>2</a:t></a:fld></a:p></p:txBody>`,
        ' type="sldNum"',
      ) +
      ph(
        `<p:txBody><a:bodyPr/><a:p><a:fld id="{2}" type="datetime1"><a:t>22-09-2026</a:t></a:fld></a:p></p:txBody>`,
        ' type="dt"',
      );
    expect(contentCount(slide(furniture), W, H), "a slide with only furniture on it is empty").toBe(0);
  });

  it("still counts a body placeholder beside the furniture", () => {
    // The pair, so "ignore the furniture" cannot quietly become "ignore
    // placeholders".
    const mixed =
      ph(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>ACME A/S</a:t></a:r></a:p></p:txBody>`, ' type="ftr"') +
      ph(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>Real content</a:t></a:r></a:p></p:txBody>`);
    expect(contentCount(slide(mixed), W, H)).toBe(1);
  });

  it("counts a placeholder holding a single character", () => {
    // One character is content. The line reads `length > 0`, and a `length > 1`
    // would drop the shortest label a slide can carry — a "1", an "A", a "%".
    const typed = ph(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody>`);
    expect(contentCount(slide(typed), W, H)).toBe(1);
    expect(occupiedBoxes(slide(typed), W, H)).toEqual([{ x: 0.1, y: 0.1, w: 0.5, h: 0.5 }]);
  });

  it("counts a placeholder holding a table, which has no text of its own", () => {
    // A content placeholder a table was dropped into: `<a:graphic>` and not one
    // `<a:t>` anywhere. Counting it as an empty ghost would offer to move an
    // element off a slide that has a table on it, or draw no grey box for it.
    const held =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/><p:cNvGraphicFramePr/>` +
      `<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></p:xfrm>` +
      `<a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="4000000"/></a:tblGrid>` +
      `<a:tr h="1000000"/></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    expect(contentCount(slide(held), W, H)).toBe(1);
    expect(occupiedBoxes(slide(held), W, H)).toHaveLength(1);
  });

  it("counts a placeholder holding a picture, which has no graphic of its own", () => {
    // The other half of the same `or`: a picture is an `<a:blip>` inside a
    // `<p:blipFill>`, with no `<a:graphic>` and no text. Both arms are needed,
    // and a slide with a photograph on it is the ordinary case for this one.
    const held =
      `<p:pic><p:nvPicPr><p:cNvPr id="8" name="Picture"/><p:cNvPicPr/>` +
      `<p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></a:xfrm></p:spPr></p:pic>`;
    expect(contentCount(slide(held), W, H)).toBe(1);
    expect(occupiedBoxes(slide(held), W, H)).toHaveLength(1);
  });

  it("does not count the title, which an insert never removes", () => {
    const title = ph(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>A heading</a:t></a:r></a:p></p:txBody>`, ' type="title"');
    expect(contentCount(slide(title), W, H)).toBe(0);
    // It is still drawn on the card, because it is still on the slide.
    expect(occupiedBoxes(slide(title), W, H)).toHaveLength(1);
  });
});
