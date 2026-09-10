import { readFileSync } from "node:fs";
import { beforeAll, describe, expect, it } from "vitest";
import { Pkg, harvest } from "../src/core/index.js";
import type { Catalogue, Element as CatalogueElement, Names, SlideSize } from "../src/core/index.js";
import { A_NS, MC_NS, P_NS, R_NS, child, elements, parseXml, serializeXml } from "../src/core/pptx/xml.js";
import {
  applyMove,
  emptyBodyPlaceholders,
  groupShapes,
  groupable,
  highestShapeId,
  parseFragment,
  rectOf,
  relIdsIn,
  renumber,
  repoint,
  shapesOf,
  slideShapes,
  topLevel,
  unionOf,
  unplaceholder,
} from "../src/core/splice/shapes.js";

/**
 * An element's shapes, made safe to live in somebody else's slide.
 *
 * Two halves, the same shape as `test/catalogue.test.ts`. The synthetic
 * fragments below hold each rewrite to a number in isolation — what a renumber
 * reaches, what a repoint is anchored on, what a scale does to a table and what
 * it must not do to a group — because every one of those is markup PowerPoint
 * either opens or calls damaged, and a screenshot cannot tell the difference.
 * The sweep at the end holds the one claim `shapes.ts` makes about the LIBRARY
 * rather than about a rule: that all 234 elements in the two committed decks
 * parse standalone, which the module's own comment says this file checks.
 *
 * Every builder here declares its prefixes on the shape element itself, exactly
 * as the harvest's serialiser does, because `parseFragment` wraps the markup in
 * a root that declares nothing. That is not decoration: `@xmldom/xmldom` throws
 * `NamespaceError` on a prefix nothing bound, so a fixture leaning on a
 * declaration from a wrapper would not parse at all — which is precisely what
 * makes the sweep a real gate rather than a formality.
 */

/** The prefixes a top-level shape out of the harvest declares for itself. */
const NS = [
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"',
].join(" ");

type Box = [x: number, y: number, cx: number, cy: number];

/** An `<a:xfrm>`: the spelling everything but a graphic frame uses. */
function xfrm(box: Box, extra = ""): string {
  return `<a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/>${extra}</a:xfrm>`;
}

/** An ordinary shape, whose rectangle lives under `<p:spPr>`. */
function sp(id: number, box?: Box, inner = ""): string {
  return (
    `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="${id}" name="Rektangel ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${box ? xfrm(box) : ""}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${inner}</p:sp>`
  );
}

/** A picture, whose blip names an image relationship with `r:embed`. */
function pic(id: number, rId: string, box: Box = [0, 0, 100, 100]): string {
  return (
    `<p:pic ${NS}><p:nvPicPr><p:cNvPr id="${id}" name="Billede ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm(box)}</p:spPr></p:pic>`
  );
}

/**
 * A group, whose rectangle lives under `<p:grpSpPr>` and which states a CHILD
 * coordinate space beside it.
 *
 * `chOff`/`chExt` are deliberately different numbers from `off`/`ext` here, so
 * a pass that scaled all four would be visible rather than plausible.
 */
function grp(id: number, box: Box, childBox: Box, ...kids: string[]): string {
  const space = `<a:chOff x="${childBox[0]}" y="${childBox[1]}"/><a:chExt cx="${childBox[2]}" cy="${childBox[3]}"/>`;
  return (
    `<p:grpSp ${NS}><p:nvGrpSpPr><p:cNvPr id="${id}" name="Gruppe ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr>${xfrm(box, space)}</p:grpSpPr>${kids.join("")}</p:grpSp>`
  );
}

/**
 * A table, whose rectangle is a `<p:xfrm>` in the PresentationML namespace and
 * whose real size is the grid inside it.
 */
function table(id: number, box: Box, cols: number[], rows: number[]): string {
  return (
    `<p:graphicFrame ${NS}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Tabel ${id}"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>` +
    cols.map((w) => `<a:gridCol w="${w}"/>`).join("") +
    `</a:tblGrid>` +
    rows.map((h) => `<a:tr h="${h}"><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>`).join("") +
    `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/**
 * A modern chart as PowerPoint writes one: the frame in `<mc:Choice>` and a
 * PICTURE of it in `<mc:Fallback>`, each carrying a `<p:cNvPr id>` of its own.
 *
 * 46 of the 234 committed elements are shaped like this, so the id inside the
 * fallback is not an exotic case: it is a collision waiting on a slide the user
 * already had a picture on.
 */
function alternate(choiceId: number, fallbackId: number, rId = "rId9"): string {
  return (
    `<mc:AlternateContent ${NS}><mc:Choice Requires="p14">` +
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${choiceId}" name="Diagram"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="urn:chartex"/></a:graphic></p:graphicFrame></mc:Choice>` +
    `<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="${fallbackId}" name="Diagram"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/></p:blipFill><p:spPr>${xfrm([0, 0, 100, 100])}</p:spPr>` +
    `</p:pic></mc:Fallback></mc:AlternateContent>`
  );
}

/** A placeholder on the destination slide: a claim, and whatever the user has typed into it. */
function placeholder(id: number, options: { type?: string; text?: string; noBody?: boolean } = {}): string {
  const type = options.type === undefined ? "" : ` type="${options.type}"`;
  const runs = options.text === undefined ? "" : `<a:r><a:rPr lang="da-DK"/><a:t>${options.text}</a:t></a:r>`;
  const body = options.noBody ? "" : `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runs}</a:p></p:txBody>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Pladsholder ${id}"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph${type} idx="${id}"/></p:nvPr></p:nvSpPr><p:spPr/>${body}</p:sp>`
  );
}

/** The destination slide's shape tree, its own properties in front of the shapes. */
function tree(...shapes: string[]): Element {
  return parseXml(`<p:spTree ${NS}><p:nvGrpSpPr/><p:grpSpPr/>${shapes.join("")}</p:spTree>`).documentElement;
}

/** The `<a:xfrm>` a group states its own frame in. */
function groupFrame(group: Element): Element | undefined {
  const props = child(group, P_NS, "grpSpPr");
  return props ? child(props, A_NS, "xfrm") : undefined;
}

/** One `<a:off>`-shaped child of an `<a:xfrm>`, as the two numbers it states. */
function pair(frame: Element | undefined, local: string, first: string, second: string): (string | null)[] {
  const el = frame ? child(frame, A_NS, local) : undefined;
  return [el?.getAttribute(first) ?? null, el?.getAttribute(second) ?? null];
}

/** Every `<p:cNvPr id>` in a tree, in document order: what a renumber has to have reached. */
function shapeIds(root: Element): string[] {
  return elements(root, P_NS, "cNvPr").map((el) => el.getAttribute("id") ?? "");
}

/** Every value of one attribute, wherever it appears under `root`. */
function valuesOf(root: Element, ns: string, local: string, attr: string): string[] {
  return elements(root, ns, local).map((el) => el.getAttribute(attr) ?? "");
}

describe("an element's shapes, parsed", () => {
  it("hands back every top-level shape in slide order", () => {
    const fragment = parseFragment(sp(4) + pic(5, "rId1") + table(6, [0, 0, 10, 10], [10], [10]));
    expect(topLevel(fragment).map((s) => s.localName)).toEqual(["sp", "pic", "graphicFrame"]);
  });

  it("counts the shapes and not the whitespace and comments between them", () => {
    // The harvest joins serialised shapes, and a serialiser is free to put a
    // newline between them. Counting nodes rather than elements makes an
    // element's shape count depend on how it was written out.
    const fragment = parseFragment(`<!-- Kasse -->\n  ${sp(1)}\n  ${sp(2)}\n`);
    expect(topLevel(fragment)).toHaveLength(2);
  });

  it("answers those same shapes through the name the splice adopts them by", () => {
    // `shapesOf` is what the splice imports into the destination document, so
    // it and `topLevel` disagreeing is an element that lands half-spliced.
    const fragment = parseFragment(sp(1) + sp(2));
    expect(shapesOf(fragment)).toEqual(topLevel(fragment));
  });

  it("refuses markup naming a prefix nothing declared, which is what makes the sweep below a gate", () => {
    /**
     * The wrapper declares nothing on purpose, so an element that leant on a
     * declaration from the library slide's root cannot be parsed at all. If
     * this ever stopped throwing, the sweep over all 234 elements would pass
     * for every element in every state, including the broken ones.
     */
    expect(() => parseFragment('<p:sp><p:nvSpPr><p:cNvPr id="1" name="x"/></p:nvSpPr></p:sp>')).toThrow(
      /NamespaceError|prefix/,
    );
  });
});

describe("ids the destination slide is not already using", () => {
  it("reads the highest id from inside a group, not only from the shapes on top", () => {
    /**
     * Ids are unique across the SLIDE, and a group's children carry them too.
     * A renumber starting above the top-level maximum alone collides with the
     * inside of a group the user already had — two shapes on one id, which is
     * a slide PowerPoint opens with one of them missing.
     */
    const spTree = tree(sp(7), grp(9, [0, 0, 10, 10], [0, 0, 10, 10], sp(42)));
    expect(highestShapeId(spTree)).toBe(42);
  });

  it("ignores an id too large to count exactly, so the number it answers is a real one", () => {
    // `max + 1 === max` above 2^53, so a deck carrying such an id would make
    // every renumbered shape collide with itself. A shape with no id at all is
    // the same question from the other side.
    const spTree = tree(sp(11), `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="9007199254740993" name="huge"/></p:nvSpPr></p:sp>`);
    expect(highestShapeId(spTree)).toBe(11);
    expect(highestShapeId(tree(`<p:sp ${NS}><p:nvSpPr><p:cNvPr name="nameless"/></p:nvSpPr></p:sp>`))).toBe(0);
  });

  it("gives every copy its own id, including a group's children and the picture in an mc:Fallback", () => {
    /**
     * The two places a `<p:cNvPr>` hides from a pass that walks only the
     * top-level shapes. Both are ordinary in the committed library — 84
     * elements carry a group and 46 carry an alternate-content fallback — and
     * an id left as the library wrote it collides with the destination's own.
     */
    const fragment = parseFragment(grp(1, [0, 0, 10, 10], [0, 0, 10, 10], sp(2), sp(3)) + alternate(4, 5));
    const next = renumber(fragment, 100);
    const ids = topLevel(fragment).flatMap((shape) => shapeIds(shape));
    expect(ids).toEqual(["100", "101", "102", "103", "104"]);
    expect(next, "the first id still free, so a second element can carry on from it").toBe(105);
  });

  it("answers the same first free id when the fragment has nothing to renumber", () => {
    const fragment = parseFragment(`<p:sp ${NS}><p:spPr/></p:sp>`);
    expect(renumber(fragment, 60)).toBe(60);
  });
});

describe("relationship references", () => {
  it("rewrites r:embed, r:id and r:link alike, because it is anchored on the namespace", () => {
    /**
     * The defect this shape exists to prevent: a rewriter listing the attribute
     * names it has seen silently skips the first one it has not, and the file
     * then names a relationship the destination has not got. `r:embed` is a
     * blip's image, `r:id` is a chart or an embedded object, `r:link` is a
     * LINKED picture — three spellings of one thing, and a library element uses
     * whichever the owner's drawing needed.
     */
    const fragment = parseFragment(
      `<p:pic ${NS}><p:blipFill><a:blip r:embed="rId1" r:link="rId2"/></p:blipFill>` +
        `<a:graphic><a:graphicData><p:oleObj r:id="rId3"/></a:graphicData></a:graphic></p:pic>`,
    );
    const changed = repoint(
      fragment,
      new Map([
        ["rId1", "rId40"],
        ["rId2", "rId41"],
        ["rId3", "rId42"],
      ]),
    );
    expect(changed).toBe(3);
    const shape = topLevel(fragment)[0] as Element;
    const blip = elements(shape, A_NS, "blip")[0] as Element;
    expect(blip.getAttributeNS(R_NS, "embed")).toBe("rId40");
    expect(blip.getAttributeNS(R_NS, "link")).toBe("rId41");
    expect(elements(shape, P_NS, "oleObj")[0]?.getAttributeNS(R_NS, "id")).toBe("rId42");
  });

  it("leaves a reference the map does not name exactly as it was", () => {
    /**
     * Not an oversight. An element naming a relationship the catalogue did not
     * carry is a defect in the HARVEST, and leaving the id alone makes it
     * visible to the integrity check as an unresolvable reference rather than
     * hiding it behind a plausible-looking one.
     */
    const fragment = parseFragment(pic(1, "rId7") + pic(2, "rId8"));
    expect(repoint(fragment, new Map([["rId7", "rId70"]]))).toBe(1);
    const embeds = topLevel(fragment).map((s) => elements(s, A_NS, "blip")[0]?.getAttributeNS(R_NS, "embed"));
    expect(embeds).toEqual(["rId70", "rId8"]);
  });

  it("leaves an ordinary attribute alone even when its value reads like a relationship id", () => {
    // The other half of "anchored on the namespace": a shape NAMED rId1 is not
    // a reference to anything, and a rewriter matching on the value would
    // rename the user's shape.
    const fragment = parseFragment(`<p:sp ${NS}><p:nvSpPr><p:cNvPr id="3" name="rId1"/></p:nvSpPr></p:sp>`);
    expect(repoint(fragment, new Map([["rId1", "rId99"]]))).toBe(0);
    expect(shapeIds(topLevel(fragment)[0] as Element)).toEqual(["3"]);
    expect(elements(topLevel(fragment)[0] as Element, P_NS, "cNvPr")[0]?.getAttribute("name")).toBe("rId1");
  });

  it("names every relationship the fragment uses, once each and in the order it uses them", () => {
    // What the splice asks before it copies anything: two shapes sharing a
    // picture must ask for it once, or the destination gets two copies of it.
    const fragment = parseFragment(pic(1, "rId5") + pic(2, "rId5") + alternate(3, 4, "rId2"));
    expect(relIdsIn(fragment)).toEqual(["rId5", "rId2"]);
    expect(relIdsIn(parseFragment(sp(1)))).toEqual([]);
  });
});

describe("where a shape says it is", () => {
  it("reads the rectangle of a shape, a group and a graphic frame, which spell it three different ways", () => {
    /**
     * The graphic frame is the one that gets missed: its `<p:xfrm>` is in the
     * PresentationML namespace and sits directly on the shape, where the other
     * two carry a DrawingML `<a:xfrm>` under their properties. 136 of the 234
     * committed elements carry a graphic frame, so a reader that knows only the
     * first two spellings mislands more than half the library.
     */
    const fragment = parseFragment(
      sp(1, [10, 20, 30, 40]) + grp(2, [50, 60, 70, 80], [0, 0, 1, 1]) + table(3, [90, 100, 110, 120], [110], [120]),
    );
    expect(topLevel(fragment).map((s) => rectOf(s))).toEqual([
      { x: 10, y: 20, cx: 30, cy: 40 },
      { x: 50, y: 60, cx: 70, cy: 80 },
      { x: 90, y: 100, cx: 110, cy: 120 },
    ]);
  });

  it("has no rectangle for a shape that does not state a whole one", () => {
    // A text box on a layout placeholder states no box at all and inherits one,
    // and half an `<a:xfrm>` is no more usable than none: guessing the missing
    // half is how an element lands somewhere nobody asked for.
    const halves = parseFragment(
      sp(1) +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:off x="1" y="2"/></a:xfrm></p:spPr></p:sp>` +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:off x="x" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>`,
    );
    expect(topLevel(halves).map((s) => rectOf(s))).toEqual([undefined, undefined, undefined]);
  });

  it("unions the shapes into the element's own frame, stepping over one that states no box", () => {
    const fragment = parseFragment(sp(1, [100, 100, 100, 100]) + sp(2) + sp(3, [400, 50, 100, 100]));
    expect(unionOf(topLevel(fragment))).toEqual({ x: 100, y: 50, cx: 400, cy: 150 });
  });

  it("has no frame at all when nothing in the run states one", () => {
    expect(unionOf([])).toBeUndefined();
    expect(unionOf(topLevel(parseFragment(sp(1))))).toBeUndefined();
  });
});

describe("moving the shapes to where they landed", () => {
  const FROM = { x: 1000, y: 2000, cx: 1000, cy: 1000 };

  it("does not touch a byte for an identity move", () => {
    // An element landing where it was authored is spliced in exactly as the
    // owner drew it, down to the grid of a table a scale would have rewritten.
    const fragment = parseFragment(sp(1, [1000, 2000, 500, 500]) + table(2, [1000, 2000, 900, 300], [400, 500], [300]));
    const before = serializeXml(fragment);
    applyMove(topLevel(fragment), FROM, { dx: 0, dy: 0, sx: 1, sy: 1 });
    expect(serializeXml(fragment)).toBe(before);
  });

  it("keeps the arrangement, repositioning each shape relative to the element's frame", () => {
    // A diagram whose arrow sits between two boxes still has it between them
    // afterwards. Absolute repositioning is how an element arrives as a heap.
    const fragment = parseFragment(sp(1, [1000, 2000, 400, 400]) + sp(2, [1500, 2500, 200, 200]));
    applyMove(topLevel(fragment), FROM, { dx: 500, dy: 100, sx: 2, sy: 2 });
    expect(topLevel(fragment).map((s) => rectOf(s))).toEqual([
      { x: 1500, y: 2100, cx: 800, cy: 800 },
      { x: 2500, y: 3100, cx: 400, cy: 400 },
    ]);
  });

  it("scales a group by its frame alone, leaving the child coordinate space and the children untouched", () => {
    /**
     * `<a:chOff>`/`<a:chExt>` are the space the children are drawn in, and
     * leaving them alone while the frame changes is exactly how PowerPoint
     * scales a group. Rewriting them, or the children, applies the scale twice
     * — a group whose contents have shrunk away from its own outline.
     */
    const fragment = parseFragment(grp(1, [1000, 2000, 1000, 1000], [0, 0, 4000, 2000], sp(2, [500, 500, 200, 200])));
    const group = topLevel(fragment)[0] as Element;
    applyMove([group], FROM, { dx: 0, dy: 0, sx: 0.5, sy: 0.5 });
    const frame = groupFrame(group);
    expect(pair(frame, "off", "x", "y")).toEqual(["1000", "2000"]);
    expect(pair(frame, "ext", "cx", "cy")).toEqual(["500", "500"]);
    expect(pair(frame, "chOff", "x", "y"), "the child coordinate space was scaled too").toEqual(["0", "0"]);
    expect(pair(frame, "chExt", "cx", "cy")).toEqual(["4000", "2000"]);
    const inner = elements(group, P_NS, "sp")[0] as Element;
    expect(rectOf(inner), "a child inside the group was rewritten as well").toEqual({
      x: 500,
      y: 500,
      cx: 200,
      cy: 200,
    });
  });

  it("scales a table's own grid, because PowerPoint draws a table from that and ignores the frame", () => {
    /**
     * `docs/DESIGN.md` section 3 records the same fact from the other end: a
     * table's box is the sum of its columns and rows, "not its frame's `ext`,
     * which PowerPoint ignores when it draws the table". So a whole-slide
     * element scaled down to clear a taller title would MOVE its table and
     * leave it at its authored size, overlapping whatever it was scaled away
     * from. 136 of the 234 committed elements carry a graphic frame.
     */
    const fragment = parseFragment(table(1, [1000, 2000, 1000, 600], [400, 600], [300, 300]));
    applyMove(topLevel(fragment), FROM, { dx: 0, dy: 0, sx: 0.5, sy: 0.5 });
    const frame = topLevel(fragment)[0] as Element;
    expect(valuesOf(frame, A_NS, "gridCol", "w")).toEqual(["200", "300"]);
    expect(valuesOf(frame, A_NS, "tr", "h")).toEqual(["150", "150"]);
    expect(rectOf(frame)).toEqual({ x: 1000, y: 2000, cx: 500, cy: 300 });
  });

  it("leaves the grid alone for a move that only shifts the table", () => {
    // A landing that merely repositions an element must not rewrite a single
    // column width: the table is the size the owner drew it.
    const fragment = parseFragment(table(1, [1000, 2000, 1000, 600], [400, 600], [300]));
    applyMove(topLevel(fragment), FROM, { dx: 7000, dy: 0, sx: 1, sy: 1 });
    const frame = topLevel(fragment)[0] as Element;
    expect(valuesOf(frame, A_NS, "gridCol", "w")).toEqual(["400", "600"]);
    expect(rectOf(frame)?.x).toBe(8000);
  });

  it("never scales a shape or a column away to nothing", () => {
    // Zero is not a size PowerPoint draws, and a rounded-down extent is how a
    // small element disappears entirely on a slide it was merely shrunk for.
    const fragment = parseFragment(sp(1, [1000, 2000, 3, 3]) + table(2, [1000, 2000, 3, 3], [3], [3]));
    applyMove(topLevel(fragment), FROM, { dx: 0, dy: 0, sx: 0.01, sy: 0.01 });
    expect(rectOf(topLevel(fragment)[0] as Element)).toEqual({ x: 1000, y: 2000, cx: 1, cy: 1 });
    const frame = topLevel(fragment)[1] as Element;
    expect(valuesOf(frame, A_NS, "gridCol", "w")).toEqual(["1"]);
    expect(valuesOf(frame, A_NS, "tr", "h")).toEqual(["1"]);
  });

  it("moves what a shape states and steps over what it does not", () => {
    /**
     * A shape with no `<a:xfrm>`, one with an offset and no extent, one whose
     * numbers are not numbers. None of them can be moved, and none of them may
     * stop the rest of the element from moving: a single inherited text box in
     * a run of twelve shapes would otherwise throw the whole splice away.
     */
    const fragment = parseFragment(
      sp(1) +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:off x="1500" y="2500"/></a:xfrm></p:spPr></p:sp>` +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:ext cx="200" cy="200"/></a:xfrm></p:spPr></p:sp>` +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:off x="odd" y="2500"/><a:ext cx="odd" cy="200"/></a:xfrm></p:spPr></p:sp>` +
        `<p:sp ${NS}><p:spPr><a:xfrm><a:off x="1500" y="odd"/><a:ext cx="200" cy="odd"/></a:xfrm></p:spPr></p:sp>`,
    );
    applyMove(topLevel(fragment), FROM, { dx: 100, dy: 100, sx: 2, sy: 2 });
    const shapes = topLevel(fragment);
    expect(pair(elements(shapes[1] as Element, A_NS, "xfrm")[0], "off", "x", "y")).toEqual(["2100", "3100"]);
    expect(pair(elements(shapes[2] as Element, A_NS, "xfrm")[0], "ext", "cx", "cy")).toEqual(["400", "400"]);
    // Guarded per attribute, not per shape: the axis that states a number still
    // moves, and `Math.round(NaN * 2)` never reaches the file as a coordinate.
    const odd = elements(shapes[3] as Element, A_NS, "xfrm")[0];
    expect(pair(odd, "off", "x", "y"), "a number that is not one was written over anyway").toEqual(["odd", "3100"]);
    expect(pair(odd, "ext", "cx", "cy")).toEqual(["odd", "400"]);
    const oddY = elements(shapes[4] as Element, A_NS, "xfrm")[0];
    expect(pair(oddY, "off", "x", "y")).toEqual(["2100", "odd"]);
    expect(pair(oddY, "ext", "cx", "cy")).toEqual(["400", "odd"]);
  });

  it("leaves a grid dimension that is not a number alone, rather than writing NaN into the file", () => {
    // The same guard on the table's own grid. A column stated as `NaN` is a
    // part PowerPoint refuses outright, and the scale is the one pass that
    // rewrites those numbers at all.
    const fragment = parseFragment(
      `<p:graphicFrame ${NS}><p:xfrm><a:off x="1000" y="2000"/><a:ext cx="10" cy="10"/></p:xfrm>` +
        `<a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="auto"/></a:tblGrid>` +
        `<a:tr h="auto"><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
    );
    applyMove(topLevel(fragment), FROM, { dx: 0, dy: 0, sx: 0.5, sy: 0.5 });
    const frame = topLevel(fragment)[0] as Element;
    expect(valuesOf(frame, A_NS, "gridCol", "w")).toEqual(["auto"]);
    expect(valuesOf(frame, A_NS, "tr", "h")).toEqual(["auto"]);
  });
});

describe("wrapping the shapes in one group", () => {
  it("refuses a run containing a graphic frame, because PowerPoint refuses one too", () => {
    /**
     * PowerPoint will not group a table through its own UI, so a table inside a
     * `<p:grpSp>` is markup it has no way to have produced. Refusing is the
     * right shape rather than failing: the element still lands, it lands loose,
     * and the pane says so.
     */
    const fragment = parseFragment(sp(1, [0, 0, 10, 10]) + table(2, [0, 0, 10, 10], [10], [10]));
    expect(groupable(topLevel(fragment))).toBe(false);
    expect(groupable(topLevel(parseFragment(sp(1) + sp(2))))).toBe(true);
  });

  it("refuses a single shape, which is already one thing to move", () => {
    expect(groupable(topLevel(parseFragment(sp(1))))).toBe(false);
    expect(groupable([])).toBe(false);
  });

  it("sets the child coordinate space equal to its own, so nothing inside has to move", () => {
    /**
     * Any other pair of values would mean rewriting every child's coordinates,
     * and getting that arithmetic subtly wrong is a group whose contents are
     * offset by a few millimetres in a way nobody notices until the deck is
     * printed.
     */
    const fragment = parseFragment(sp(1, [200, 0, 100, 100]) + sp(2, [400, 0, 100, 100]));
    const group = groupShapes(fragment, topLevel(fragment), "Boxes, 2 vertical", 77);
    const frame = groupFrame(group);
    expect(pair(frame, "off", "x", "y")).toEqual(["200", "0"]);
    expect(pair(frame, "ext", "cx", "cy")).toEqual(["300", "100"]);
    expect(pair(frame, "chOff", "x", "y")).toEqual(pair(frame, "off", "x", "y"));
    expect(pair(frame, "chExt", "cx", "cy")).toEqual(pair(frame, "ext", "cx", "cy"));
    // The group is a shape like any other, and the splice asks it where it is.
    expect(rectOf(group)).toEqual({ x: 200, y: 0, cx: 300, cy: 100 });
    expect(elements(group, P_NS, "cNvPr")[0]?.getAttribute("name")).toBe("Boxes, 2 vertical");
    expect(elements(group, P_NS, "cNvPr")[0]?.getAttribute("id")).toBe("77");
  });

  it("takes the first shape's place in the run, so an element drawn behind another stays behind it", () => {
    // Slide order IS z-order. A group appended at the end of the fragment comes
    // out in front of whatever was drawn over it in the library.
    const fragment = parseFragment(sp(1, [0, 0, 100, 100]) + sp(2, [200, 0, 100, 100]) + sp(3, [400, 0, 100, 100]));
    const [first, second, third] = topLevel(fragment) as [Element, Element, Element];
    const group = groupShapes(fragment, [second, third], "Boxes", 9);
    expect(topLevel(fragment)).toEqual([first, group]);
    expect(Array.from(group.childNodes).filter((n) => n.nodeType === 1).length).toBe(4);
    expect(elements(group, P_NS, "sp")).toEqual([second, third]);
  });

  it("gives a run that states no box a frame of nothing rather than of NaN", () => {
    // An element whose shapes all inherit their geometry has no union to take.
    // A group sized `NaN` is a part PowerPoint refuses outright.
    const fragment = parseFragment(sp(1) + sp(2));
    const frame = groupFrame(groupShapes(fragment, topLevel(fragment), "Inherited", 3));
    expect(pair(frame, "off", "x", "y")).toEqual(["0", "0"]);
    expect(pair(frame, "ext", "cx", "cy")).toEqual(["0", "0"]);
  });
});

describe("the placeholder claim a library shape carries", () => {
  it("strips a p:ph and leaves every visual property the shape had", () => {
    /**
     * Two shapes claiming one placeholder index is resolved by PowerPoint
     * inheriting geometry from the LAYOUT, so the element jumps to wherever the
     * layout's placeholder is and the landing the splice computed is thrown
     * away. What the shape loses is the claim, which was never true of the
     * destination.
     */
    const fragment = parseFragment(
      `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="1" name="Indhold"/><p:cNvSpPr/>` +
        `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm([1, 2, 3, 4])}</p:spPr></p:sp>`,
    );
    const shapes = topLevel(fragment);
    expect(unplaceholder(shapes)).toBe(1);
    expect(elements(shapes[0] as Element, P_NS, "ph")).toEqual([]);
    expect(rectOf(shapes[0] as Element)).toEqual({ x: 1, y: 2, cx: 3, cy: 4 });
  });

  it("reaches a claim made by a shape inside a group, and counts every one it stripped", () => {
    const inner =
      `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="2" name="Indhold"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`;
    const fragment = parseFragment(grp(1, [0, 0, 10, 10], [0, 0, 10, 10], inner, inner));
    expect(unplaceholder(topLevel(fragment))).toBe(2);
    expect(elements(topLevel(fragment)[0] as Element, P_NS, "ph")).toEqual([]);
  });

  it("answers nothing stripped for shapes that claim nothing, which is most of the library", () => {
    expect(unplaceholder(topLevel(parseFragment(sp(1) + pic(2, "rId1"))))).toBe(0);
  });
});

describe("the destination slide's own shapes", () => {
  it("lists the shapes on the slide without the shape tree's own properties", () => {
    // `<p:nvGrpSpPr>` and `<p:grpSpPr>` are the tree's, not shapes on it.
    // Counting them is how a slide with nothing on it looks occupied.
    const spTree = tree(sp(1), pic(2, "rId1"), table(3, [0, 0, 10, 10], [10], [10]));
    expect(slideShapes(spTree).map((s) => s.localName)).toEqual(["sp", "pic", "graphicFrame"]);
    expect(slideShapes(tree())).toEqual([]);
    // A producer that pretty-prints its parts leaves a text node between every
    // shape, and the destination deck is whatever the user brought.
    expect(slideShapes(tree(`\n  ${sp(1)}\n  ${pic(2, "rId1")}\n`)).map((s) => s.localName)).toEqual(["sp", "pic"]);
  });

  it("finds an empty body placeholder to remove, and never the title", () => {
    /**
     * `docs/DESIGN.md` section 6: a whole-slide element removes the "Click to
     * add text" ghosts behind it, and the TITLE stays. A rule that took the
     * title as well leaves the user's own heading gone from a slide they only
     * added a picture to.
     */
    const spTree = tree(
      placeholder(1, { type: "title" }),
      placeholder(2, { type: "ctrTitle" }),
      placeholder(3, { type: "body" }),
    );
    expect(emptyBodyPlaceholders(spTree).map((s) => shapeIds(s)[0])).toEqual(["3"]);
  });

  it("never one the user has typed into, because content is not a ghost", () => {
    // Whitespace is not content: a placeholder holding a stray space is still
    // the ghost the layout put there.
    const spTree = tree(
      placeholder(1, { type: "body", text: "Vores tal for 2026" }),
      placeholder(2, { type: "body", text: "   " }),
    );
    expect(emptyBodyPlaceholders(spTree).map((s) => shapeIds(s)[0])).toEqual(["2"]);
  });

  it("treats a placeholder that states no type as a body, which is what PowerPoint does", () => {
    // `<p:ph idx="1"/>` with no type IS a body placeholder, and it is the
    // commonest spelling in a deck built from the standard layouts.
    expect(emptyBodyPlaceholders(tree(placeholder(1))).map((s) => shapeIds(s)[0])).toEqual(["1"]);
  });

  it("passes over a shape that claims nothing, states no text body, or is not a shape at all", () => {
    // A picture is never a ghost, a placeholder with no `<p:txBody>` is a
    // picture or a table placeholder holding content, and an ordinary text box
    // is the user's. The two bare shapes are a slide stripped by another
    // add-in: markup with no `<p:nvSpPr>` at all, and one whose `<p:nvSpPr>`
    // states no `<p:nvPr>` — neither of which may be read as a claim.
    const spTree = tree(
      sp(1),
      pic(2, "rId1"),
      placeholder(3, { type: "body", noBody: true }),
      grp(4, [0, 0, 10, 10], [0, 0, 10, 10]),
      `<p:sp><p:spPr/></p:sp>`,
      `<p:sp><p:nvSpPr><p:cNvPr id="6" name="Stripped"/><p:cNvSpPr/></p:nvSpPr><p:spPr/></p:sp>`,
    );
    expect(emptyBodyPlaceholders(spTree)).toEqual([]);
  });
});

/**
 * The claim `shapes.ts` makes about the library rather than about a rule.
 *
 * `FRAGMENT_ROOT`'s comment says the wrapper "declares nothing, and it does not
 * have to: every prefix each shape uses is declared inside that shape", and that
 * `test/splice-shapes.test.ts` checks it holds for all 234 elements in the
 * committed library rather than for the one the comment was written against.
 * This is that check. It is worth the two harvests: the property belongs to the
 * SERIALISER, so a change to how the harvest writes a shape out — or one
 * element redrawn in a way that leans on a declaration from the library slide's
 * root — is a splice that throws in the pane on that element and only that one.
 */
const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

const library: { size: SlideSize; catalogue: Catalogue }[] = [];

beforeAll(async () => {
  for (const [size, file] of [
    ["16:9", "template/library-16x9.pptx"],
    ["4:3", "template/library-4x3.pptx"],
  ] as [SlideSize, string][]) {
    const pkg = await Pkg.open(new Uint8Array(readFileSync(file)));
    const { catalogue } = await harvest(pkg, { size, names: NAMES });
    library.push({ size, catalogue });
  }
}, 180_000);

/** Every element in both committed decks, each labelled by the deck it came from. */
function committed(): { size: SlideSize; el: CatalogueElement }[] {
  return library.flatMap((deck) => deck.catalogue.elements.map((el) => ({ size: deck.size, el })));
}

describe("the committed library", () => {
  it("parses every one of its 234 elements standalone, because each declares the prefixes it uses", () => {
    const problems: string[] = [];
    for (const { size, el } of committed()) {
      try {
        const tops = topLevel(parseFragment(el.markup.xml));
        // A fragment that parsed to nothing would pass a check for "did not
        // throw" while carrying no element at all.
        if (tops.length !== el.shapes) {
          problems.push(`${size} ${el.id}: parsed ${tops.length} shapes, the catalogue counted ${el.shapes}`);
        }
      } catch (error) {
        problems.push(`${size} ${el.id}: ${(error as Error).message}`);
      }
    }
    // Listed all at once rather than one at a time, the way the harvest reports
    // a missing name: one run says which elements are wrong, not the first.
    expect(problems).toEqual([]);
    expect(committed()).toHaveLength(234);
  }, 120_000);

  it("is made of the markup that makes parsing it worth checking, so the sweep is not vacuous", () => {
    /**
     * Measured against the committed decks. A library of plain rectangles would
     * parse under any wrapper at all; these numbers are the prefixes beyond
     * `p` and `a` that the sweep above actually has to resolve — `r` on 376
     * relationship references, `mc` on 57 alternate-content fallbacks, and the
     * `c`, `a14`, `a16` and `p14` a real drawing carries.
     */
    let groups = 0;
    let frames = 0;
    let fallbacks = 0;
    let references = 0;
    for (const { el } of committed()) {
      const fragment = parseFragment(el.markup.xml);
      for (const shape of topLevel(fragment)) {
        groups += elements(shape, P_NS, "grpSp").length + (shape.localName === "grpSp" ? 1 : 0);
        frames += elements(shape, P_NS, "graphicFrame").length + (shape.localName === "graphicFrame" ? 1 : 0);
        fallbacks += elements(shape, MC_NS, "Fallback").length;
      }
      references += relIdsIn(fragment).length;
    }
    expect(groups).toBe(205);
    expect(frames).toBe(163);
    expect(fallbacks).toBe(57);
    expect(references).toBe(376);
  }, 120_000);

  it("gives every shape of every element an id of its own when renumbered", () => {
    /**
     * The synthetic case above proves a renumber reaches a group's children and
     * an `<mc:Fallback>`; this proves there is no THIRD place in the library
     * where a `<p:cNvPr>` hides. A duplicate id is a slide PowerPoint opens
     * with one of the shapes missing, which is not a failure anybody reports as
     * a splice bug.
     */
    const problems: string[] = [];
    for (const { size, el } of committed()) {
      const fragment = parseFragment(el.markup.xml);
      const next = renumber(fragment, 5000);
      const ids = topLevel(fragment).flatMap((shape) => shapeIds(shape));
      if (new Set(ids).size !== ids.length) problems.push(`${size} ${el.id}: two shapes share an id`);
      if (next !== 5000 + ids.length) problems.push(`${size} ${el.id}: answered ${next} after ${ids.length} shapes`);
    }
    expect(problems).toEqual([]);
  }, 120_000);

  it("names exactly the relationships the harvest resolved for each element", () => {
    /**
     * Two readers of the same markup, written independently — the harvest's own
     * relationship scan and `relIdsIn` here — held against each other over 234
     * elements. They are the pair the splice depends on agreeing: the harvest
     * decides which parts to carry, `relIdsIn` and `repoint` decide which
     * references get rewritten, and an id one sees and the other does not is
     * either a dangling reference in the destination or a part copied for
     * nothing. 73 of the elements name at least one.
     */
    const problems: string[] = [];
    let naming = 0;
    for (const { size, el } of committed()) {
      const named = [...relIdsIn(parseFragment(el.markup.xml))].sort();
      const resolved = el.markup.rels.map((rel) => rel.id).sort();
      if (named.length > 0) naming += 1;
      if (named.join(",") !== resolved.join(",")) {
        problems.push(
          `${size} ${el.id}: markup names ${named.join("/")}, the catalogue resolved ${resolved.join("/")}`,
        );
      }
    }
    expect(problems).toEqual([]);
    expect(naming).toBe(73);
  });
});
