import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { framesOf, layoutOf, masterOf, placeholderRect, slideSize, type Rect } from "../src/core/pptx/layout.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * Where the destination slide's placeholders are, on decks the fixture built.
 *
 * The module exists because a real slide almost never carries its own geometry.
 * PowerPoint writes `<p:spPr/>` and nothing else on a title the user never
 * dragged, and leaves the rectangle to the layout, which leaves it to the
 * master — so a reader that looks only at the slide finds nothing on almost
 * every real deck and concludes the slide has no title. The element then lands
 * straight through a customer's two-line heading.
 *
 * So most of what is tested below is the WALK: what stops it, what it must not
 * pick up on the way, and which placeholder is the SAME placeholder rather than
 * merely the same kind. Every case is one rectangle `splice/landing.ts` would
 * otherwise be handed wrongly, and it has no way to tell a wrong rectangle from
 * a right one.
 *
 * The fixture deck wires a slide to a layout to a master and puts placeholders
 * on none of them, which is the useful starting point: each test writes over
 * exactly the parts it is about and inherits an empty shape tree everywhere
 * else, so nothing can pass by accident. The last block runs against
 * `template/library-16x9.pptx`, the deck the product ships, where the
 * inheritance is the owner's rather than mine.
 */

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * The relationship types, spelled out rather than imported from `parts.ts`.
 *
 * A fixture that names a type by the same constant the module reads agrees
 * with a typo in it, and the test would go green on a package no PowerPoint
 * could follow. `test/fixtures/deck.ts` keeps its own copy for the same reason.
 */
const LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const NOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const SLIDE = "ppt/slides/slide1.xml";
const SLIDE_RELS = "ppt/slides/_rels/slide1.xml.rels";
const LAYOUT_RELS = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
const LAYOUT = "ppt/slideLayouts/slideLayout1.xml";
const LAYOUT2 = "ppt/slideLayouts/slideLayout2.xml";
const MASTER = "ppt/slideMasters/slideMaster1.xml";
const PRESENTATION = "ppt/presentation.xml";

/**
 * The rectangles, one per role.
 *
 * Distinct on every side, so a failure message says which part the answer came
 * from rather than only that it was wrong. The layout and master numbers are
 * the sizes a real 16:9 template uses; the point is that they differ.
 */
const SLIDE_OWN: Rect = { x: 1000000, y: 1000000, cx: 5000000, cy: 1000000 };
const LAYOUT_TITLE: Rect = { x: 838200, y: 365125, cx: 10515600, cy: 1325563 };
const LAYOUT_BODY: Rect = { x: 838200, y: 1825625, cx: 10515600, cy: 4351338 };
const MASTER_TITLE: Rect = { x: 628650, y: 274638, cx: 10934700, cy: 1143000 };
const MASTER_BODY: Rect = { x: 628650, y: 1600200, cx: 10934700, cy: 4525963 };
const LEFT_COLUMN: Rect = { x: 838200, y: 1825625, cx: 5100000, cy: 4351338 };
const RIGHT_COLUMN: Rect = { x: 6253800, y: 1825625, cx: 5100000, cy: 4351338 };
const CHROME_STRIP: Rect = { x: 838200, y: 6356350, cx: 2743200, cy: 365125 };
const FULL_BLEED: Rect = { x: 0, y: 0, cx: 12192000, cy: 6858000 };

function offExt(rect: Rect): string {
  return `<a:off x="${rect.x}" y="${rect.y}"/><a:ext cx="${rect.cx}" cy="${rect.cy}"/>`;
}

function xfrm(rect: Rect): string {
  return `<a:xfrm>${offExt(rect)}</a:xfrm>`;
}

interface PhSpec {
  type?: string;
  idx?: string;
  rect?: Rect;
  /** Raw `<p:spPr>` content, for the half-written geometry the walk has to survive. */
  spPr?: string;
  id?: number;
}

/** A placeholder shape, carrying geometry only when it is given some. */
function ph(spec: PhSpec = {}): string {
  const type = spec.type === undefined ? "" : ` type="${spec.type}"`;
  const idx = spec.idx === undefined ? "" : ` idx="${spec.idx}"`;
  const id = spec.id ?? 2;
  const geometry = spec.spPr ?? (spec.rect ? xfrm(spec.rect) : "");
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type}${idx}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${geometry}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
  );
}

/**
 * A table placeholder, whose geometry is a `<p:xfrm>` of its own.
 *
 * A `<p:graphicFrame>` has no `<p:spPr>` at all — tables, charts and SmartArt
 * on a slide are all this shape — so a reader that only looks inside `spPr`
 * decides the frame is geometryless and inherits the layout's rectangle
 * instead of the one the user dragged the table to.
 */
function frame(spec: { idx?: string; rect: Rect }): string {
  const idx = spec.idx === undefined ? "" : ` idx="${spec.idx}"`;
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Tabel 6"/><p:cNvGraphicFramePr/>` +
    `<p:nvPr><p:ph type="tbl"${idx}/></p:nvPr></p:nvGraphicFramePr>` +
    `<p:xfrm>${offExt(spec.rect)}</p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic></p:graphicFrame>`
  );
}

/** A shape that is not a placeholder: the decoration a real layout is full of. */
function decoration(rect: Rect): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="8" name="Rektangel 7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(rect)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
  );
}

/** The group shape every real shape tree opens with, and no placeholder. */
const TREE_HEAD = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`;

function slidePart(...shapes: string[]): string {
  return (
    `${DECL}<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>${TREE_HEAD}${shapes.join("")}</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

function layoutPart(...shapes: string[]): string {
  return (
    `${DECL}<p:sldLayout ${P} ${A} ${R} type="obj"><p:cSld><p:spTree>${TREE_HEAD}${shapes.join("")}</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

function masterPart(...shapes: string[]): string {
  return (
    `${DECL}<p:sldMaster ${P} ${A} ${R}><p:cSld><p:spTree>${TREE_HEAD}${shapes.join("")}</p:spTree></p:cSld>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`
  );
}

/** The same part, pretty-printed: a producer that indents its XML is a real one. */
function indented(part: string): string {
  return part.replaceAll("><", ">\n  <");
}

function rel(id: string, type: string, target?: string, mode?: string): string {
  const to = target === undefined ? "" : ` Target="${target}"`;
  const external = mode === undefined ? "" : ` TargetMode="${mode}"`;
  return `<Relationship Id="${id}" Type="${type}"${to}${external}/>`;
}

function relsPart(...rels: string[]): string {
  return `${DECL}<Relationships ${RELS}>${rels.join("")}</Relationships>`;
}

function presentationPart(sldSz: string): string {
  return (
    `${DECL}<p:presentation ${P} ${A} ${R}>` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `${sldSz}<p:notesSz cx="6858000" cy="9144000"/></p:presentation>`
  );
}

/**
 * The fixture deck, with the parts a test cares about written over.
 *
 * The slide the fixture builds is already the common case: a title placeholder
 * carrying `<p:spPr/>` and a body shape that is not a placeholder at all.
 */
async function deck(parts: Record<string, string> = {}): Promise<Pkg> {
  const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Body text"]], title: "A destination slide" }]));
  for (const [path, xml] of Object.entries(parts)) pkg.setText(path, xml);
  return pkg;
}

describe("the layout and the master a slide inherits from", () => {
  it("follows a slide to its layout and that layout to its master", async () => {
    // The layout carries a logo, so its master is not its first relationship.
    // Both hops are by TYPE, and a hop that takes the first relationship it
    // finds answers a picture part on any template with branding on the layout.
    const pkg = await deck({
      [LAYOUT_RELS]: relsPart(
        rel("rId1", IMAGE_REL, "../media/image1.png"),
        rel("rId2", MASTER_REL, "../slideMasters/slideMaster1.xml"),
      ),
    });
    const layout = (await layoutOf(pkg, SLIDE)) ?? "";
    expect(layout).toBe(LAYOUT);
    expect(await masterOf(pkg, layout)).toBe(MASTER);
  });

  it("answers nothing for a slide the package does not hold", async () => {
    // The pane asks by path, and a path that is not in the package has no rels
    // part either. Both walks have to answer rather than throw: `splice` reads
    // the frames before it has done anything, and a throw there is a pane that
    // says the deck is broken when the caller merely asked about slide 9.
    const pkg = await deck();
    expect(await layoutOf(pkg, "ppt/slides/slide9.xml")).toBeUndefined();
    expect(await placeholderRect(pkg, "ppt/slides/slide9.xml", "title")).toBeUndefined();
  });

  it("answers nothing when the slide's relationships name no layout", async () => {
    const pkg = await deck({ [SLIDE_RELS]: relsPart(rel("rId2", NOTES_REL, "../notesSlides/notesSlide1.xml")) });
    expect(await layoutOf(pkg, SLIDE)).toBeUndefined();
    // And the rectangle walk stops there rather than reaching for whichever
    // layout the package happens to hold: a slide with no layout of its own
    // has no inherited geometry, and guessing one is worse than none.
    expect(await placeholderRect(pkg, SLIDE, "title")).toBeUndefined();
  });

  it("steps over a relationship whose target has left the package", async () => {
    // A deck a third-party tool has edited keeps relationships to parts it
    // deleted. Taking the first one of the right TYPE answers a part that is
    // not there, `spTreeOf` finds nothing, and every rectangle is lost — on a
    // deck PowerPoint itself opens without a murmur.
    const pkg = await deck({
      [SLIDE_RELS]: relsPart(
        rel("rId1", LAYOUT_REL, "../slideLayouts/slideLayout9.xml"),
        rel("rId2", LAYOUT_REL, "../slideLayouts/slideLayout1.xml"),
      ),
      [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE })),
    });
    expect(await layoutOf(pkg, SLIDE)).toBe(LAYOUT);
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(LAYOUT_TITLE);
  });

  it("does not follow an external relationship, even when its target reads like a part", async () => {
    // `TargetMode="External"` means the target is not a part in this package,
    // whatever it is spelled like — and a target spelled like a part that IS
    // in the package is the one case where dropping the check is invisible.
    // Here it would answer layout 1, and the slide's real layout is layout 2.
    const pkg = await deck({
      [SLIDE_RELS]: relsPart(
        rel("rId1", LAYOUT_REL, "../slideLayouts/slideLayout1.xml", "External"),
        rel("rId2", LAYOUT_REL, "../slideLayouts/slideLayout2.xml"),
      ),
      [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE })),
      [LAYOUT2]: layoutPart(ph({ type: "title", rect: MASTER_TITLE })),
    });
    expect(await layoutOf(pkg, SLIDE)).toBe(LAYOUT2);
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(MASTER_TITLE);
  });

  it("ignores a relationship with no target at all", async () => {
    // `Target` is required, so this is malformed — and resolving the missing
    // attribute as the empty string names the slide's own directory, which
    // `has` answers for on a package where a directory entry exists.
    const pkg = await deck({
      [SLIDE_RELS]: relsPart(rel("rId1", LAYOUT_REL), rel("rId2", LAYOUT_REL, "../slideLayouts/slideLayout1.xml")),
    });
    expect(await layoutOf(pkg, SLIDE)).toBe(LAYOUT);
  });
});

describe("where the destination's title is", () => {
  it("takes the slide's own rectangle over the layout's", async () => {
    // The one deck it would be wrong to overrule: a user who dragged the title
    // is telling the add-in where the title is.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "title", rect: SLIDE_OWN })),
      [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(SLIDE_OWN);
  });

  it("inherits the layout's rectangle when the slide's title carries no geometry", async () => {
    // The common case, and the reason the module exists. The fixture writes the
    // title the way PowerPoint does — `<p:spPr/>`, empty — so a reader that
    // stops at the slide answers "this deck has no title" for almost every
    // deck a customer owns, and the element lands through the heading.
    const pkg = await deck({ [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE })) });
    expect(await pkg.text(SLIDE), "the fixture's title must carry no geometry, or this proves nothing").toContain(
      '<p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>',
    );
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(LAYOUT_TITLE);
  });

  it("falls through to the master when the layout does not place it either", async () => {
    // The layout's title is PRESENT and geometryless, which is the shape that
    // ends a walk one part early: a placeholder found is not a rectangle found.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "title" })),
      [LAYOUT]: layoutPart(ph({ type: "title" })),
      [MASTER]: masterPart(ph({ type: "title", rect: MASTER_TITLE })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(MASTER_TITLE);
  });

  it("answers nothing when no part in the chain places the placeholder", async () => {
    // The fixture's layout and master are empty shape trees. `landing.ts` has
    // its own fallback for this and needs to be told, not guessed at.
    const pkg = await deck({ [SLIDE]: slidePart(ph({ type: "title" })) });
    expect(await placeholderRect(pkg, SLIDE, "title")).toBeUndefined();
    expect(await framesOf(pkg, SLIDE)).toEqual({});
  });
});

describe("which placeholder is a body", () => {
  it("treats a placeholder with no type at all as a body, which is what the schema says", async () => {
    // `type` defaults to `body`, and PowerPoint omits it constantly. A reader
    // keying on the attribute misses the body of an ordinary content layout
    // entirely, so the element lands against the slide edge instead of inside
    // the body area.
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(ph({ idx: "1", rect: LAYOUT_BODY })),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(LAYOUT_BODY);
    expect(await placeholderRect(pkg, SLIDE, "title"), "a typeless placeholder is not a title").toBeUndefined();
  });

  it("never lands on the footer, the slide number or the date", async () => {
    // All three are chrome along the bottom of the slide, and all three are
    // placeholders with neither a title's type nor a body's. Listed FIRST here
    // on purpose: a walk that takes the first placeholder it recognises has to
    // reject these by name, and the answer must be the body's rectangle
    // whatever order the layout happens to list them in.
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(
        ph({ type: "dt", idx: "10", rect: CHROME_STRIP, id: 3 }),
        ph({ type: "ftr", idx: "11", rect: CHROME_STRIP, id: 4 }),
        ph({ type: "sldNum", idx: "12", rect: CHROME_STRIP, id: 5 }),
        ph({ type: "body", idx: "1", rect: LAYOUT_BODY, id: 6 }),
      ),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(LAYOUT_BODY);
  });

  it("answers nothing for a layout that is chrome and nothing else", async () => {
    // The title-slide shape: date, footer and slide number, no content
    // placeholder anywhere. An element landing in the footer strip is the
    // failure this half of the rule prevents.
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(
        ph({ type: "dt", idx: "10", rect: CHROME_STRIP, id: 3 }),
        ph({ type: "ftr", idx: "11", rect: CHROME_STRIP, id: 4 }),
        ph({ type: "sldNum", idx: "12", rect: CHROME_STRIP, id: 5 }),
      ),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toBeUndefined();
    expect(await placeholderRect(pkg, SLIDE, "title")).toBeUndefined();
  });

  it("counts ctrTitle as a title and subTitle as a body, the way PowerPoint does", async () => {
    // The layout of a real title slide, in the order the shipped library writes
    // it: the subtitle first, the centred title second. Two shapes that a
    // reader matching on the exact strings "title" and "body" reads as neither.
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(
        ph({ type: "subTitle", idx: "1", rect: LAYOUT_BODY, id: 3 }),
        ph({ type: "ctrTitle", rect: LAYOUT_TITLE, id: 4 }),
      ),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(LAYOUT_TITLE);
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(LAYOUT_BODY);
  });

  it("does not read a rectangle off a shape that is no placeholder at all", async () => {
    // A layout's decoration comes first and is usually the biggest thing on it.
    // Read as a placeholder it hands back a full-bleed background, and the
    // "body area" the element lands in is the whole slide.
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(decoration(FULL_BLEED), ph({ type: "title", rect: LAYOUT_TITLE, id: 4 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(LAYOUT_TITLE);
    expect(await placeholderRect(pkg, SLIDE, "body"), "decoration is not a typeless body").toBeUndefined();
  });
});

describe("the same placeholder, not merely the same kind", () => {
  it("matches the column the slide named, so a two-column layout gives back the right one", async () => {
    // Two content placeholders side by side are distinguished only by `idx`.
    // Taking the first would land every element in the left column whichever
    // one the user was in.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "body", idx: "2", id: 3 })),
      [LAYOUT]: layoutPart(
        ph({ type: "body", idx: "1", rect: LEFT_COLUMN, id: 4 }),
        ph({ type: "body", idx: "2", rect: RIGHT_COLUMN, id: 5 }),
      ),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(RIGHT_COLUMN);
  });

  it("answers nothing rather than the wrong column when the index is nowhere in the layout", async () => {
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "body", idx: "3", id: 3 })),
      [LAYOUT]: layoutPart(ph({ type: "body", idx: "1", rect: LEFT_COLUMN, id: 4 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toBeUndefined();
  });

  it("does not settle for a layout placeholder that has no index at all", async () => {
    // A `<p:ph>` with no `idx` is index zero, not "any index". A layout's own
    // title and its first content placeholder are both written that way, so
    // treating a missing index as a match hands back the layout's first
    // content placeholder for every slide, and the two-column rule above stops
    // meaning anything on the layouts where it matters most.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "body", idx: "1", id: 3 })),
      [LAYOUT]: layoutPart(
        ph({ type: "body", rect: LEFT_COLUMN, id: 4 }),
        ph({ type: "body", idx: "1", rect: RIGHT_COLUMN, id: 5 }),
      ),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(RIGHT_COLUMN);
  });

  it("carries the index all the way to the master", async () => {
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "body", idx: "3", id: 3 })),
      [LAYOUT]: layoutPart(ph({ type: "body", idx: "1", rect: LEFT_COLUMN, id: 4 })),
      [MASTER]: masterPart(ph({ type: "body", idx: "3", rect: MASTER_BODY, id: 5 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(MASTER_BODY);
  });

  it("matches a title on its type alone, because a title has no index", async () => {
    // The mirror of the rule above, and the reason it is written as "the index
    // the slide named" rather than "the indices agree": a slide's title has no
    // `idx`, a layout's may have one, and requiring them to match loses the
    // title on any layout whose author gave it an index.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "title", id: 3 })),
      [LAYOUT]: layoutPart(ph({ type: "title", idx: "0", rect: LAYOUT_TITLE, id: 4 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(LAYOUT_TITLE);
  });
});

describe("a rectangle worth answering with", () => {
  it("reads a table placeholder's geometry, which is not inside p:spPr", async () => {
    const pkg = await deck({
      [SLIDE]: slidePart(frame({ idx: "1", rect: SLIDE_OWN })),
      [LAYOUT]: layoutPart(ph({ type: "body", idx: "1", rect: LAYOUT_BODY, id: 4 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "body")).toEqual(SLIDE_OWN);
  });

  it("walks past half a rectangle instead of completing it", async () => {
    // An `<a:off>` with no `<a:ext>` says where the placeholder starts and
    // nothing about its size. Defaulting the missing half to zero answers a
    // title of no height, and everything landing "below the title" lands at the
    // very top of the slide, over it.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "title" })),
      [LAYOUT]: layoutPart(ph({ type: "title", spPr: `<a:xfrm><a:off x="838200" y="365125"/></a:xfrm>`, id: 4 })),
      [MASTER]: masterPart(ph({ type: "title", rect: MASTER_TITLE, id: 5 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(MASTER_TITLE);
  });

  it("walks past numbers that are not numbers", async () => {
    // `Number("auto")` is NaN, every landing rule adds to it, and a shape whose
    // `<a:off>` is NaN is markup PowerPoint repairs by dropping the shape — the
    // element the user asked for is simply not there, with nothing said.
    const pkg = await deck({
      [SLIDE]: slidePart(ph({ type: "title" })),
      [LAYOUT]: layoutPart(
        ph({ type: "title", spPr: `<a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="auto"/></a:xfrm>` }),
      ),
      [MASTER]: masterPart(ph({ type: "title", rect: MASTER_TITLE, id: 5 })),
    });
    expect(await placeholderRect(pkg, SLIDE, "title")).toEqual(MASTER_TITLE);
  });

  it("reads parts a producer pretty-printed", async () => {
    // Indentation between elements is text nodes, in the shape tree and inside
    // every shape, and a walk that treats each child node as an element asks a
    // run of spaces for its `<p:nvPr>`. Nothing PowerPoint writes is indented
    // and nothing the fixture writes is either, so without this case both
    // guards are untested — and the deck that breaks them comes from whichever
    // tool the customer's colleague used, not from here.
    const pkg = await deck({
      [SLIDE]: indented(slidePart(ph({ type: "title", id: 3 }))),
      [LAYOUT]: indented(
        layoutPart(
          decoration(FULL_BLEED),
          ph({ type: "title", rect: LAYOUT_TITLE, id: 4 }),
          ph({ type: "body", idx: "1", rect: LAYOUT_BODY, id: 5 }),
        ),
      ),
    });
    expect(await framesOf(pkg, SLIDE)).toEqual({ title: LAYOUT_TITLE, body: LAYOUT_BODY });
  });

  it("walks past a part that has no shape tree instead of stopping at it", async () => {
    // Three shapes of empty part, all of which a deck generator has produced:
    // a layout with no `<p:cSld>`, a master whose `<p:cSld>` holds no shape
    // tree, and a slide with neither. The walk has to continue through the
    // first, answer nothing for the second, and still inherit for the third.
    const noCSld = `${DECL}<p:sldLayout ${P} ${A} ${R}/>`;
    const past = await deck({
      [SLIDE]: slidePart(ph({ type: "title" })),
      [LAYOUT]: noCSld,
      [MASTER]: masterPart(ph({ type: "title", rect: MASTER_TITLE, id: 5 })),
    });
    expect(await placeholderRect(past, SLIDE, "title")).toEqual(MASTER_TITLE);

    const nowhere = await deck({
      [SLIDE]: slidePart(ph({ type: "title" })),
      [LAYOUT]: noCSld,
      [MASTER]: `${DECL}<p:sldMaster ${P} ${A} ${R}><p:cSld/></p:sldMaster>`,
    });
    expect(await placeholderRect(nowhere, SLIDE, "title")).toBeUndefined();

    const emptySlide = await deck({
      [SLIDE]: `${DECL}<p:sld ${P} ${A} ${R}/>`,
      [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE, id: 4 })),
    });
    expect(await placeholderRect(emptySlide, SLIDE, "title")).toEqual(LAYOUT_TITLE);
  });
});

describe("both frames at once", () => {
  it("answers the title and the body together", async () => {
    const pkg = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(
        ph({ type: "title", rect: LAYOUT_TITLE, id: 3 }),
        ph({ type: "body", idx: "1", rect: LAYOUT_BODY, id: 4 }),
      ),
    });
    expect(await framesOf(pkg, SLIDE)).toEqual({ title: LAYOUT_TITLE, body: LAYOUT_BODY });
  });

  it("leaves out a frame the deck does not place, rather than holding a key with nothing in it", async () => {
    // `underTitle` reads `frames.body ?? …`, so an absent key and a present
    // undefined land the same today. What differs is what the answer SAYS, and
    // a `body` key holding undefined says the deck has a body and it is
    // nowhere. `Object.keys` is what makes the difference visible at all —
    // `toEqual` treats the two spellings as one.
    const titleOnly = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(ph({ type: "title", rect: LAYOUT_TITLE, id: 3 })),
    });
    expect(Object.keys(await framesOf(titleOnly, SLIDE))).toEqual(["title"]);

    const bodyOnly = await deck({
      [SLIDE]: slidePart(),
      [LAYOUT]: layoutPart(ph({ type: "body", idx: "1", rect: LAYOUT_BODY, id: 3 })),
    });
    expect(Object.keys(await framesOf(bodyOnly, SLIDE))).toEqual(["body"]);
  });
});

describe("the deck's slide size", () => {
  it("reads the size the deck declares", async () => {
    expect(await slideSize(await deck())).toEqual({ width: 12192000, height: 6858000 });
  });

  it("falls back to 4:3 when a deck declares no size", async () => {
    // Not a deck PowerPoint wrote — but every landing rule divides by these,
    // and 4:3 is what the format itself falls back to. The alternative is an
    // element positioned at NaN, which PowerPoint repairs by dropping it.
    const pkg = await deck({ [PRESENTATION]: presentationPart("") });
    expect(await slideSize(pkg)).toEqual({ width: 9144000, height: 6858000 });
  });

  it("treats a size of zero, or one that is not a number, as no size at all", async () => {
    // Each of the four ways `<p:sldSz>` can be present and useless. A zero
    // width is the sharp one: it is finite, it passes a `Number.isFinite`
    // check on its own, and dividing by it is how a centred element ends up
    // at Infinity.
    const DEFAULT_4_3 = { width: 9144000, height: 6858000 };
    for (const sz of [
      `<p:sldSz cx="0" cy="6858000"/>`,
      `<p:sldSz cx="wide" cy="6858000"/>`,
      `<p:sldSz cx="12192000" cy="0"/>`,
      `<p:sldSz cx="12192000" cy="tall"/>`,
      // No `cx` attribute at all, which reads as zero rather than as missing.
      `<p:sldSz cy="6858000"/>`,
    ]) {
      const pkg = await deck({ [PRESENTATION]: presentationPart(sz) });
      expect(await slideSize(pkg), sz).toEqual(DEFAULT_4_3);
    }
  });
});

describe("the deck the product ships", () => {
  it("reads a real slide's title and body out of the layout it inherits from", async () => {
    /**
     * The synthetic decks above prove each rule in isolation; this one proves
     * the rules add up on a deck the owner drew in PowerPoint, where the
     * placeholders are `ctrTitle` and `subTitle idx="1"`, both carrying
     * `<p:spPr/>` and nothing else, and the walk steps over an embedded object
     * inside `<mc:AlternateContent>` and a hidden picture to reach them.
     *
     * The numbers are `slideLayout1.xml`'s, read off the shipped
     * `template/library-16x9.pptx` on 2026-09-10. If the owner redraws the
     * title slide they change, and the assertion above them is what says they
     * can only have come from the layout.
     */
    const pkg = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
    const xml = await pkg.text(SLIDE);
    expect(xml).toContain('<p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>');
    expect(xml).toContain('<p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>');

    expect(await layoutOf(pkg, SLIDE)).toBe(LAYOUT);
    expect(await framesOf(pkg, SLIDE)).toEqual({
      title: { x: 6340368, y: 2799923, cx: 5510012, cy: 1292225 },
      body: { x: 6340368, y: 3986434, cx: 5510012, cy: 369122 },
    });
    expect(await slideSize(pkg)).toEqual({ width: 12192000, height: 6858000 });
  });
});
