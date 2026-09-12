import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { parseXml } from "../src/core/pptx/xml.js";
import { boxOf, contentCount, occupiedBoxes, rotationOf, topLevelShapes } from "../src/core/catalogue/boxes.js";
import { onlySlide, splice, type SpliceElement } from "../src/core/splice/splice.js";
import { removeElement } from "../src/core/splice/remove.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * Slides that are not shaped the way the splice assumes.
 *
 * The companion to `pptx-malformed.test.ts`, one layer in. That file damages the
 * PACKAGE — content types, relationships, the slide list. This one damages the
 * SLIDE, which is what `splice.ts`, `remove.ts` and `boxes.ts` read and rewrite.
 *
 * Measured 2026-09-12, after `pkg.ts` reached 98.96%: these three were then the
 * thinnest measured code in the repo — `splice.ts` at **82.4% of its branches**,
 * `boxes.ts` at 87.3%, `remove.ts` at 86.2% — and the uncovered lines are all
 * one family of guard. A `<p:cSld>` with no `<p:spTree>`. An `<a:xfrm>` that is
 * not there. `?? ""` on an attribute read. And, four times over,
 * `if (node.nodeType !== 1) continue`.
 *
 * That last one is not exotic at all, which is the point of writing these down.
 * **Any slide saved with the XML indented has text nodes between its elements.**
 * PowerPoint writes its parts without them, so the fixture decks in this suite
 * do too, and every one of those four guards had therefore never run — against
 * a shape of file that any editor, any diff tool and any `xmllint --format`
 * produces. These functions edit a deck the user opened; the deck was written by
 * whatever wrote it.
 *
 * **Two kinds of case live here and they are labelled**, the same way as in
 * `pptx-malformed.test.ts`: writing them meant removing each guard to watch its
 * case go red, and the ones that stayed green are marked BELT AND BRACES with
 * what actually catches them. A case that pins the BEHAVIOUR is worth keeping
 * whichever line delivers it; a comment claiming it proves something it does not
 * is worth nothing.
 */

const NS = [
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join(" ");

const SLIDE2 = "ppt/slides/slide2.xml";
const SLIDE2_RELS = "ppt/slides/_rels/slide2.xml.rels";
const THREE = [{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }];

/** A plain rectangle, as an element's markup: enough to be spliced, and nothing more. */
const BOX: SpliceElement = {
  id: "one-box",
  name: "One box",
  kind: "slide",
  box: { x: 0.1, y: 0.2, w: 0.5, h: 0.4 },
  landing: "layout",
  markup: {
    xml:
      `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="90" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`,
    rels: [],
  },
};

const CATALOGUE = { version: "test-version", carried: {}, theme: {} };
const store = (): Promise<undefined> => Promise.resolve(undefined);

/** A three-slide fixture deck with named parts replaced, or removed with `null`. */
async function craft(edits: Record<string, string | null>): Promise<Uint8Array> {
  const source = await JSZip.loadAsync(await makeDeck(THREE));
  const out = new JSZip();
  for (const [name, file] of Object.entries(source.files)) {
    if (file.dir) continue;
    if (!(name in edits)) {
      out.file(name, await file.async("uint8array"));
      continue;
    }
    const replacement = edits[name];
    if (replacement !== null && replacement !== undefined) out.file(name, replacement);
  }
  return out.generateAsync({ type: "uint8array" });
}

/** The fixture's slide 2, with its `<p:spTree>` contents replaced by `body`. */
async function slideWith(body: string, wrapper?: (tree: string) => string): Promise<Uint8Array> {
  const source = await JSZip.loadAsync(await makeDeck(THREE));
  const xml = await (source.file(SLIDE2) as JSZip.JSZipObject).async("string");
  const open = xml.indexOf("<p:spTree>");
  const close = xml.indexOf("</p:spTree>") + "</p:spTree>".length;
  const tree = `<p:spTree>${body}</p:spTree>`;
  return craft({ [SLIDE2]: xml.slice(0, open) + (wrapper ? wrapper(tree) : tree) + xml.slice(close) });
}

/** The `<p:nvGrpSpPr>` and `<p:grpSpPr>` every `<p:spTree>` opens with. */
const TREE_HEAD =
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;

/** A body placeholder holding one paragraph of one run. */
function placeholder(id: number, text = "Second"): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Body ${id}"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

const onNewSlide = (deck: Uint8Array) =>
  splice({
    deck,
    slide: 1,
    element: BOX,
    options: { target: "new", group: true, colours: "deck" },
    catalogue: CATALOGUE,
    store,
  });

describe("a slide the splice cannot find a shape tree in", () => {
  it("names the part when there is no <p:spTree>", async () => {
    // The `<p:cSld>` is there and empty. Not a slide PowerPoint would write, and
    // the failure has to name which part, because the caller is a task pane
    // whose user is looking at a deck they did not author.
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:sld ${NS}><p:cSld/><p:clrMapOvr><a:overrideClrMapping/></p:clrMapOvr></p:sld>`;
    await expect(onNewSlide(await craft({ [SLIDE2]: xml }))).rejects.toThrow(
      /ppt\/slides\/slide\d+\.xml has no <p:spTree>/,
    );
  });

  it("names the part when there is no <p:cSld> either", async () => {
    // BELT AND BRACES for `spTreeOf`'s `cSld ? child(…) : undefined`, found by
    // trying on 2026-09-12: `cloneSlide` runs first and raises its own,
    // better-aimed sentence, so that ternary's false arm is not what a caller
    // ever meets. What the case pins is that the sentence names the missing
    // ELEMENT and the part — which is all a user of a task pane can act on.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}/>`;
    await expect(onNewSlide(await craft({ [SLIDE2]: xml }))).rejects.toThrow(
      /ppt\/slides\/slide\d+\.xml has no <p:cSld>/,
    );
  });

  it("refuses a removal from such a slide before it ever looks for the tree", async () => {
    /**
     * `remove.ts` carries its own copy of `spTreeOf`, and it CANNOT be reached
     * through `removeElement` — measured 2026-09-12 while trying to.
     *
     * The tag check runs first and needs at least one shape carrying this
     * add-in's tag; a slide with no `<p:spTree>` has no shapes, so it fails
     * there instead, and the clone `spTreeOf` is actually handed is a copy of a
     * slide that demonstrably had a tree. The guard stays — it is what lets the
     * function return an `Element` rather than `Element | undefined` — and this
     * case pins the message the user actually gets, which is the better one:
     * it says what is missing rather than naming a part they cannot see.
     */
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld/></p:sld>`;
    const deck = await craft({ [SLIDE2]: xml });
    await expect(removeElement({ deck, slide: 1, element: "one-box" })).rejects.toThrow(/slide 2 carries no "one-box"/);
  });
});

describe("a slide whose XML somebody has indented", () => {
  /**
   * The case none of the fixtures could produce.
   *
   * PowerPoint writes its parts with no whitespace between elements, so every
   * deck in this suite is one long line and the four `nodeType !== 1` guards had
   * never run. A slide that has been through an editor, a diff tool or
   * `xmllint --format` has a text node between every pair of elements — and the
   * splice walks `childNodes`, not `children`.
   *
   * BELT AND BRACES, all four, and only trying them said so — measured
   * 2026-09-12. Remove any one and these cases stay green, because what a text
   * node meets next already turns it away: it belongs to no namespace the
   * checks below accept, and `child()` finds nothing under a node with no
   * element children. What the cases pin is the OUTCOME on a reformatted slide,
   * which is worth having whichever line delivers it — and which nothing had
   * ever asked for, on a shape of file any editor produces.
   */
  const indent = (tree: string): string => tree.replace(/></g, ">\n  <");

  it("still empties the placeholders of a new slide", async () => {
    // `blank` walks the first paragraph's `childNodes` looking for runs to
    // remove, so on this slide it meets whitespace between every pair.
    const deck = await slideWith(TREE_HEAD + placeholder(4), indent);
    const report = await onNewSlide(deck);
    const pkg = await Pkg.open(report.base64);
    expect(await pkg.text(report.slidePath), "the previous slide's text came with the new one").not.toContain("Second");
    // Blanked and then dropped, which is the whole sequence: `blank` empties the
    // placeholder and `emptyBodyPlaceholders` takes the ghost away. The count is
    // the evidence that the first half happened — an un-emptied placeholder is
    // not empty, so it would survive and this would be 0.
    expect(report.placeholders).toBe(1);
  });

  it("still recognises which shapes are placeholders", async () => {
    // `placeholderIn` walks a shape's own `childNodes` for `<p:nvSpPr>`. Get
    // this wrong and the placeholder reads as ordinary content and is REMOVED,
    // so the new slide loses the layout's own boxes instead of emptying them.
    const deck = await slideWith(TREE_HEAD + placeholder(4), indent);
    const report = await onNewSlide(deck);
    // Miss the `<p:nvSpPr>` and the shape reads as ordinary content, so `blank`
    // REMOVES it outright rather than emptying it — and nothing is left for
    // `emptyBodyPlaceholders` to count. Same end state on screen, different
    // route, and the count is what tells them apart.
    expect(report.placeholders).toBe(1);
  });

  it("still counts what the slide already holds", async () => {
    // `boxes.ts` walks the tree the same way, and its answer decides whether the
    // pane offers "Move to a new slide" — an offer made or withheld on a count
    // that read the whitespace as shapes.
    const deck = await slideWith(TREE_HEAD + placeholder(4), indent);
    const report = await splice({
      deck,
      slide: 1,
      element: BOX,
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: CATALOGUE,
      store,
    });
    expect(report.held, "one placeholder with text on it, not a count of text nodes").toBe(1);
  });
});

describe("placeholders that are not the tidy case", () => {
  it("leaves a placeholder that has no text body alone", async () => {
    // A picture or table placeholder has no `<p:txBody>` to empty. Skipped
    // rather than reached into — and it still has to SURVIVE, because it is a
    // placeholder and the new slide keeps the layout's boxes.
    const bare =
      `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Picture placeholder"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="pic" idx="2"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp>`;
    const report = await onNewSlide(await slideWith(TREE_HEAD + bare));
    const pkg = await Pkg.open(report.base64);
    // A picture placeholder is not a BODY placeholder, so nothing drops it: it
    // survives onto the new slide, which is what a layout's own box should do.
    expect(await pkg.text(report.slidePath)).toContain('type="pic"');
    expect(report.placeholders, "nothing to empty and nothing to drop").toBe(0);
  });

  it("leaves a text body that has no paragraph alone", async () => {
    // `<p:txBody>` with no `<a:p>` is schema-invalid and PowerPoint reports the
    // file as damaged without naming it — so the one thing this must not do is
    // ADD to the damage by reaching into a paragraph that is not there.
    const empty =
      `<p:sp><p:nvSpPr><p:cNvPr id="6" name="Body"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp>`;
    const report = await onNewSlide(await slideWith(TREE_HEAD + empty));
    // It is a body placeholder and it is already empty, so the ghost pass takes
    // it. What matters is that the splice REACHED that pass: reaching into the
    // paragraph that is not there would have raised before it.
    expect(report.placeholders).toBe(1);
  });
});

describe("a slide whose relationships are missing or damaged", () => {
  it("makes a new slide from one that has no relationships part at all", async () => {
    // `blank` strips the notes and comment relationships on the way out. With no
    // `.rels` there is nothing to strip, and the early return is what keeps that
    // from being a read of a part that is not there.
    const deck = await craft({ [SLIDE2_RELS]: null });
    const report = await onNewSlide(deck);
    expect(report.slidePath).toBeTruthy();
    expect(await Pkg.open(report.base64)).toBeTruthy();
  });

  it("skips a relationship that declares no type", async () => {
    // BELT AND BRACES, measured 2026-09-12. The `?? ""` is not what saves this:
    // `null` equals neither the notes type nor any comment type, and
    // `COMMENT_REL_TYPES.includes(null)` is false, so the entry is left alone
    // without it. What the case pins is that ONE malformed entry does not take
    // the whole "as a new slide" path with it.
    const source = await JSZip.loadAsync(await makeDeck(THREE));
    const rels = await (source.file(SLIDE2_RELS) as JSZip.JSZipObject).async("string");
    const damaged = rels.replace(
      "<Relationship ",
      '<Relationship Id="rIdX" Target="../slideLayouts/slideLayout1.xml"/><Relationship ',
    );
    const report = await onNewSlide(await craft({ [SLIDE2_RELS]: damaged }));
    expect(report.slidePath).toBeTruthy();
  });
});

describe("what the element itself can be missing", () => {
  it("splices a catalogue that carries no theme map", async () => {
    // The theme map is what "As in the library" pins colours TO. A catalogue
    // without one is the harvest of a deck whose theme could not be read, and
    // the `?? {}` is what makes that a no-op rather than a raise.
    const deck = await makeDeck(THREE);
    const report = await splice({
      deck,
      slide: 1,
      element: BOX,
      options: { target: "onto", group: true, colours: "library" },
      catalogue: { version: "v1", carried: {} },
      store,
    });
    expect(report.slidePath).toBeTruthy();
  });
});

describe("reducing a package to one slide", () => {
  it("keeps working on a deck whose slide XML is indented", async () => {
    // `onlySlide` is the undo's half of the same walk, and the undo is the one
    // operation whose failure the user cannot work around.
    const deck = await slideWith(TREE_HEAD + placeholder(4), (tree) => tree.replace(/></g, ">\n  <"));
    const only = await onlySlide(deck, 1);
    const pkg = await Pkg.open(only.base64);
    expect(await pkg.slidePaths()).toHaveLength(1);
  });
});

describe("shapes the box reader cannot measure", () => {
  /**
   * `boxes.ts` answers where a shape is, and three different things ask it: the
   * harvest (against the owner's decks), the preview card's grey boxes, and
   * `contentCount`, which decides whether "Move to a new slide" is offered. The
   * last two run against the USER'S deck, so every shape below is one somebody
   * else's software could have written.
   *
   * Called directly rather than through a splice: these are pure functions over
   * a parsed slide, and a case that has to build a package first is a case that
   * proves less and takes longer.
   */
  const W = 12192000;
  const H = 6858000;

  const slide = (body: string): Document =>
    parseXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<p:sld ${NS}><p:cSld><p:spTree>${TREE_HEAD}${body}</p:spTree></p:cSld></p:sld>`,
    );

  const only = (body: string): globalThis.Element => topLevelShapes(slide(body))[0] as globalThis.Element;

  it("cannot measure a shape with no <p:spPr>", () => {
    const bare = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Bare"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>`;
    expect(boxOf(only(bare), W, H)).toBeUndefined();
  });

  it("cannot measure a group with no <p:grpSpPr>", () => {
    // A group is the one kind whose rectangle lives somewhere else, so it has
    // its own arm and its own way of being absent.
    const bare = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr></p:grpSp>`;
    expect(boxOf(only(bare), W, H)).toBeUndefined();
  });

  it("cannot measure an <a:xfrm> that is missing any one of the four numbers", () => {
    // All four, one at a time, because the check is a four-way `or` and a case
    // that drops the whole `<a:ext>` only ever exercises the first of them —
    // which is what the first version of this did.
    const sp = (id: number, xfrm: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Partial"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm>${xfrm}</a:xfrm></p:spPr></p:sp>`;
    const full = { off: `<a:off x="100" y="100"/>`, ext: `<a:ext cx="10" cy="10"/>` };
    expect(boxOf(only(sp(6, full.off)), W, H), "no <a:ext> at all").toBeUndefined();
    expect(boxOf(only(sp(7, full.ext)), W, H), "no <a:off> at all").toBeUndefined();
    expect(boxOf(only(sp(8, `<a:off y="100"/>${full.ext}`)), W, H), "no x").toBeUndefined();
    expect(boxOf(only(sp(9, `<a:off x="100"/>${full.ext}`)), W, H), "no y").toBeUndefined();
    expect(boxOf(only(sp(10, `${full.off}<a:ext cy="10"/>`)), W, H), "no cx").toBeUndefined();
    expect(boxOf(only(sp(11, `${full.off}<a:ext cx="10"/>`)), W, H), "no cy").toBeUndefined();
    // And the whole thing present still measures, so the case is not vacuous.
    expect(boxOf(only(sp(12, full.off + full.ext)), W, H)).toBeDefined();
  });

  it("cannot measure a coordinate that is not a number", () => {
    // `Number("12pt")` is NaN, and a NaN through the arithmetic below would come
    // out as a rectangle whose every edge is NaN — drawn as nothing, silently,
    // rather than reported as unmeasurable.
    const odd =
      `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Odd"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="12pt" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp>`;
    expect(boxOf(only(odd), W, H)).toBeUndefined();
  });

  it("gives a table with no column grid a width of nothing rather than a raise", () => {
    // A table's drawn size is its columns and rows added up, because the frame
    // it sits in lies about its height. With no `<a:tblGrid>` there are no
    // columns to add, and the frame's own width is what is left.
    const table =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm rot="1200000"><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></p:xfrm>` +
      `<a:graphic><a:graphicData><a:tbl><a:tr h="900000"/><a:tr/></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    const box = boxOf(only(table), W, H);
    expect(box, "still measurable, from the frame").toBeDefined();
    // The rotated frame is also what reaches `rotationOf`'s table arm — the two
    // read the same shape and only one of them had ever seen a table.
    const spun = rotationOf(only(table), W, H);
    expect(spun?.deg).toBe(20);
  });

  it("adds up only the rows that say how tall they are", () => {
    // `?? 0` on each row and column. A `<a:tr>` with no `h` is what a table
    // still being laid out looks like, and one of them must not make the whole
    // height NaN.
    const table =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></p:xfrm>` +
      `<a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="3000000"/><a:gridCol/></a:tblGrid>` +
      `<a:tr h="900000"/><a:tr/></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    const box = boxOf(only(table), W, H);
    expect(box?.w, "one column of 3000000 EMU on a 12192000 EMU slide").toBeCloseTo(3000000 / W, 6);
    expect(box?.h).toBeCloseTo(900000 / H, 6);
  });

  it("counts an unmeasurable shape as content, because it is on the slide", () => {
    // The direction that matters. `contentCount` decides whether the pane offers
    // "Move to a new slide"; a shape it cannot measure is still something the
    // element would land on top of, so it counts. Withholding the offer is the
    // failure a reader would not notice.
    const bare = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Bare"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp>`;
    expect(contentCount(slide(bare), W, H)).toBe(1);
    // And it draws no grey box for it, because there is no rectangle to draw.
    expect(occupiedBoxes(slide(bare), W, H)).toEqual([]);
  });
});

describe("taking an element back off a slide that has been reformatted", () => {
  /**
   * A removal reads the slide the same way the splice does, and it is the
   * operation that runs once per slide in a loop — so a walk that miscounts
   * costs the user a slide it reported as cleaned and did not clean.
   *
   * The deck here is built by an actual insert, because the tag the removal
   * looks for is written BY the insert. Indenting the result afterwards is the
   * whole point: the same bytes, through a formatter.
   */
  async function insertedThenIndented(): Promise<Uint8Array> {
    const report = await splice({
      deck: await makeDeck(THREE),
      slide: 1,
      element: BOX,
      options: { target: "onto", group: false, colours: "deck" },
      catalogue: CATALOGUE,
      store,
    });
    const source = await JSZip.loadAsync(Buffer.from(report.base64, "base64"));
    const out = new JSZip();
    for (const [name, file] of Object.entries(source.files)) {
      if (file.dir) continue;
      const bytes = await file.async("uint8array");
      if (name !== report.slidePath) {
        out.file(name, bytes);
        continue;
      }
      out.file(name, Buffer.from(bytes).toString("utf8").replace(/></g, ">\n  <"));
    }
    return out.generateAsync({ type: "uint8array" });
  }

  it("still finds the shapes it put there", async () => {
    // `strip` walks `childNodes`, so every whitespace node between two shapes
    // reaches `idOf(shape)` unless the guard drops it first. BELT AND BRACES
    // like the others: `idOf` finds no `<p:cNvPr>` under a text node and the
    // `namespaceURI !== P_NS` line after it turns the node away regardless.
    const report = await removeElement({ deck: await insertedThenIndented(), slide: 0, element: BOX.id });
    expect(report.removed, "the shape the insert tagged").toBe(1);
    const pkg = await Pkg.open(report.base64);
    expect(await pkg.text(report.slidePath)).not.toContain('name="Box"');
  });
});

describe("elements that are not one plain shape", () => {
  it("sends a marker AROUND the selection rather than onto it", async () => {
    // `wraps` is the marker's whole behaviour (`docs/DESIGN.md` section 5) and
    // `splice-landing.test.ts` proves the rule against `place` directly. What
    // had never run is the splice PASSING it: the flag is threaded through one
    // conditional spread, and a spread that never fires is a feature that never
    // reaches the engine.
    const marker: SpliceElement = { ...BOX, id: "markering-1", kind: "part", landing: "cursor", wraps: true };
    const selection = { x: 4000000, y: 2000000, cx: 2000000, cy: 1000000 };
    const report = await splice({
      deck: await makeDeck(THREE),
      slide: 1,
      element: marker,
      options: { target: "onto", group: false, colours: "deck" },
      catalogue: CATALOGUE,
      store,
      selection,
    });
    // AROUND it: a little bigger than the shape it wraps, and centred on it.
    // The measurement, 2026-09-12: with the flag the marker comes out
    // 2 320 000 EMU wide — the selection plus eight per cent of air on each
    // side — and without it, 6 096 000, the element's own authored width
    // scaled to the slide, which ignores the selection's size completely. So
    // "a bit bigger" is the assertion, and an upper bound is what makes it one:
    // `toBeGreaterThan` alone passes either way.
    const middle = (r: { x: number; cx: number }): number => r.x + r.cx / 2;
    expect(report.landed.cx).toBeGreaterThan(selection.cx);
    expect(report.landed.cx, "sized to the selection, not to the element").toBeLessThan(selection.cx * 1.5);
    expect(middle(report.landed), "centred on what it wraps").toBeCloseTo(middle(selection), 6);
  });

  it("tags what it can inside a group and leaves the rest alone", async () => {
    /**
     * A group holding something with no `<p:nvPr>` — which is how a modern chart
     * sits on a slide, inside `<mc:AlternateContent>`.
     *
     * The rule the guard implements is in `splice.ts`: skipped rather than
     * refused, because failing an insert that works today would be a worse trade
     * than losing a tag on an ungroup that may never happen. Nothing was holding
     * it to that, so an insert of any grouped element carrying a chart was
     * relying on a line no case had run.
     */
    const grouped: SpliceElement = {
      ...BOX,
      id: "grouped-1",
      markup: {
        xml:
          `<p:grpSp ${NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<p:nvGrpSpPr><p:cNvPr id="80" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/>` +
          `<a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="500000"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="81" name="Inner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp>` +
          `<mc:AlternateContent><mc:Fallback/></mc:AlternateContent></p:grpSp>`,
        rels: [],
      },
    };
    const report = await splice({
      deck: await makeDeck(THREE),
      slide: 1,
      element: grouped,
      options: { target: "onto", group: false, colours: "deck" },
      catalogue: CATALOGUE,
      store,
    });
    // The insert works, which is the trade the rule chose.
    const pkg = await Pkg.open(report.base64);
    expect(await pkg.text(report.slidePath)).toContain("AlternateContent");
  });
});
