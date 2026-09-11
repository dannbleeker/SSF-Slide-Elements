import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { REL_TYPE } from "../src/core/pptx/parts.js";
import {
  TAG_CATALOGUE,
  TAG_ELEMENT,
  mergeTagPart,
  nextTagNumber,
  readShapeTags,
  tagPartXml,
  writeShapeTags,
} from "../src/core/pptx/tags.js";
import { P_NS, R_NS, child, children, elements, parseXml } from "../src/core/pptx/xml.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

/**
 * The stamp this add-in leaves on a shape, written into the package rather than
 * asked of Office.js.
 *
 * It is what "Used in this deck" and "Remove from N slides" are read back from,
 * so a defect here is not a broken file — it is a pane that quietly cannot see
 * an element it inserted, on a deck that opens perfectly well. That shape of
 * failure is why almost every case below asserts on the MARKUP the writer
 * produced, by the path the schema names, rather than on what `readShapeTags`
 * says about it: reader and writer share `nvPrOf`, so a round trip through both
 * agrees with itself whatever either one does.
 *
 * `src/core/pptx/tags.ts` was ported from SSF-Merge, which tags a SLIDE, and
 * this one tags a SHAPE. Everything that changed in the move is a question
 * about where `<p:nvPr>` is — five containers, one per shape kind, and a group
 * that has a second one inside it — so those cases are the bulk of the file.
 * The incidents the module's comments narrate (`Ben &amp;amp; Jerry`, the
 * dropped foreign tag, the damaged file) happened in that engine, and the cases
 * named after them are here because the code that produced them is here.
 */

const SLIDE = "ppt/slides/slide1.xml";
/** The part the fixture's `shapeTags` option writes: another add-in's, already in the deck. */
const VENDOR_PART = "ppt/tags/tag9.xml";
/** What `[Content_Types].xml` has to say about a tag part, or the file opens as damaged. */
const TAGS_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

async function deck(slides: SlideSpec[]): Promise<Pkg> {
  return Pkg.open(await makeDeck(slides));
}

/** Where `<p:nvPr>` lives, per shape kind. The whole of what the port had to get right. */
const CONTAINER: Record<string, string> = {
  sp: "nvSpPr",
  grpSp: "nvGrpSpPr",
  pic: "nvPicPr",
  graphicFrame: "nvGraphicFramePr",
  cxnSp: "nvCxnSpPr",
};

/**
 * A shape's own `<p:nvPr>`, reached by the path the schema names for that KIND.
 *
 * Spelled out rather than borrowed from the module, deliberately. `nvPrOf`
 * there searches for the first `<p:nvPr>` one level down, so a test using it
 * would agree with a writer that had put the stamp on the wrong element — which
 * is exactly the group case below.
 */
function nvPrOn(shape: Element): Element {
  const container = CONTAINER[shape.localName ?? ""];
  if (!container) throw new Error(`the test does not know where <p:nvPr> lives in <p:${shape.localName}>`);
  const props = child(shape, P_NS, container);
  const nvPr = props ? child(props, P_NS, "nvPr") : undefined;
  if (!nvPr) throw new Error(`<p:${shape.localName}> has no <p:${container}><p:nvPr>`);
  return nvPr;
}

/** The `<p:tags>` reference on a shape, by the one path a reader is allowed to look down. */
function referenceOn(shape: Element): Element | undefined {
  const list = child(nvPrOn(shape), P_NS, "custDataLst");
  return list ? child(list, P_NS, "tags") : undefined;
}

/** Every `<p:custDataLst>` ANYWHERE under a shape, so a list left on a child is counted too. */
function everyCustomDataList(shape: Element): Element[] {
  return elements(shape, P_NS, "custDataLst");
}

/** The element children of a node, by local name: what the schema orders. */
function order(parent: Element): string[] {
  return Array.from(parent.childNodes)
    .filter((n) => n.nodeType === 1)
    .map((n) => (n as Element).localName ?? "");
}

async function shapeTree(pkg: Pkg, slidePath = SLIDE): Promise<Element> {
  const doc = await pkg.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) throw new Error(`the fixture changed shape: ${slidePath} has no <p:spTree>`);
  return spTree;
}

async function topLevel(pkg: Pkg, slidePath = SLIDE): Promise<Element[]> {
  return Array.from((await shapeTree(pkg, slidePath)).childNodes).filter((n) => n.nodeType === 1) as Element[];
}

/** A top-level shape by the name in its `<p:cNvPr>`, which is how the fixtures below are read. */
async function shapeNamed(pkg: Pkg, name: string, slidePath = SLIDE): Promise<Element> {
  for (const shape of await topLevel(pkg, slidePath)) {
    for (const container of Array.from(shape.childNodes)) {
      if (container.nodeType !== 1) continue;
      const cNvPr = child(container as Element, P_NS, "cNvPr");
      if (cNvPr?.getAttribute("name") === name) return shape;
    }
  }
  throw new Error(`the fixture has no top-level shape named "${name}"`);
}

/** The package path a shape's reference resolves to, or undefined if it resolves to nothing. */
async function partFor(pkg: Pkg, shape: Element, slidePath = SLIDE): Promise<string | undefined> {
  const ref = referenceOn(shape);
  const rId = ref?.getAttributeNS(R_NS, "id") ?? ref?.getAttribute("r:id");
  if (!rId) return undefined;
  return pkg.relTarget(slidePath, rId);
}

/** The tags in a part, read off the bytes the package would ship rather than a cached document. */
async function entriesIn(pkg: Pkg, part: string): Promise<[string, string][]> {
  return elements(parseXml(await pkg.text(part)), P_NS, "tag").map((tag) => [
    tag.getAttribute("name") ?? "",
    tag.getAttribute("val") ?? "",
  ]);
}

const OURS: [string, string][] = [
  [TAG_ELEMENT, "hvid-kasse-2x1-vertikale"],
  [TAG_CATALOGUE, "2026-09-10"],
];

/** A group with one shape inside it: the case the slide-level writer got wrong. */
const GROUP =
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="Gruppe 50"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="51" name="Inde i gruppen"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
  `</p:grpSp>`;

const PICTURE =
  `<p:pic><p:nvPicPr><p:cNvPr id="60" name="Billede 60"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>`;

/** The same picture as a formatter or a hand edit leaves it, with whitespace between every element. */
const PRETTY_PICTURE = [
  "\n  <p:pic>",
  "\n    <p:nvPicPr>",
  '\n      <p:cNvPr id="61" name="Billede 61"/>',
  "\n      <p:cNvPicPr/>",
  "\n      <p:nvPr/>",
  "\n    </p:nvPicPr>",
  "\n    <p:blipFill><a:stretch><a:fillRect/></a:stretch></p:blipFill>",
  "\n    <p:spPr/>",
  "\n  </p:pic>\n  ",
].join("");

/** A table: `p:graphicFrame` is also how a chart and SmartArt arrive. */
const FRAME =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="70" name="Tabel 70"/><p:cNvGraphicFramePr/>` +
  `<p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic></p:graphicFrame>`;

const CONNECTOR =
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="80" name="Lige forbindelse 80"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
  `<p:spPr/></p:cxnSp>`;

/** A shape whose `<p:nvPr>` already carries the extension list the schema puts LAST. */
const WITH_EXTENSION_LIST =
  `<p:sp><p:nvSpPr><p:cNvPr id="90" name="Med extLst"/><p:cNvSpPr/>` +
  `<p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"/></p:extLst></p:nvPr></p:nvSpPr>` +
  `<p:spPr/></p:sp>`;

/** A custom-data list holding only `<p:custData>`: legal, common, and no tag reference at all. */
const WITH_CUSTOM_DATA =
  `<p:sp><p:nvSpPr><p:cNvPr id="91" name="Med custData"/><p:cNvSpPr/>` +
  `<p:nvPr><p:custDataLst><p:custData r:id="rId1"/></p:custDataLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`;

const REFERENCE_WITHOUT_ID =
  `<p:sp><p:nvSpPr><p:cNvPr id="92" name="Uden r-id"/><p:cNvSpPr/>` +
  `<p:nvPr><p:custDataLst><p:tags/></p:custDataLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`;

const REFERENCE_TO_NOTHING =
  `<p:sp><p:nvSpPr><p:cNvPr id="93" name="Uden relation"/><p:cNvSpPr/>` +
  `<p:nvPr><p:custDataLst><p:tags r:id="rId404"/></p:custDataLst></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`;

/** Two shapes no PowerPoint wrote: one whose `<p:cNvPr>` states no id, one carrying none at all. */
const WITHOUT_SHAPE_ID = `<p:sp><p:nvSpPr><p:cNvPr name="Uden id"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;
const WITHOUT_SHAPE_PROPS = `<p:sp><p:nvSpPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;

/** A deck whose one shape carries another add-in's tag part, holding exactly these tags. */
async function taggedByAnother(foreign: string): Promise<Pkg> {
  const pkg = await deck([{ paragraphs: [["a"]], shapeTags: true }]);
  pkg.setText(
    VENDOR_PART,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:tagLst xmlns:p="${P_NS}">${foreign}</p:tagLst>`,
  );
  return pkg;
}

describe("the keys a shape is stamped with", () => {
  it("names the element and the catalogue, uppercase, as PowerPoint stores tag keys", () => {
    /**
     * Pinned as literals rather than derived, because these two strings are
     * already in decks on disk. Renaming either one orphans every shape a
     * shipped release stamped: the pane would report an empty deck and offer no
     * way to remove what is plainly on the slide.
     */
    expect(TAG_ELEMENT).toBe("SSF_SLIDE_ELEMENT");
    expect(TAG_CATALOGUE).toBe("SSF_SLIDE_ELEMENTS_CATALOGUE");
    expect([TAG_ELEMENT, TAG_CATALOGUE].map((k) => k.toUpperCase())).toEqual([TAG_ELEMENT, TAG_CATALOGUE]);
  });
});

describe("the tag part's own markup", () => {
  it("declares itself the way PowerPoint declares a part", () => {
    const xml = tagPartXml(OURS);
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')).toBe(true);
    const root = parseXml(xml).documentElement;
    expect(root.localName).toBe("tagLst");
    expect(root.namespaceURI).toBe(P_NS);
  });

  it("writes one tag per entry, in the order it was given them", () => {
    const doc = parseXml(tagPartXml(OURS));
    expect(elements(doc, P_NS, "tag").map((t) => [t.getAttribute("name"), t.getAttribute("val")])).toEqual(OURS);
  });

  it("writes an empty list when there is nothing to say, rather than broken markup", () => {
    // Reachable through `mergeTagPart` on a part whose every tag we own and
    // whose caller passed nothing: the part still has to parse.
    expect(elements(parseXml(tagPartXml([])), P_NS, "tag")).toHaveLength(0);
  });

  it("escapes the five markup characters, so a value carrying one is still one attribute", () => {
    const value = `Ben & Jerry's <b>"best"</b>`;
    const doc = parseXml(tagPartXml([["A&B", value]]));
    const tag = elements(doc, P_NS, "tag")[0];
    expect(tag?.getAttribute("name")).toBe("A&B");
    expect(tag?.getAttribute("val")).toBe(value);
  });

  it("escapes the whitespace an XML parser would otherwise read back as a space", () => {
    /**
     * The half that is not obvious. Attribute-value normalisation turns a
     * literal newline, carriage return or tab inside an attribute into a SPACE,
     * so writing them literally loses them on the FIRST insert and looks stable
     * ever after — the shape of defect that never gets reported.
     *
     * It cannot reach our own two keys, which have no whitespace in them. It
     * reaches a foreign tag, which the merge below carries through untouched.
     */
    const value = "line one\nline two\rcolumn\tcolumn";
    const xml = tagPartXml([["FOREIGN", value]]);
    expect(xml).toContain("&#10;");
    expect(xml).toContain("&#13;");
    expect(xml).toContain("&#9;");
    expect(elements(parseXml(xml), P_NS, "tag")[0]?.getAttribute("val")).toBe(value);
  });

  it("replaces a character XML cannot carry at all, rather than writing a part PowerPoint refuses", () => {
    /**
     * Escaping was the whole of this once and it is not enough: `&#11;` is
     * exactly as ill-formed as the byte, so a C0 control or a lone surrogate in
     * a foreign tag produced a part PowerPoint reports as a damaged file with
     * nothing naming the cause. `xmlSafe` runs FIRST, because escaping a
     * character that may not be written is writing it.
     */
    const xml = tagPartXml([["FOREIGN", "a\u000Bb\u0000c\uD800d"]]);
    expect(xml).not.toContain("&#11;");
    expect(elements(parseXml(xml), P_NS, "tag")[0]?.getAttribute("val")).toBe("a b c d");
  });
});

describe("choosing a name for the new part", () => {
  it("never answers a name the package already holds", async () => {
    /**
     * Blind use of `tag1.xml` is the trap: a deck that already carries one has
     * it overwritten and every shape pointing at it silently loses its tags.
     * The library's own elements make that certain rather than theoretical:
     * six of the 117 harvested 16:9 elements arrive carrying tag
     * relationships, 78 of them between the six, and one carries 41 alone.
     */
    const pkg = await deck([{ paragraphs: [["a"]], shapeTags: true }]);
    expect(pkg.has(VENDOR_PART), "the fixture stopped writing a vendor tag part").toBe(true);
    const n = nextTagNumber(pkg);
    expect(pkg.has(`ppt/tags/tag${n}.xml`)).toBe(false);
    expect(n).toBeGreaterThan(9);
  });

  it("answers one past the highest, not the first gap", async () => {
    // A package holding tag2 and tag5 and no tag1 answers 6, where a walk from
    // 1 answered 1. Both are names the package does not hold, which is all this
    // has to guarantee — but only one of them is what `Pkg`'s counter promises.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    pkg.setText("ppt/tags/tag2.xml", tagPartXml([]));
    pkg.setText("ppt/tags/tag5.xml", tagPartXml([]));
    expect(nextTagNumber(pkg)).toBe(6);
  });

  it("starts at one in a deck that has never been tagged", async () => {
    expect(nextTagNumber(await deck([{ paragraphs: [["a"]] }]))).toBe(1);
  });

  it("keeps answering a free name as parts are written", async () => {
    // The counter is memoised, so a part written since it was first asked has
    // to reach it — otherwise the second shape stamped in one insert overwrites
    // the first shape's part.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const first = nextTagNumber(pkg);
    pkg.setText(`ppt/tags/tag${first}.xml`, tagPartXml(OURS));
    const second = nextTagNumber(pkg);
    expect(second).not.toBe(first);
    expect(pkg.has(`ppt/tags/tag${second}.xml`)).toBe(false);
  });
});

describe("stamping a shape that has never been tagged", () => {
  it("writes the part, declares it, relates it from the slide, and points the shape at it", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const shape = await shapeNamed(pkg, "Body");
    await writeShapeTags(pkg, SLIDE, shape, OURS);

    // The reference, down the one path `CT_ApplicationNonVisualDrawingProps`
    // allows: nvPr, then the custom data list, then `<p:tags>`.
    const ref = referenceOn(shape);
    expect(ref, "the shape has no <p:nvPr><p:custDataLst><p:tags>").toBeDefined();
    const rId = ref?.getAttributeNS(R_NS, "id");
    expect(rId).toMatch(/^rId\d+$/);

    // The relationship it names, and what that resolves to.
    const part = await pkg.relTarget(SLIDE, rId ?? "");
    expect(part).toBe("ppt/tags/tag1.xml");
    expect(pkg.has(part ?? "")).toBe(true);
    expect(await entriesIn(pkg, part ?? "")).toEqual(OURS);

    // The relationship is a TAGS relationship, and the target is written
    // relative the way PowerPoint writes one.
    const rels = await pkg.text(Pkg.relsPathFor(SLIDE));
    expect(rels).toContain(`Id="${rId}" Type="${REL_TYPE.tags}" Target="../tags/tag1.xml"`);

    // Undeclared, the file opens as damaged and PowerPoint does not say which
    // part it could not classify.
    expect(await pkg.text("[Content_Types].xml")).toContain(
      `<Override PartName="/ppt/tags/tag1.xml" ContentType="${TAGS_TYPE}"/>`,
    );
  });

  it("puts the custom data list before an extension list, which is the order the schema states", async () => {
    /**
     * `CT_ApplicationNonVisualDrawingProps` orders its children
     * `ph?, (audio|video)?, custDataLst?, extLst?`. Appended blindly the list
     * lands AFTER the extension list, and PowerPoint reports the file as
     * damaged without naming which part — from a deck the add-in produced, on
     * an insert that reported success.
     */
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [WITH_EXTENSION_LIST] }]);
    const shape = await shapeNamed(pkg, "Med extLst");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect(order(nvPrOn(shape))).toEqual(["custDataLst", "extLst"]);
  });

  it("puts it after a placeholder claim, which the same rule orders first", async () => {
    // The other side of that ordering, and the commoner shape: a title
    // placeholder's `<p:nvPr>` opens with `<p:ph>`, and the list goes after it.
    const pkg = await deck([{ paragraphs: [["a"]], title: "Kasser" }]);
    const shape = await shapeNamed(pkg, "Title 1");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect(order(nvPrOn(shape))).toEqual(["ph", "custDataLst"]);
  });

  it("adds its reference to a custom data list the shape already had", async () => {
    // `CT_CustomerDataList` allows one `<p:tags>` and any number of
    // `<p:custData>`, so a second LIST is as invalid as a second reference —
    // and a list holding only `<p:custData>` is legal and common.
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [WITH_CUSTOM_DATA] }]);
    const shape = await shapeNamed(pkg, "Med custData");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    const lists = everyCustomDataList(shape);
    expect(lists).toHaveLength(1);
    expect(order(lists[0] as Element)).toEqual(["custData", "tags"]);
  });

  it("gives two shapes on one slide a part each", async () => {
    // Both are stamped in one insert, and the second must not be handed the
    // first one's part name.
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [PICTURE] }]);
    const body = await shapeNamed(pkg, "Body");
    const picture = await shapeNamed(pkg, "Billede 60");
    await writeShapeTags(pkg, SLIDE, body, OURS);
    await writeShapeTags(pkg, SLIDE, picture, OURS);
    const parts = [await partFor(pkg, body), await partFor(pkg, picture)];
    expect(new Set(parts).size, "the two shapes share one tag part").toBe(2);
    for (const part of parts) expect(await entriesIn(pkg, part ?? "")).toEqual(OURS);
  });
});

describe("stamping a shape another add-in has already tagged", () => {
  it("merges into the part that is there rather than adding a second reference", async () => {
    /**
     * `CT_CustomerDataList` allows at most one `<p:tags>` child, so a shape
     * that already has a tag part must have its entries appended. A shape
     * harvested out of the library can: the owner's decks were built in
     * PowerPoint, and six of the 117 elements arrive already tagged — a
     * minority, but one of them carries 41 relationships, so this path is
     * exercised rather than theoretical.
     */
    const pkg = await taggedByAnother(`<p:tag name="VENDOR" val="do not delete"/>`);
    const shape = await shapeNamed(pkg, "Body");
    const before = referenceOn(shape)?.getAttributeNS(R_NS, "id");
    const tagParts = () => pkg.partNames().filter((p) => p.startsWith("ppt/tags/"));
    expect(tagParts()).toEqual([VENDOR_PART]);
    await writeShapeTags(pkg, SLIDE, shape, OURS);

    expect(elements(shape, P_NS, "tags"), "a second <p:tags> is schema-invalid").toHaveLength(1);
    expect(referenceOn(shape)?.getAttributeNS(R_NS, "id"), "the shape was repointed").toBe(before);
    expect(await partFor(pkg, shape)).toBe(VENDOR_PART);
    expect(tagParts(), "a second tag part was written anyway").toEqual([VENDOR_PART]);
    expect(await entriesIn(pkg, VENDOR_PART)).toEqual([["VENDOR", "do not delete"], ...OURS]);
  });

  it("keeps a foreign value with an ampersand in it intact through two round trips", async () => {
    /**
     * The case that broke the sibling. Its regex read attribute VALUES as raw
     * source, so `val="Ben &amp; Jerry"` came back with the entity intact and
     * was escaped again on write: one merge produced `Ben &amp;amp; Jerry`, two
     * produced `Ben &amp;amp;amp; Jerry`, and a reader saw the literal `&amp;`
     * on screen. The parser decodes, `tagPartXml` encodes exactly once, and
     * that round trip is the whole of what makes repeated inserts stable.
     */
    const value = `Ben & Jerry's <b>"best"</b>`;
    const pkg = await taggedByAnother(
      `<p:tag name="VENDOR" val="Ben &amp; Jerry&apos;s &lt;b&gt;&quot;best&quot;&lt;/b&gt;"/>`,
    );
    const shape = await shapeNamed(pkg, "Body");
    expect(await entriesIn(pkg, VENDOR_PART), "the fixture does not hold what this case is about").toEqual([
      ["VENDOR", value],
    ]);

    await writeShapeTags(pkg, SLIDE, shape, OURS);
    await writeShapeTags(pkg, SLIDE, shape, OURS);

    expect(await entriesIn(pkg, VENDOR_PART)).toEqual([["VENDOR", value], ...OURS]);
    expect(await pkg.text(VENDOR_PART)).not.toContain("&amp;amp;");
  });

  it("keeps a foreign tag PowerPoint spelled its own legal way", async () => {
    /**
     * The regex insisted on one attribute order and a self-closing tag, so
     * `val` before `name`, single quotes and a separate closing tag matched
     * nothing at all and the foreign tag was DROPPED — somebody else's data,
     * gone from a deck this add-in only meant to add to.
     */
    const pkg = await taggedByAnother(`<p:tag val='keep me' name='VENDOR_ALT'></p:tag>`);
    const shape = await shapeNamed(pkg, "Body");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect(await entriesIn(pkg, VENDOR_PART)).toEqual([["VENDOR_ALT", "keep me"], ...OURS]);
  });

  it("replaces its own entry on a second insert rather than writing it twice", async () => {
    // Re-inserting the same element, or upgrading the catalogue behind a shape
    // already stamped: one key must hold one value, or the reader takes the
    // first and reports a version the deck no longer carries.
    const pkg = await taggedByAnother(`<p:tag name="VENDOR" val="do not delete"/>`);
    const shape = await shapeNamed(pkg, "Body");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    await writeShapeTags(pkg, SLIDE, shape, [
      [TAG_ELEMENT, "bred-bjaelke"],
      [TAG_CATALOGUE, "2026-12-01"],
    ]);
    expect(await entriesIn(pkg, VENDOR_PART)).toEqual([
      ["VENDOR", "do not delete"],
      [TAG_ELEMENT, "bred-bjaelke"],
      [TAG_CATALOGUE, "2026-12-01"],
    ]);
  });
});

describe("merging a tag part on its own", () => {
  it("keeps what it does not own and replaces what it does", () => {
    const merged = mergeTagPart(
      tagPartXml([
        ["VENDOR", "theirs"],
        [TAG_ELEMENT, "old"],
      ]),
      [[TAG_ELEMENT, "new"]],
    );
    expect(elements(parseXml(merged), P_NS, "tag").map((t) => [t.getAttribute("name"), t.getAttribute("val")])).toEqual(
      [
        ["VENDOR", "theirs"],
        [TAG_ELEMENT, "new"],
      ],
    );
  });

  it("drops a tag with no name rather than carrying it through as a nameless one", () => {
    // `getAttribute` answers null for an absent name, and a tag written back
    // with `name="null"` is a key nothing can ever match or clean up.
    const merged = mergeTagPart(`<p:tagLst xmlns:p="${P_NS}"><p:tag val="orphan"/></p:tagLst>`, OURS);
    expect(merged).not.toContain("orphan");
    expect(merged).not.toContain("null");
  });

  it("keeps a foreign tag that states no value at all, as an empty one", () => {
    const merged = mergeTagPart(`<p:tagLst xmlns:p="${P_NS}"><p:tag name="VENDOR"/></p:tagLst>`, OURS);
    expect(elements(parseXml(merged), P_NS, "tag")[0]?.getAttribute("val")).toBe("");
  });
});

describe("a reference that leads nowhere", () => {
  /**
   * A relationship can point at a part that is not in the package — a deck
   * another tool wrote, or one PowerPoint repaired by dropping the part and
   * leaving the reference. `pkg.text` throws by name for a missing part, so a
   * shape like that would kill the whole insert, while the reader guards with
   * the same test and answers nothing. A reader that degrades and a writer that
   * throws on the same markup is the pair worth never shipping.
   */
  async function danglingDeck(): Promise<{ pkg: Pkg; shape: Element; rId: string }> {
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [REFERENCE_TO_NOTHING] }]);
    const shape = await shapeNamed(pkg, "Uden relation");
    const rId = await pkg.addRel(SLIDE, REL_TYPE.tags, "../tags/tag77.xml");
    referenceOn(shape)?.setAttributeNS(R_NS, "r:id", rId);
    expect(pkg.has("ppt/tags/tag77.xml"), "the fixture accidentally holds the part").toBe(false);
    return { pkg, shape, rId };
  }

  it("replaces a reference to a part the package does not hold, rather than throwing", async () => {
    const { pkg, shape } = await danglingDeck();
    await writeShapeTags(pkg, SLIDE, shape, OURS);

    // Exactly one, because the fall-through appends into the SAME custom data
    // list: leaving the dead reference produces two, which is schema-invalid,
    // and a reader takes the FIRST — so the insert's own tag would be invisible
    // to every reader of it, on a deck that opens perfectly well.
    expect(elements(shape, P_NS, "tags")).toHaveLength(1);
    const part = await partFor(pkg, shape);
    expect(part).not.toBe("ppt/tags/tag77.xml");
    expect(await entriesIn(pkg, part ?? "")).toEqual(OURS);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([
      { element: "hvid-kasse-2x1-vertikale", catalogue: "2026-09-10", shapeId: "93" },
    ]);
  });

  it("leaves the dangling relationship exactly where it was", async () => {
    /**
     * Deliberate. It was in the deck before this ran, and removing
     * relationships is the operation that has twice produced damage in the
     * sibling: an id freed by a delete is handed to the next thing that asks
     * for one, and the vendor's shape came out of the run pointing at OUR
     * metadata — a reference that still resolves, to somebody else's data.
     */
    const { pkg, shape, rId } = await danglingDeck();
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    const rels = await pkg.text(Pkg.relsPathFor(SLIDE));
    expect(rels).toContain(`Id="${rId}"`);
    expect(rels).toContain(`Target="../tags/tag77.xml"`);
  });

  it("stamps a shape whose reference names a relationship the slide does not have", async () => {
    // `relTarget` answers nothing, which is not the same as a target that
    // resolves to a missing part, and reaches the same fall-through.
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [REFERENCE_TO_NOTHING] }]);
    const shape = await shapeNamed(pkg, "Uden relation");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect(elements(shape, P_NS, "tags")).toHaveLength(1);
    expect(await entriesIn(pkg, (await partFor(pkg, shape)) ?? "")).toEqual(OURS);
  });

  it("stamps a shape whose reference states no relationship id at all", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [REFERENCE_WITHOUT_ID] }]);
    const shape = await shapeNamed(pkg, "Uden r-id");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect(elements(shape, P_NS, "tags")).toHaveLength(1);
    expect(await entriesIn(pkg, (await partFor(pkg, shape)) ?? "")).toEqual(OURS);
  });
});

describe("which shape the writer stamps", () => {
  it("finds the non-visual properties of a group, a picture, a frame and a connector", async () => {
    /**
     * Each kind wraps `<p:nvPr>` in a differently named container, and the port
     * from the slide-level writer is exactly this question. An element out of
     * the library is any of them: the collection holds tables and pictures, and
     * a grouped insert lands as one `<p:grpSp>`.
     */
    const pkg = await deck([{ paragraphs: [["a"]], noBody: true, shapes: [GROUP, PICTURE, FRAME, CONNECTOR] }]);
    for (const name of ["Gruppe 50", "Billede 60", "Tabel 70", "Lige forbindelse 80"]) {
      const shape = await shapeNamed(pkg, name);
      await writeShapeTags(pkg, SLIDE, shape, [
        [TAG_ELEMENT, name],
        [TAG_CATALOGUE, "2026-09-10"],
      ]);
      expect(referenceOn(shape), `<p:${shape.localName}> was not stamped on its own <p:nvPr>`).toBeDefined();
    }
    expect((await readShapeTags(pkg, SLIDE)).map((t) => [t.element, t.shapeId])).toEqual([
      ["Gruppe 50", "50"],
      ["Billede 60", "60"],
      ["Tabel 70", "70"],
      ["Lige forbindelse 80", "80"],
    ]);
  });

  it("never stamps a shape inside the group it was asked to stamp", async () => {
    /**
     * `element` walks DESCENDANTS, so a search that is not scoped to one level
     * finds the `<p:nvPr>` of a shape INSIDE the group and puts the element's
     * identity there. The group is then untagged and a child claims to be the
     * element: "Remove from N slides" would take the child out and leave the
     * rest of the group on the slide.
     */
    const pkg = await deck([{ paragraphs: [["a"]], noBody: true, shapes: [GROUP] }]);
    const group = await shapeNamed(pkg, "Gruppe 50");
    await writeShapeTags(pkg, SLIDE, group, OURS);

    expect(everyCustomDataList(group), "a list was left on a shape inside the group").toHaveLength(1);
    const inner = children(group, P_NS, "sp")[0];
    expect(inner, "the fixture stopped putting a shape inside the group").toBeDefined();
    expect(child(nvPrOn(inner as Element), P_NS, "custDataLst")).toBeUndefined();
    expect(referenceOn(group)).toBeDefined();
  });

  it("refuses a top-level thing with no non-visual properties, naming the slide", async () => {
    /**
     * A modern chart sits on the slide as `<mc:AlternateContent>`, whose
     * children are `<mc:Choice>` and `<mc:Fallback>` and which has no
     * `<p:nvPr>` of its own at all. Refusing by name is what tells a caller
     * which slide to look at; silently skipping would ship an untagged element
     * the pane can never see again.
     */
    const pkg = await deck([{ paragraphs: [["a"]], modernChart: { series: "Sales", categories: ["A"] } }]);
    const alternate = (await topLevel(pkg)).find((s) => s.localName === "AlternateContent");
    expect(alternate, "the fixture stopped writing <mc:AlternateContent>").toBeDefined();
    await expect(writeShapeTags(pkg, SLIDE, alternate as Element, OURS)).rejects.toThrow(
      /a top-level shape in ppt\/slides\/slide1\.xml has no <p:nvPr> to tag/,
    );
  });

  it("refuses a shape that belongs to no document", async () => {
    // The new elements are created through the shape's own document, so a
    // detached shape cannot be stamped. The message names the slide rather than
    // reporting a null somewhere inside the DOM.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const shape = await shapeNamed(pkg, "Body");
    Object.defineProperty(shape, "ownerDocument", { value: undefined, configurable: true });
    await expect(writeShapeTags(pkg, SLIDE, shape, OURS)).rejects.toThrow(/belongs to no document/);
  });
});

describe("reading a deck back", () => {
  it("answers every stamped shape on the slide, with the element and the shape it is on", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [PICTURE] }]);
    await writeShapeTags(pkg, SLIDE, await shapeNamed(pkg, "Body"), OURS);
    await writeShapeTags(pkg, SLIDE, await shapeNamed(pkg, "Billede 60"), [
      [TAG_ELEMENT, "bred-bjaelke"],
      [TAG_CATALOGUE, "2026-09-10"],
    ]);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([
      { element: "hvid-kasse-2x1-vertikale", catalogue: "2026-09-10", shapeId: "2" },
      { element: "bred-bjaelke", catalogue: "2026-09-10", shapeId: "60" },
    ]);
  });

  it("says nothing about a shape whose tag part is another add-in's", async () => {
    // A deck touched by think-cell has exactly this on a hidden shape in every
    // slide it has seen. Answering nothing is what keeps "Used in this deck"
    // from reporting somebody else's bookkeeping as an element of ours.
    const pkg = await taggedByAnother(`<p:tag name="VENDOR" val="do not delete"/>`);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
  });

  it("says nothing about a tag part that carries the catalogue but not the element", async () => {
    // Half a stamp is not an element. The element id is what every caller keys
    // on, and a record without one would reach the pane as a nameless row.
    const pkg = await taggedByAnother(`<p:tag name="${TAG_CATALOGUE}" val="2026-09-10"/>`);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
  });

  it("leaves the catalogue out rather than inventing one when the tag carried none", async () => {
    /**
     * The version is what lets a later release tell a shape inserted by an
     * older catalogue from one of its own. An empty string would read as a
     * version that exists, so the absent key has to stay absent.
     */
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    await writeShapeTags(pkg, SLIDE, await shapeNamed(pkg, "Body"), [[TAG_ELEMENT, "hvid-kasse"]]);
    const tagged = await readShapeTags(pkg, SLIDE);
    expect(tagged).toEqual([{ element: "hvid-kasse", shapeId: "2" }]);
    expect("catalogue" in (tagged[0] ?? {})).toBe(false);
  });

  it("says nothing about a slide the deck does not have", async () => {
    // The pane asks per slide, off a list it read earlier; a deck can lose a
    // slide between the two.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    expect(await readShapeTags(pkg, "ppt/slides/slide7.xml")).toEqual([]);
  });

  it("says nothing about a slide with no shape tree to read", async () => {
    // Neither half of `<p:cSld><p:spTree>` can be assumed of a part another
    // tool wrote, and a thrown TypeError here would take down the whole pane.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    pkg.setText(SLIDE, `<?xml version="1.0"?><p:sld xmlns:p="${P_NS}"/>`);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
    pkg.setText(SLIDE, `<?xml version="1.0"?><p:sld xmlns:p="${P_NS}"><p:cSld/></p:sld>`);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
  });

  it("skips a reference whose part has gone, where the writer replaces it", async () => {
    // The pair the module exists to keep in step: the same markup that the
    // writer must survive is markup the reader must not throw on.
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [REFERENCE_TO_NOTHING] }]);
    const shape = await shapeNamed(pkg, "Uden relation");
    const rId = await pkg.addRel(SLIDE, REL_TYPE.tags, "../tags/tag77.xml");
    referenceOn(shape)?.setAttributeNS(R_NS, "r:id", rId);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
    await expect(writeShapeTags(pkg, SLIDE, shape, OURS)).resolves.toBeUndefined();
  });

  it("skips a shape with no tag reference at all, which is most of a real slide", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], noBody: true, shapes: [GROUP, PICTURE, FRAME, CONNECTOR] }]);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([]);
  });

  it("reads a reference whose `r:id` carries no namespace", async () => {
    /**
     * `getAttributeNS` is paired with a plain `getAttribute` fallback for the
     * reason `xml.ts` already gives: a document that came out of a host may not
     * carry the namespace where a reader expects it, and the reference is then
     * invisible — the pane reports an empty deck full of stamped shapes.
     */
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const shape = await shapeNamed(pkg, "Body");
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    const ref = referenceOn(shape) as Element;
    const rId = ref.getAttributeNS(R_NS, "id") ?? "";
    ref.removeAttributeNS(R_NS, "id");
    ref.setAttribute("r:id", rId);
    expect(ref.getAttributeNS(R_NS, "id"), "the attribute is still namespaced, so this proves nothing").toBeNull();
    expect((await readShapeTags(pkg, SLIDE)).map((t) => t.element)).toEqual(["hvid-kasse-2x1-vertikale"]);
  });

  it("reads and stamps a shape whose markup has been pretty-printed", async () => {
    /**
     * Whitespace is text, not an element, and it sits both BETWEEN the shapes
     * and inside each one — so every walk that looks for a named child has to
     * step over it. A deck that has been through a formatter, an XML editor or
     * another tool's writer is full of it, and PowerPoint opens all of them.
     */
    const pkg = await deck([{ paragraphs: [["a"]], shapes: [PRETTY_PICTURE] }]);
    const text = (node: Element): boolean => Array.from(node.childNodes).some((n) => n.nodeType !== 1);
    expect(text(await shapeTree(pkg)), "the whitespace never reached the tree").toBe(true);
    const shape = await shapeNamed(pkg, "Billede 61");
    expect(text(shape), "the whitespace inside the shape never reached the tree").toBe(true);
    await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect((await readShapeTags(pkg, SLIDE)).map((t) => t.shapeId)).toEqual(["61"]);
  });

  it("answers an empty shape id for a shape that states none, either way round", async () => {
    // The id is how a caller finds the shape again in the same markup, and a
    // shape without one is markup no PowerPoint wrote — a `<p:cNvPr>` with no
    // `id`, or no `<p:cNvPr>` at all. Degrading keeps the ELEMENT readable,
    // which is what the pane's "used in this deck" count is made of.
    const pkg = await deck([{ paragraphs: [["a"]], noBody: true, shapes: [WITHOUT_SHAPE_ID, WITHOUT_SHAPE_PROPS] }]);
    const shapes = (await topLevel(pkg)).filter((s) => s.localName === "sp");
    expect(shapes, "the fixture changed shape").toHaveLength(2);
    for (const shape of shapes) await writeShapeTags(pkg, SLIDE, shape, OURS);
    expect((await readShapeTags(pkg, SLIDE)).map((t) => t.shapeId)).toEqual(["", ""]);
  });

  it("ignores a nameless tag and reads a valueless one as empty", async () => {
    // A hand-written or generated tag part, which the reader has to survive:
    // `getAttribute` answers null for both, and null is not a key or a value.
    const pkg = await taggedByAnother(`<p:tag val="orphan"/><p:tag name="${TAG_ELEMENT}"/>`);
    expect(await readShapeTags(pkg, SLIDE)).toEqual([{ element: "", shapeId: "2" }]);
  });

  it("survives the deck being written out and opened again", async () => {
    // Everything above reads a package still holding parsed documents. What
    // PowerPoint gets is bytes, and a stamp that only exists in memory is one
    // the pane can never read back.
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    await writeShapeTags(pkg, SLIDE, await shapeNamed(pkg, "Body"), OURS);
    const again = await Pkg.open(await pkg.toBytes());
    expect(await readShapeTags(again, SLIDE)).toEqual([
      { element: "hvid-kasse-2x1-vertikale", catalogue: "2026-09-10", shapeId: "2" },
    ]);
  });
});
