/**
 * PowerPoint tags, written into the file instead of asked of the API.
 *
 * Ported from SSF-Merge (`src/core/pptx/tags.ts`) on 2026-09-10 and moved down
 * one level: that engine tags a SLIDE, because a merged slide is the thing it
 * owns, and this one tags a SHAPE, because what this add-in puts into a deck is
 * an element and an element is one shape or one group. Everything below the
 * writer — the part shape, the escaping, the numbering, the merge of an
 * existing part — is that file's, comments and all, and the incidents they
 * narrate happened there.
 *
 * The decision they rest on is the same one, and it is the most load-bearing in
 * either engine. On PowerPoint for the web a slide the run has just added does
 * not round-trip through `slides.getItem(id)`, and tag writes through a shape
 * proxy are refused outright: a sibling add-in logged 46
 * `InvalidParam passed to GetItem(id)` failures in one run and needed a whole
 * recovery pass to claw some of them back. A tag written into the package
 * before the insert cannot be refused, because nothing is asked.
 *
 * The shape is `ppt/tags/tagN.xml` holding `<p:tagLst>`, related from the owner
 * part, and referenced from the owner's `<p:custDataLst><p:tags r:id="…"/>`. On
 * a shape that list lives inside `<p:nvPr>`, which every shape kind has:
 * `p:sp`, `p:grpSp`, `p:pic` and `p:graphicFrame` each carry one.
 *
 * **A tag does not survive cut and paste on the web.** `docs/MANUAL.md` says so
 * rather than the code trying to detect it.
 */
import { Pkg } from "./pkg.js";
import { REL_TYPE } from "./parts.js";
import { P_NS, R_NS, child, elements, parseXml, xmlSafe } from "./xml.js";

const TAGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const TAG_LST_NS = `xmlns:p="${P_NS}"`;

/**
 * Tag keys are stored uppercase by PowerPoint, and some APIs require it. Ours
 * are written that way.
 *
 * Two of them, which is what `docs/DESIGN.md` section 6 asks for: which element
 * this shape is, and which catalogue it came out of. The version is what lets a
 * later release tell a shape inserted by an older catalogue from one of its own
 * without guessing from the markup.
 */
export const TAG_ELEMENT = "SSF_SLIDE_ELEMENT";
export const TAG_CATALOGUE = "SSF_SLIDE_ELEMENTS_CATALOGUE";

/**
 * Escape a value for an XML attribute — including the whitespace.
 *
 * The five markup characters are the obvious half. The other half is that an
 * XML parser NORMALISES an attribute value: a literal newline, carriage return
 * or tab inside one is read back as a SPACE. Writing them literally therefore
 * loses them, and the loss happens on the first insert and looks stable
 * afterwards, which is the shape that never gets reported.
 *
 * It cannot touch our own tags — an element id and a catalogue version have no
 * whitespace in them. It reaches a FOREIGN tag, which `mergeTagPart` carries
 * through untouched: a shape from a deck another add-in has touched can already
 * carry a tag part, and an add-in keeping anything formatted in a tag got it
 * back on one line.
 *
 * **A third half: the characters XML cannot carry AT ALL.** Escaping was the
 * whole of this and it is not enough — `&#11;` is exactly as ill-formed as the
 * byte, so a C0 control, a lone surrogate or U+FFFE in a foreign tag produced a
 * part PowerPoint refuses, reported as a damaged file with nothing naming the
 * cause. `xmlSafe` is that rule.
 *
 * It runs FIRST, because escaping a character that may not be written is
 * writing it.
 */
function xmlAttr(s: string): string {
  return xmlSafe(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#13;")
    .replace(/\t/g, "&#9;");
}

export function tagPartXml(entries: [string, string][]): string {
  const tags = entries.map(([name, val]) => `<p:tag name="${xmlAttr(name)}" val="${xmlAttr(val)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:tagLst ${TAG_LST_NS}>${tags}</p:tagLst>`;
}

/**
 * An unused `ppt/tags/tagN.xml`: the highest this package has held, plus one.
 *
 * Blind use of `tag1.xml` is the trap here: a deck that already carries one
 * would have it overwritten, and every shape pointing at it would silently lose
 * its tags. The number is taken from the package, never assumed — and the
 * library's own elements make that certain rather than theoretical: six of the
 * 117 harvested 16:9 elements arrive carrying tag relationships of their own,
 * 78 of them between the six, and one element carries 41.
 */
export function nextTagNumber(pkg: Pkg): number {
  // Through `Pkg`'s own counter, which reads the package once and then keeps
  // itself current.
  //
  // NOT "the first free number", and an earlier spelling of this comment in the
  // sibling said it was: a package holding tag2 and tag5 and no tag1 answers 6
  // here, where a walk from 1 answered 1. Both are names the package does not
  // hold, which is all this has to guarantee.
  return pkg.nextNumber("ppt/tags/tag");
}

/**
 * Append or replace entries in an existing tag part, keeping the ones we do not
 * own.
 *
 * Parsed, not pattern-matched. The regex this replaced in the sibling read
 * attribute VALUES as raw source, so `val="Ben &amp; Jerry"` came back with the
 * entity intact and was escaped a second time on write: one merge turned it
 * into `Ben &amp;amp; Jerry`, two into `Ben &amp;amp;amp; Jerry`, and a reader
 * saw the literal `&amp;` on screen. It also insisted on one attribute order
 * and a self-closing tag, so PowerPoint's own perfectly legal spellings — `val`
 * before `name`, single quotes, a separate closing tag — matched nothing and
 * the foreign tag was DROPPED.
 *
 * The parser decodes; `tagPartXml` encodes exactly once. That round trip is
 * what makes repeated inserts stable.
 */
export function mergeTagPart(xml: string, entries: [string, string][]): string {
  const kept: [string, string][] = [];
  const incoming = new Set(entries.map(([k]) => k));
  for (const tag of elements(parseXml(xml), P_NS, "tag")) {
    const name = tag.getAttribute("name");
    if (name && !incoming.has(name)) kept.push([name, tag.getAttribute("val") ?? ""]);
  }
  return tagPartXml([...kept, ...entries]);
}

/**
 * The `<p:nvPr>` of a top-level shape, whatever kind of shape it is.
 *
 * Each kind wraps it in a differently named container — `p:nvSpPr` for a shape,
 * `p:nvGrpSpPr` for a group, `p:nvPicPr` for a picture, `p:nvGraphicFramePr`
 * for a table, a chart or SmartArt, `p:nvCxnSpPr` for a connector — and
 * `p:nvPr` is the last child of all five.
 *
 * Searched one level down rather than by descendant, which is the same scoping
 * the sibling's slide-level writer had to learn: `element` walks descendants, so
 * a search that was not scoped found a `<p:nvPr>` belonging to a shape INSIDE a
 * group and put the element's identity there instead of on the group.
 */
function nvPrOf(shape: Element): Element | undefined {
  for (const container of Array.from(shape.childNodes)) {
    if (container.nodeType !== 1) continue;
    const found = child(container as Element, P_NS, "nvPr");
    if (found) return found;
  }
  return undefined;
}

/** A top-level shape's `<p:cNvPr id>`, found the same way and for the same reason. */
function idOf(shape: Element): string | undefined {
  for (const container of Array.from(shape.childNodes)) {
    if (container.nodeType !== 1) continue;
    const cNvPr = child(container as Element, P_NS, "cNvPr");
    if (cNvPr) return cNvPr.getAttribute("id") ?? undefined;
  }
  return undefined;
}

/**
 * Attach tags to one top-level shape, merging with whatever is already there.
 *
 * `CT_CustomerDataList` allows at most one `<p:tags>` child, so a shape that
 * already has a tag part must have its entries appended rather than a second
 * part added. A shape harvested out of the library can: the owner's decks were
 * built in PowerPoint, and six of the 117 elements come with tags already on
 * them — a minority, but one of them carries 41 relationships, so the append
 * path is exercised rather than theoretical.
 *
 * The shape must already be IN `slidePath` — the relationship is added to that
 * slide's rels part and the reference is written into the shape's own markup,
 * so a caller that tagged a detached element and appended it afterwards would
 * write a relationship into one part and a reference into another.
 */
export async function writeShapeTags(
  pkg: Pkg,
  slidePath: string,
  shape: Element,
  entries: [string, string][],
): Promise<void> {
  const nvPr = nvPrOf(shape);
  if (!nvPr) throw new Error(`ssf-slide-elements: a top-level shape in ${slidePath} has no <p:nvPr> to tag`);

  const custData = child(nvPr, P_NS, "custDataLst");
  const existing = custData ? child(custData, P_NS, "tags") : undefined;
  if (existing) {
    const rId = existing.getAttributeNS(R_NS, "id") ?? existing.getAttribute("r:id");
    const target = rId ? await pkg.relTarget(slidePath, rId) : undefined;
    // `pkg.has`, not just a resolved target. `relTarget` answers what the
    // relationship POINTS AT, and a relationship can point at a part that is
    // not in the package — a deck another tool wrote, or one PowerPoint
    // repaired by dropping the part and leaving the reference. `pkg.text`
    // throws by name for a missing part, so a shape like that would kill the
    // whole insert while `readShapeTags` below guards with exactly this test
    // and returns nothing. A reader that degrades and a writer that throws on
    // the same markup is the pair worth never shipping.
    if (target && pkg.has(target)) {
      pkg.setText(target, mergeTagPart(await pkg.text(target), entries));
      return;
    }
    // The reference is there and leads nowhere. It has to GO before a fresh one
    // is written, because `CT_CustomerDataList` allows at most one `<p:tags>`
    // child and the fall-through below appends into this same
    // `<p:custDataLst>` — so leaving it produces two, which is schema-invalid,
    // and a reader takes the FIRST. The insert's own tag would then be
    // invisible to every reader of it: the pane could not report the elements
    // it had put in the deck, on a deck that opens perfectly well.
    //
    // The dangling RELATIONSHIP is deliberately left alone. It was in the deck
    // before this ran, and removing relationships is the operation that has
    // twice produced damage in the sibling — an id freed by a delete is handed
    // to the next thing that asks for one.
    existing.parentNode?.removeChild(existing);
  }

  const n = nextTagNumber(pkg);
  const part = `ppt/tags/tag${n}.xml`;
  pkg.setText(part, tagPartXml(entries));
  await pkg.addContentTypeOverride(`/${part}`, TAGS_CONTENT_TYPE);
  const rId = await pkg.addRel(slidePath, REL_TYPE.tags, `../tags/tag${n}.xml`);

  const doc = shape.ownerDocument;
  if (!doc) throw new Error(`ssf-slide-elements: the shape to tag in ${slidePath} belongs to no document`);
  const tags = doc.createElementNS(P_NS, "p:tags");
  tags.setAttributeNS(R_NS, "r:id", rId);
  if (custData) {
    // `CT_CustomerDataList` allows one `<p:tags>`, and by here there is none —
    // either the list never had one, or the one that led nowhere was taken out
    // above. A list holding only `<p:custData>` children is legal and common.
    custData.appendChild(tags);
    return;
  }
  const custDataLst = doc.createElementNS(P_NS, "p:custDataLst");
  custDataLst.appendChild(tags);
  // `CT_ApplicationNonVisualDrawingProps` orders its children
  // `ph?, (audioCd|wavAudioFile|audioFile|videoFile|quickTimeFile)?,
  // custDataLst?, extLst?`, so the list goes before an `<p:extLst>` and after
  // everything else. Appended blindly it would land after the extension list,
  // which is a part PowerPoint reports as damaged without naming which.
  const extLst = child(nvPr, P_NS, "extLst");
  if (extLst) nvPr.insertBefore(custDataLst, extLst);
  else nvPr.appendChild(custDataLst);
}

/** One tagged top-level shape: which element it is, and where it sits. */
export interface TaggedShape {
  /** The value of `SSF_SLIDE_ELEMENT`: the catalogue id of the element. */
  element: string;
  /** The value of `SSF_SLIDE_ELEMENTS_CATALOGUE`, absent on a shape whose tag carried none. */
  catalogue?: string;
  /** The shape's `<p:cNvPr id>`, so a caller can find it again in the same markup. */
  shapeId: string;
}

/**
 * Every top-level shape on a slide that this add-in put there.
 *
 * What "Used in this deck" and "Remove from N slides" are read from
 * (`docs/DESIGN.md` sections 4 and 6). Read from the FILE rather than through
 * the API for the reason at the top of this file, and read per top-level shape
 * rather than per slide because one slide can carry several elements.
 *
 * A shape carrying a tag part that is not ours answers nothing rather than
 * throwing: a deck touched by another add-in has exactly that on every slide.
 */
export async function readShapeTags(pkg: Pkg, slidePath: string): Promise<TaggedShape[]> {
  const out: TaggedShape[] = [];
  if (!pkg.has(slidePath)) return out;
  const doc = await pkg.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) return out;

  for (const node of Array.from(spTree.childNodes)) {
    if (node.nodeType !== 1) continue;
    const shape = node as Element;
    const nvPr = nvPrOf(shape);
    if (!nvPr) continue;
    const custData = child(nvPr, P_NS, "custDataLst");
    const ref = custData ? child(custData, P_NS, "tags") : undefined;
    const rId = ref?.getAttributeNS(R_NS, "id") ?? ref?.getAttribute("r:id");
    if (!rId) continue;
    const target = await pkg.relTarget(slidePath, rId);
    if (!target || !pkg.has(target)) continue;
    const values = new Map<string, string>();
    for (const tag of elements(await pkg.doc(target), P_NS, "tag")) {
      const name = tag.getAttribute("name");
      if (name) values.set(name, tag.getAttribute("val") ?? "");
    }
    const element = values.get(TAG_ELEMENT);
    if (element === undefined) continue;
    const catalogue = values.get(TAG_CATALOGUE);
    out.push({
      element,
      ...(catalogue === undefined ? {} : { catalogue }),
      shapeId: idOf(shape) ?? "",
    });
  }
  return out;
}

/** One library element the deck already carries, and where. */
export interface DeckUse {
  /** The catalogue id out of the shape's `SSF_SLIDE_ELEMENT` tag. */
  element: string;
  /** The slides it is on, 1-based, in order and without repeats. */
  slides: number[];
  /** How many tagged top-level shapes carry it, which can be more than one per slide. */
  shapes: number;
  /** Every catalogue version stamped on those shapes, oldest spelling first; empty when none was. */
  catalogues: string[];
}

/**
 * Every library element already in a deck, read from the tags the add-in wrote.
 *
 * `docs/DESIGN.md` section 4's "Used in this deck", and section 15 is why it
 * can be built at all: the tags an insert writes were measured surviving
 * `insertSlidesFromBase64` on the web on 2026-09-11 — seven of them across four
 * slides, with their relationships intact — so this rests on a measured host
 * fact rather than a hope.
 *
 * Read from the FILE, never through the shape collection. A deck that has been
 * through another add-in carries ITS tags in the same folder, numbered around
 * ours — think-cell's were sitting there in the measured deck — and a sweep
 * that assumed `ppt/tags/tagN.xml` was ours would report somebody else's
 * bookkeeping as our elements. `readShapeTags` keys on the tag NAME instead,
 * per slide, and this only groups what it answers.
 *
 * An id the current catalogue no longer has still appears: ids are slugs of the
 * element keys and eleven of them changed on 2026-09-11 (section 2), so a deck
 * built before that names elements this library cannot. Naming them is the
 * caller's problem; hiding them here would make the count wrong.
 */
export async function usedInDeck(pkg: Pkg): Promise<DeckUse[]> {
  const found = new Map<string, { slides: Set<number>; shapes: number; catalogues: Set<string> }>();
  const paths = await pkg.slidePaths();
  for (let i = 0; i < paths.length; i++) {
    const path = paths[i] as string;
    for (const tagged of await readShapeTags(pkg, path)) {
      const entry = found.get(tagged.element) ?? {
        slides: new Set<number>(),
        shapes: 0,
        catalogues: new Set<string>(),
      };
      entry.slides.add(i + 1);
      entry.shapes += 1;
      if (tagged.catalogue !== undefined) entry.catalogues.add(tagged.catalogue);
      found.set(tagged.element, entry);
    }
  }
  return (
    [...found.entries()]
      .map(([element, entry]) => ({
        element,
        slides: [...entry.slides].sort((a, b) => a - b),
        shapes: entry.shapes,
        catalogues: [...entry.catalogues].sort(),
      }))
      // By where they first appear, which is the order somebody scrolling the
      // deck would meet them in.
      .sort((a, b) => (a.slides[0] ?? 0) - (b.slides[0] ?? 0) || a.element.localeCompare(b.element))
  );
}
