/**
 * An element into the slide the user is on, in the file.
 *
 * This is the whole mechanism `docs/DESIGN.md` section 6 describes, and it is
 * one shape: **rebuild the destination slide with the element on it, hand
 * PowerPoint the rebuild, and take the original away.** Nothing is drawn
 * through the shape collection — `docs/BACKLOG.md` lists that under "Rejected —
 * do not re-propose", with the sibling rounds that rejected it.
 *
 * The package handed back holds the user's ENTIRE deck with one slide added and
 * every other slide UNLISTED: its part is still in the zip, its relationship
 * still points at it, but no `<p:sldId>` names it. PowerPoint for the web was
 * asked directly whether it accepts that, on 2026-09-10, because the whole cost
 * of the insert turns on it — `docs/host-answers/` holds the sheets and
 * `docs/PROBE.md` question 1 the reading. It does: both prunings landed exactly
 * one slide, so the host reads the slide list and ignores a part nothing lists.
 * The expensive alternative was removing every other slide properly, orphans
 * walked, on a deck that can be fifty megabytes.
 *
 * The second measured fact this rests on is question 1b. An insert of the
 * deck's OWN bytes added no master under either formatting option (1 → 1),
 * while a foreign fixture deck added one (1 → 2). This package is built from
 * the user's own deck, so the master, the layout and the theme the rebuilt
 * slide points at are the ones already in it, and the insert does not grow the
 * deck by a master. Building the slide from the library deck instead would have.
 */
import { contentCount } from "../catalogue/boxes.js";
import type { MarkupRel } from "../catalogue/types.js";
import { cloneSlide } from "../pptx/clone.js";
import { framesOf, slideSize } from "../pptx/layout.js";
import { Pkg } from "../pptx/pkg.js";
import { COMMENT_REL_TYPES, REL_TYPE } from "../pptx/parts.js";
import { TAG_CATALOGUE, TAG_ELEMENT, taggable, writeShapeTags } from "../pptx/tags.js";
import { A_NS, PKG_REL_NS, P_NS, child, children, elements } from "../pptx/xml.js";
import { carry, type PartStore } from "./carry.js";
import { keepOnly } from "./listing.js";
import { pinColoursInXml, pinSchemeColours } from "./colours.js";
import { authored, moveFrom, place, type Box, type Landing, type Rect } from "./landing.js";
import {
  applyMove,
  emptyBodyPlaceholders,
  groupShapes,
  groupable,
  highestShapeId,
  parseFragment,
  renumber,
  repoint,
  slideShapes,
  topLevel,
  unplaceholder,
} from "./shapes.js";

/** What the pane knows about the element it is inserting. */
export interface SpliceElement {
  /** The catalogue id, written into the shape's tag so the deck can be read back. */
  id: string;
  /** The English name, which becomes the shape's name in the selection pane. */
  name: string;
  kind: "slide" | "part";
  box: Box;
  landing: Landing;
  /** A marker wraps the selected shape rather than merely sitting on it. */
  wraps?: boolean;
  markup: { xml: string; rels: MarkupRel[] };
}

export interface SpliceOptions {
  /** `docs/DESIGN.md` section 7: onto this slide, or as a new slide after it. */
  target: "onto" | "new";
  /** Section 7 again: multi-shape elements land as one group, or loose. */
  group: boolean;
  /**
   * Section 7 once more: `deck` leaves every scheme colour alone, so the
   * element takes the destination's theme, and `library` pins each of them to
   * the value the library's theme gave it.
   *
   * `deck` is the default and costs nothing — it is what happens when this pass
   * does not run at all.
   */
  colours: "deck" | "library";
}

export interface SpliceRequest {
  /** The user's deck, as `getFileAsync` handed it over. */
  deck: Uint8Array | string;
  /** Which slide the user is on, counting from zero. */
  slide: number;
  element: SpliceElement;
  options: SpliceOptions;
  /**
   * What the catalogue knows that this element's own entry does not: the
   * version stamped into every inserted shape's tag, the content type of every
   * carried part, and the library theme's colour map the "As in the library"
   * setting pins to.
   */
  catalogue: { version: string; carried: Record<string, string>; theme?: Record<string, string> };
  store: PartStore;
  /** The selected shape's rectangle, when the host could name one. */
  selection?: Rect;
}

/**
 * What the splice did, in numbers the pane can show and the suite can assert.
 *
 * The base64 is the package; everything else is evidence. `docs/DESIGN.md`
 * section 6 makes the deck delta the evidence an insert worked, never the
 * absence of an error, and a report that only carried the bytes would leave the
 * pane with nothing to say afterwards.
 */
export interface SpliceReport {
  base64: string;
  /** How many slides the user's deck holds. The splice does not change it. */
  deckSlides: number;
  /** The rebuilt slide inside the package: the one slide the package lists. */
  slidePath: string;
  /** Where the element landed, in EMU on the destination slide. */
  landed: Rect;
  /** How many top-level shapes were added: one when they were grouped. */
  shapes: number;
  grouped: boolean;
  /** Carried parts copied into the package. */
  parts: number;
  /**
   * Scheme colours pinned to the library's own values, across the element's
   * markup and every XML part it carried. Zero under "This deck's theme", and
   * zero for the five library elements that state no scheme colour at all.
   */
  pinned: number;
  /** Empty content placeholders taken off the rebuilt slide. */
  placeholders: number;
  /**
   * What the slide the user was ON already held, counted by `contentCount`:
   * its own title and the empty placeholders an insert removes do not count.
   *
   * The slide the request NAMED, always — not the rebuilt copy and not the
   * blanked clone "as a new slide" starts from. It is what the footer's "Move
   * to a new slide" is decided on (`docs/DESIGN.md` section 6), and that offer
   * is about the slide the user was looking at.
   */
  held: number;
}

/**
 * The deck with exactly one of its own slides listed, and nothing changed.
 *
 * What Undo hands back to PowerPoint. Taking an insert back that landed "onto
 * this slide" means putting the REPLACED slide back, and the only way to put a
 * slide into a deck is to hand the host a package containing it — so this is
 * the same unlisting trick the splice itself uses, over the bytes the pane read
 * before the insert.
 *
 * Nothing is rebuilt and nothing is cloned: the slide that comes back is the
 * user's own, byte for byte, which is the whole point of an undo.
 */
export async function onlySlide(deck: Uint8Array | string, slide: number): Promise<{ base64: string; path: string }> {
  const pkg = await Pkg.open(deck);
  const paths = await pkg.slidePaths();
  const path = paths[slide];
  if (path === undefined) {
    throw new Error(`ssf-slide-elements: slide ${slide + 1} is not in this deck, which has ${paths.length} slide(s)`);
  }
  await keepOnly(pkg, path);
  return { base64: await pkg.toBase64(), path };
}

/**
 * The shape tree of a slide part.
 *
 * The missing-`<p:spTree>` raise is reachable and `splice-malformed.test.ts`
 * holds it. The `cSld ? … : undefined` arm beside it is NOT: measured
 * 2026-09-12, a slide with no `<p:cSld>` is refused by `cloneSlide` first, with
 * a sentence naming that element instead. It stays because it is what lets this
 * return an `Element` rather than an `Element | undefined`, which every caller
 * would then have to re-check.
 */
async function spTreeOf(pkg: Pkg, slidePath: string): Promise<Element> {
  const doc = await pkg.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) throw new Error(`ssf-slide-elements: ${slidePath} has no <p:spTree>`);
  return spTree;
}

/**
 * Empty a cloned slide, so "as a new slide" starts from the layout rather than
 * from what the user happened to be looking at.
 *
 * The clone is the basis rather than a slide built from nothing, because that
 * is what guarantees the new slide sits on the same layout and master as the
 * one it follows — and a slide whose layout is not in the deck is a slide
 * PowerPoint has to invent a design for.
 *
 * Placeholders stay and are emptied; everything else goes. So do the notes page
 * and the COMMENTS, and both for the same reason: a new slide carrying the
 * previous slide's speaker notes or somebody's review thread is a surprise
 * nobody asked for. The parts they point at are left in the package as orphans,
 * which `scripts/package-integrity.mjs` treats as weight rather than damage and
 * which the insert never reaches, because nothing lists them.
 *
 * The comments half was measured rather than reasoned. On PowerPoint for the
 * web on 2026-09-10 a slide carrying one comment was replaced twice and then
 * used as the basis for a new slide, and the deck came back with the comment on
 * BOTH — `ppt/comments/modernComment_104_*` and `_105_*`, one each. A modern
 * comment on the web is anchored from the slide's own extension list, so
 * `cloneSlide`'s drop-what-nothing-names pass keeps it, which is right for
 * "onto this slide" — the user's comment survives their slide being rebuilt —
 * and wrong for a slide that is meant to be new.
 */
async function blank(pkg: Pkg, slidePath: string): Promise<void> {
  const spTree = await spTreeOf(pkg, slidePath);
  for (const shape of slideShapes(spTree)) {
    const isPlaceholder = placeholderIn(shape);
    if (!isPlaceholder) {
      shape.parentNode?.removeChild(shape);
      continue;
    }
    const txBody = child(shape, P_NS, "txBody");
    if (!txBody) continue;
    // Every paragraph but the first goes, and the first is emptied of runs. A
    // `<p:txBody>` with no `<a:p>` at all is schema-invalid; PowerPoint reports
    // the file as damaged and names nothing.
    const paragraphs = children(txBody, A_NS, "p");
    for (const p of paragraphs.slice(1)) p.parentNode?.removeChild(p);
    const first = paragraphs[0];
    if (!first) continue;
    for (const node of Array.from(first.childNodes)) {
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      // `<a:pPr>` and `<a:endParaRPr>` are the paragraph's own formatting and
      // are kept, so the empty placeholder still looks like itself. Runs go.
      if (el.namespaceURI === A_NS && (el.localName === "pPr" || el.localName === "endParaRPr")) continue;
      first.removeChild(el);
    }
  }

  const relsPath = Pkg.relsPathFor(slidePath);
  if (!pkg.has(relsPath)) return;
  const rels = await pkg.doc(relsPath);
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    const type = rel.getAttribute("Type") ?? "";
    if (type === REL_TYPE.notesSlide || COMMENT_REL_TYPES.includes(type)) rel.parentNode?.removeChild(rel);
  }
}

/** Whether a top-level shape claims a placeholder on its slide. */
function placeholderIn(shape: Element): boolean {
  for (const node of Array.from(shape.childNodes)) {
    if (node.nodeType !== 1) continue;
    const nvPr = child(node as Element, P_NS, "nvPr");
    if (nvPr && child(nvPr, P_NS, "ph")) return true;
  }
  return false;
}

/**
 * Splice an element into a deck and answer the package to hand PowerPoint.
 *
 * The order of the steps is not arrangeable. Ids are renumbered against the
 * destination slide BEFORE anything is appended to it, so the numbering sees
 * the slide as it was; the parts are carried and the references repointed
 * BEFORE the shapes are adopted, so a failure to find a part leaves the user's
 * deck untouched rather than half-written; and the tags are written LAST,
 * because a tag is a relationship on the slide and the shape has to be in the
 * slide to carry the reference.
 */
export async function splice(request: SpliceRequest): Promise<SpliceReport> {
  const pkg = await Pkg.open(request.deck);
  const paths = await pkg.slidePaths();
  const deckSlides = paths.length;
  if (deckSlides === 0) throw new Error("ssf-slide-elements: this deck has no slides to insert into");
  if (!Number.isInteger(request.slide) || request.slide < 0 || request.slide >= deckSlides) {
    throw new Error(
      `ssf-slide-elements: slide ${request.slide + 1} is not in this deck, which has ${deckSlides} slide(s)`,
    );
  }
  const source = paths[request.slide] as string;

  const size = await slideSize(pkg);
  const frames = await framesOf(pkg, source);
  // Read from `source` BEFORE anything is cloned or blanked, so it is the
  // user's own slide that is counted whichever target this call is for.
  const held = contentCount(await pkg.doc(source), size.width, size.height);

  const rebuilt = await cloneSlide(pkg, source);
  if (request.options.target === "new") await blank(pkg, rebuilt);

  const spTree = await spTreeOf(pkg, rebuilt);
  const fragment = parseFragment(request.element.markup.xml);
  const tops = topLevel(fragment);
  if (tops.length === 0) throw new Error(`ssf-slide-elements: the element "${request.element.id}" has no shapes`);

  // A library shape can claim to BE the slide's body. On a destination that
  // already has one, two shapes claim the same index and PowerPoint resolves it
  // by inheriting geometry from the layout — the element jumps to wherever the
  // layout's placeholder is and the landing computed below is thrown away.
  unplaceholder(tops);

  // `docs/DESIGN.md` section 7's colour switch, and it is a rewrite of the
  // element's OWN markup only — the destination slide is never touched, so a
  // deck whose theme the user likes keeps it everywhere except on what they
  // just inserted. Done here, before the shapes are moved or grouped, so
  // everything after this works on markup that already says what it means.
  const theme = request.catalogue.theme ?? {};
  let pinned = 0;
  if (request.options.colours === "library") pinned += pinSchemeColours(fragment, theme).pinned;

  let nextId = highestShapeId(spTree) + 1;
  nextId = renumber(fragment, nextId);

  const carried = await carry({
    pkg,
    owner: rebuilt,
    rels: request.element.markup.rels,
    types: request.catalogue.carried,
    store: request.store,
    ...(request.options.colours === "library"
      ? {
          transform: (_path: string, xml: string) => {
            const done = pinColoursInXml(xml, theme);
            pinned += done.result.pinned;
            return done.xml;
          },
        }
      : {}),
  });
  repoint(fragment, carried.ids);

  // The BOX is the reference frame, not the union of the shapes' own
  // rectangles, and the two are not the same thing: the catalogue's box is a
  // rotated shape's rotated extent (`docs/DESIGN.md` section 3), which is the
  // ink the user sees. Landing the ink where the rule says is the promise;
  // landing the unrotated frame there would put a 29° stamp off the edge.
  const from = authored(request.element.box, size);
  const landed = place({
    box: request.element.box,
    landing: request.element.landing,
    slide: size,
    frames,
    ...(request.selection ? { selection: request.selection } : {}),
    ...(request.element.wraps ? { wraps: true } : {}),
  });
  applyMove(tops, from, moveFrom(from, landed));

  // The group takes the next id after the shapes it wraps, which `renumber`
  // has just told us. Nothing needs an id after this one — a second element is
  // a second call to `splice`, which reads the slide again — so the counter is
  // spent here rather than advanced for a caller that does not exist.
  const wanted = request.options.group && groupable(tops);
  if (wanted) groupShapes(fragment, tops, request.element.name, nextId);

  // Not reachable, and kept for the type rather than for the case: `spTree` was
  // found by walking a document `pkg.doc` parsed, so it has one. `ownerDocument`
  // is nullable on the DOM's own `Node`, and a non-null assertion here would be
  // a claim with no sentence attached — this at least says what went wrong if
  // the impossible happens. Measured uncovered 2026-09-12; nothing can reach it.
  const doc = spTree.ownerDocument;
  if (!doc) throw new Error(`ssf-slide-elements: ${rebuilt} belongs to no document`);
  const added: Element[] = [];
  for (const shape of topLevel(fragment)) {
    const imported = doc.importNode(shape, true);
    spTree.appendChild(imported);
    added.push(imported);
  }

  // `docs/DESIGN.md` section 6: a whole-slide element removes the "Click to add
  // text" ghosts it lands over. The TITLE stays, and a placeholder the user has
  // typed into is content rather than a ghost — `emptyBodyPlaceholders` is
  // where both of those live.
  let placeholders = 0;
  if (request.element.kind === "slide") {
    for (const ghost of emptyBodyPlaceholders(spTree)) {
      ghost.parentNode?.removeChild(ghost);
      placeholders += 1;
    }
  }

  const stamp: [string, string][] = [
    [TAG_ELEMENT, request.element.id],
    [TAG_CATALOGUE, request.catalogue.version],
  ];
  for (const shape of added) {
    await writeShapeTags(pkg, rebuilt, shape, stamp);
  }

  // The shapes INSIDE any group that lands are stamped too.
  //
  // Ungrouping is one gesture, and it destroys the group and the tag on it
  // together: a five-shape element ungrouped went from "used once in this deck"
  // to not in the deck at all, with the shapes still on the slide. Stamping the
  // insides means one ungroup leaves five tagged shapes — which is the truth,
  // and exactly what the loose setting writes anyway. Measured at about 250
  // bytes a shape.
  //
  // **Any group, not only one this code made.** The first version asked
  // `wanted`, which is true only when the splice WRAPPED the shapes — and 23
  // elements in the shipped libraries are already a single `<p:grpSp>` the
  // owner drew, so there is nothing to wrap and `wanted` is false. Those were
  // exactly the elements the fix was meant to cover: ungrouping `confidential`
  // left two untagged shapes on the slide and the removal then refused by name,
  // "slide 1 carries no confidential this add-in inserted". Found by sweeping
  // every element in both libraries rather than the one the fix was written
  // against.
  //
  // The reader stops AT a tagged shape, so while the group is still a group
  // this changes nothing it answers: one element, one use.
  //
  // A shape with no `<p:nvPr>` is skipped rather than refused. `<mc:Alter`
  // `nateContent`, which is how a modern chart sits on a slide, has none — and
  // failing an insert that works today would be a worse trade than losing a tag
  // on an ungroup that may never happen.
  for (const shape of added) {
    if (shape.namespaceURI !== P_NS || shape.localName !== "grpSp") continue;
    for (const inner of slideShapes(shape)) {
      if (!taggable(inner)) continue;
      await writeShapeTags(pkg, rebuilt, inner, stamp);
    }
  }

  await keepOnly(pkg, rebuilt);

  return {
    base64: await pkg.toBase64(),
    deckSlides,
    slidePath: rebuilt,
    landed,
    shapes: added.length,
    grouped: wanted,
    parts: carried.parts.size,
    pinned,
    placeholders,
    held,
  };
}
