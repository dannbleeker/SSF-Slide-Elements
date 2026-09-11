/**
 * An element taken back off a slide it was inserted onto.
 *
 * `docs/DESIGN.md` sections 4 and 6: a stamp already in the deck can be removed
 * from every slide it is on with one click, found by the tag the insert wrote.
 * This is the engine half — one slide at a time, as a package the host can be
 * handed.
 *
 * **The same mechanism as the insert, in the other direction.** A slide cannot
 * be edited in place through Office.js, so removing a shape means rebuilding
 * the slide without it, handing PowerPoint a package containing the rebuild,
 * and taking the original away — exactly the sequence `splice` uses and the
 * only one this family has measured. Nothing here draws or deletes through the
 * shape collection.
 *
 * **What it removes is what the add-in put there, and nothing else.** The
 * shapes are found by the `SSF_SLIDE_ELEMENT` tag, per shape, through
 * `readShapeTags` — which keys on the tag NAME rather than on the part's path,
 * because a deck that has been through another add-in carries ITS tags in the
 * same folder, numbered around ours. A user's own shape has no tag of ours and
 * cannot be reached from here at all.
 *
 * **Wherever the user has since put it.** Grouping our stamp with a shape of
 * their own is one gesture, and it puts the tag one level down. Sweeping only
 * the top level left that shape on the slide while reporting the element
 * removed, so the sweep goes into the user's groups too — taking out the tagged
 * shape, leaving their own beside it, and taking the group as well only when
 * the removal is what emptied it.
 *
 * **Two things it deliberately leaves behind.** The tag parts the removed
 * shapes pointed at stay in the package as orphans, and so do any pictures they
 * carried: `scripts/package-integrity.mjs` treats an orphan as weight rather
 * than damage, and a sweep that deleted them would be a sweep deciding what
 * else in the user's deck still needs a part — which is the one class of
 * mistake that produces a file PowerPoint calls damaged.
 *
 * **What that weighs, measured rather than assumed** (2026-09-11, 50
 * insert-then-remove cycles on `template/validators.pptx`, 37.9 KB to start):
 *
 * | element                        | per cycle | after 50 |
 * | ------------------------------ | --------- | -------- |
 * | `hvid-kasse-1-stor`, 5 shapes  |  5.5 KB   |  315 KB  |
 * | `confidential`, 1 shape        |  4.1 KB   |  242 KB  |
 * | `markeringer-1`, carries an EMF|  3.5 KB   |  211 KB  |
 *
 * **And the host reclaims all of it.** That was written here as an open
 * question and then measured the same day (section 15): a deck after three
 * insert-then-remove cycles, opened through COM with no window and saved as
 * `.pptx`, went from 73 parts to 39, from 9 slide parts to 1, and from **18
 * orphaned tag parts to none** — 55,072 bytes back to 37,556, which is the deck
 * it started as. PowerPoint drops what nothing references when it writes the
 * file.
 *
 * So the number above is what chaining THIS engine's own packages costs
 * between saves, and the decision to leave orphans rather than walk them costs
 * a user nothing. Leaving them remains right for the original reason — a sweep
 * that deleted them would be deciding what else in the user's deck still needs
 * a part — and now also for a measured one: the host does it better, and it is
 * the host's file.
 */
import { cloneSlide } from "../pptx/clone.js";
import { Pkg } from "../pptx/pkg.js";
import { readShapeTags } from "../pptx/tags.js";
import { P_NS, child } from "../pptx/xml.js";
import { keepOnly } from "./listing.js";
import { slideShapes } from "./shapes.js";

export interface RemoveRequest {
  /** The user's deck, as `getFileAsync` handed it over. */
  deck: Uint8Array | string;
  /** Which slide to take it off, counting from zero. */
  slide: number;
  /** The catalogue id in the shapes' tags. */
  element: string;
}

export interface RemoveReport {
  /** The package to hand PowerPoint: the user's deck with one rebuilt slide listed. */
  base64: string;
  /** How many slides the user's deck holds. The removal does not change it. */
  deckSlides: number;
  /** The rebuilt slide inside the package. */
  slidePath: string;
  /** How many tagged shapes came off, at any depth. */
  removed: number;
  /** What is left of that element on that slide: zero unless something was refused. */
  left: number;
}

/**
 * Which slides carry an element, out of a deck already open.
 *
 * The pane asks this before offering "Remove from N slides", and again is not
 * asked to trust it: the removal itself reads the slide it is about to rebuild.
 */
export async function slidesHolding(pkg: Pkg, element: string): Promise<number[]> {
  const out: number[] = [];
  const paths = await pkg.slidePaths();
  for (let i = 0; i < paths.length; i++) {
    const tagged = await readShapeTags(pkg, paths[i] as string);
    if (tagged.some((t) => t.element === element)) out.push(i);
  }
  return out;
}

/** The shape tree of a slide part. */
async function spTreeOf(pkg: Pkg, slidePath: string): Promise<Element> {
  const doc = await pkg.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) throw new Error(`ssf-slide-elements: ${slidePath} has no <p:spTree>`);
  return spTree;
}

/** A shape's `<p:cNvPr id>`, by the path the schema names for its kind. */
function idOf(shape: Element): string | undefined {
  for (const node of Array.from(shape.childNodes)) {
    if (node.nodeType !== 1) continue;
    const cNvPr = child(node as Element, P_NS, "cNvPr");
    const id = cNvPr?.getAttribute("id");
    if (id) return id;
  }
  return undefined;
}

/**
 * Rebuild one slide without an element, and answer the package to hand over.
 *
 * The rebuilt slide is a clone of the user's own, so it keeps their layout,
 * their master and everything else on it; only the tagged shapes go. The
 * package lists that one slide and nothing else, which is the same shape of
 * package the insert hands over and the same one probe question 1 measured the
 * host accepting.
 */
export async function removeElement(request: RemoveRequest): Promise<RemoveReport> {
  const pkg = await Pkg.open(request.deck);
  const paths = await pkg.slidePaths();
  const deckSlides = paths.length;
  if (!Number.isInteger(request.slide) || request.slide < 0 || request.slide >= deckSlides) {
    throw new Error(
      `ssf-slide-elements: slide ${request.slide + 1} is not in this deck, which has ${deckSlides} slide(s)`,
    );
  }
  const source = paths[request.slide] as string;

  const wanted = new Set(
    (await readShapeTags(pkg, source)).filter((t) => t.element === request.element).map((t) => t.shapeId),
  );
  if (wanted.size === 0) {
    throw new Error(
      `ssf-slide-elements: slide ${request.slide + 1} carries no "${request.element}" this add-in inserted`,
    );
  }

  const rebuilt = await cloneSlide(pkg, source);
  const spTree = await spTreeOf(pkg, rebuilt);
  let removed = 0;
  /**
   * Take the tagged shapes out, wherever the user has since put them.
   *
   * A shape that has been grouped with one of the user's own is one level down,
   * and a removal that only swept the top level left it on the slide while
   * reporting the element gone. What comes out is still only what carries our
   * tag: the user's own shape in that group is not touched, and neither is a
   * group of theirs that still holds something.
   */
  const strip = (container: Element): void => {
    for (const node of Array.from(container.childNodes)) {
      if (node.nodeType !== 1) continue;
      const shape = node as Element;
      const id = idOf(shape);
      if (id !== undefined && wanted.has(id)) {
        container.removeChild(shape);
        removed += 1;
        continue;
      }
      if (shape.namespaceURI !== P_NS || shape.localName !== "grpSp") continue;
      strip(shape);
      // A group with nothing left in it is not a group any more. PowerPoint
      // writes `<p:grpSp>` with at least one shape in it, and the schema says
      // the same; leaving an empty one behind would be this add-in producing
      // the very thing it refuses to produce elsewhere.
      if (slideShapes(shape).length === 0) container.removeChild(shape);
    }
  };
  strip(spTree);

  // Read back off the REBUILT slide rather than trusting the loop: this is the
  // evidence the caller reports, and the whole point of the feature is that the
  // element is gone.
  const left = (await readShapeTags(pkg, rebuilt)).filter((t) => t.element === request.element).length;

  await keepOnly(pkg, rebuilt);
  return { base64: await pkg.toBase64(), deckSlides, slidePath: rebuilt, removed, left };
}
