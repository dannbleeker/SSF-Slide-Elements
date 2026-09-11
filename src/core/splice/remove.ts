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
 * shapes are found by the `SSF_SLIDE_ELEMENT` tag, per top-level shape, through
 * `readShapeTags` — which keys on the tag NAME rather than on the part's path,
 * because a deck that has been through another add-in carries ITS tags in the
 * same folder, numbered around ours. A user's own shape has no tag of ours and
 * cannot be reached from here at all.
 *
 * **Two things it deliberately leaves behind.** The tag parts the removed
 * shapes pointed at stay in the package as orphans, and so do any pictures they
 * carried: `scripts/package-integrity.mjs` treats an orphan as weight rather
 * than damage, and a sweep that deleted them would be a sweep deciding what
 * else in the user's deck still needs a part — which is the one class of
 * mistake that produces a file PowerPoint calls damaged.
 */
import { cloneSlide } from "../pptx/clone.js";
import { Pkg } from "../pptx/pkg.js";
import { readShapeTags } from "../pptx/tags.js";
import { P_NS, child } from "../pptx/xml.js";
import { keepOnly } from "./listing.js";

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
  /** How many top-level shapes came off. */
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

/** A top-level shape's `<p:cNvPr id>`, by the path the schema names for its kind. */
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
  for (const node of Array.from(spTree.childNodes)) {
    if (node.nodeType !== 1) continue;
    const shape = node as Element;
    const id = idOf(shape);
    if (id === undefined || !wanted.has(id)) continue;
    spTree.removeChild(shape);
    removed += 1;
  }

  // Read back off the REBUILT slide rather than trusting the loop: this is the
  // evidence the caller reports, and the whole point of the feature is that the
  // element is gone.
  const left = (await readShapeTags(pkg, rebuilt)).filter((t) => t.element === request.element).length;

  await keepOnly(pkg, rebuilt);
  return { base64: await pkg.toBase64(), deckSlides, slidePath: rebuilt, removed, left };
}
