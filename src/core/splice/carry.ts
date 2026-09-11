/**
 * The parts an element brings with it, copied into the destination package.
 *
 * An element is rarely only markup. A picture is a `.png` or an `.emf` in
 * `ppt/media/`; an embedded object is a `.bin` in `ppt/embeddings/` plus the
 * picture PowerPoint draws in its place; a chart is a part with its own
 * relationships to a whole Excel workbook; and most of the library's shapes
 * carry a tag part the owner's PowerPoint wrote. Every one of those has to
 * arrive in the user's deck under a name that deck is not already using, with a
 * content type declared for it, and with every relationship pointing at the new
 * name rather than the library's.
 *
 * Three rules, each learned rather than chosen:
 *
 * - **What to copy comes from the element's relationships, never from its parts
 *   list.** The harvest collects each part once per deck, so the second element
 *   to use a picture lists nothing for it while still naming it: measured on
 *   the committed library, 64 relationship targets across the two sizes are
 *   absent from their own element's `parts`. `types.ts` says the same thing
 *   from the other end.
 * - **A copied part's own relationships are copied too, and rewritten.** The
 *   library's one chart points at `../embeddings/Microsoft_Excel_Worksheet.xlsx`
 *   through its own rels part. Copy the chart and not that, and PowerPoint
 *   opens a chart it cannot edit; copy both and leave the target alone, and the
 *   copy points at the library's workbook, which is not in the deck at all.
 * - **Every copied part gets an Override.** An Override is legal for any part
 *   name and takes precedence over a Default, so declaring one is correct
 *   whatever the destination already says about that extension. The alternative
 *   — deciding per part whether the library used a Default or an Override —
 *   is a guess with an undeclared part as its failure, and PowerPoint refuses
 *   a package with one of those without naming it.
 */
import type { MarkupRel } from "../catalogue/types.js";
import { Pkg } from "../pptx/pkg.js";
import { PKG_REL_NS, elements, parseXml, serializeXml } from "../pptx/xml.js";

/**
 * Reads a carried part out of the catalogue, by the path the library knew it
 * by. Answers undefined for a part the store does not hold.
 *
 * A function rather than a map, because the pane fetches these over the network
 * one at a time and a whole library of parts is about 16 MB.
 */
export type PartStore = (path: string) => Promise<Uint8Array | string | undefined>;

export interface CarryRequest {
  pkg: Pkg;
  /** The part that will own the new relationships: the slide the element lands on. */
  owner: string;
  /** The element's relationships, as the catalogue recorded them. */
  rels: MarkupRel[];
  /** Content type by library path, from the catalogue's `carried` map. */
  types: Record<string, string>;
  store: PartStore;
  /**
   * A last pass over a copied XML part's text, before it is written.
   *
   * One caller and one purpose: pinning the scheme colours inside a carried
   * chart when the user asked for the library's own colours
   * (`docs/DESIGN.md` section 7). The library's one chart states 17 of them,
   * and a switch that rewrote the shapes but not the chart they sit beside
   * would leave the two disagreeing.
   *
   * Given the LIBRARY path, so a caller can tell one part from another, and
   * never applied to a `.rels` part — those carry targets, not colours, and
   * this runs before they are rewritten.
   */
  transform?: (path: string, xml: string) => string;
}

export interface Carried {
  /** The element's old relationship ids to the ones now in the destination. */
  ids: Map<string, string>;
  /** Every part copied, by the library path it came from. */
  parts: Map<string, string>;
}

/**
 * The directory a part lives in, as segments.
 *
 * Deliberately not sharing an implementation with `resolveTarget` in `pkg.ts`,
 * which goes the other way and is held to a corpus by `test/integrity.test.ts`.
 * `test/architecture.test.ts` holds that resolver to one file by its own
 * signature, and this is not it: this builds a relationship Target from two
 * part names, where that one reads a Target and answers a part name.
 */
function directoryOf(part: string): string[] {
  const segments = part.split("/");
  segments.pop();
  return segments;
}

/**
 * A relationship `Target` that points from one part to another.
 *
 * Relative, with `..` for each directory the two do not share, which is the
 * spelling PowerPoint itself writes: a slide names a picture
 * `../media/image3.png`. An absolute `/ppt/media/image3.png` is legal OPC and
 * PowerPoint accepts it, but a package where half the targets are written one
 * way and half the other is a package nobody can eyeball.
 */
export function targetFrom(owner: string, part: string): string {
  const from = directoryOf(owner);
  const to = part.split("/");
  const name = to.pop() ?? part;
  let shared = 0;
  while (shared < from.length && shared < to.length && from[shared] === to[shared]) shared += 1;
  const up = new Array(from.length - shared).fill("..");
  return [...up, ...to.slice(shared), name].join("/");
}

/**
 * A part name in the same family that the destination package is not using.
 *
 * The family is the part's name with its trailing digits removed, so
 * `ppt/media/image7.emf` extends the destination's own `image` sequence and
 * `ppt/embeddings/Microsoft_Excel_Worksheet.xlsx`, which has no number at all,
 * becomes `…Worksheet1.xlsx`. Both are names the package does not hold, which
 * is the only thing this has to guarantee.
 *
 * Through `Pkg`'s own counter, which reads the package once and stays current
 * as parts are written — so copying forty parts does not walk the zip forty
 * times inside a task-pane WebView.
 */
export function freeName(pkg: Pkg, part: string): string {
  const dot = part.lastIndexOf(".");
  const slash = part.lastIndexOf("/");
  const extension = dot > slash ? part.slice(dot) : "";
  const stem = dot > slash ? part.slice(0, dot) : part;
  const family = stem.replace(/\d+$/, "");
  return `${family}${pkg.nextNumber(family, extension)}${extension}`;
}

/**
 * A fingerprint of some bytes: 64 bits, as hex, computed here rather than asked
 * of a platform.
 *
 * Two independent 32-bit hashes rather than one 64-bit one, because doing 64
 * bits in JavaScript means `BigInt` and this runs over every picture an element
 * carries. Not a cryptographic hash and not used as one: it names a part, and
 * the name carries the byte LENGTH beside it, so two different pictures would
 * have to collide in both to be mistaken for each other.
 */
function fingerprint(bytes: Uint8Array): string {
  let fnv = 0x811c9dc5;
  let djb = 5381;
  for (let i = 0; i < bytes.length; i += 1) {
    const b = bytes[i] as number;
    fnv = Math.imul(fnv ^ b, 0x01000193);
    djb = Math.imul(djb, 33) ^ b;
  }
  const hex = (n: number): string => (n >>> 0).toString(16).padStart(8, "0");
  return hex(fnv) + hex(djb);
}

/** Whether a library part is media, which is the one family that may be shared. */
const isMedia = (part: string): boolean => part.startsWith("ppt/media/");

/**
 * The name a carried MEDIA part gets: one derived from its own bytes.
 *
 * This is what makes the second insert of a picture free. A name from
 * `freeName` extends the destination's `image` sequence, so inserting the same
 * marker twice wrote `image3.emf` and `image4.emf` — byte for byte the same 29
 * KB picture, twice. Measured on the validators' deck: four inserts of
 * `markeringer-1` left four identical copies and cost 11.6 KB each. A user who
 * stamps thirty slides carries thirty.
 *
 * Derived from the CONTENT rather than found by searching the package, and that
 * is the whole point: a search means decompressing every picture in the user's
 * deck on every insert, which on a deck full of photographs is the cost this
 * engine spent a day removing from the base64 path. `pkg.has(name)` is one
 * lookup in an index the package already holds.
 *
 * Only `ppt/media/`. An embedded workbook or a chart part is a document, not a
 * picture: two charts sharing one workbook would mean editing one edits both,
 * which is not what anybody asked for. PowerPoint shares identical images
 * itself and does not share those, and this follows it.
 */
function mediaName(part: string, bytes: Uint8Array): string {
  const dot = part.lastIndexOf(".");
  const extension = dot > part.lastIndexOf("/") ? part.slice(dot) : "";
  return `ppt/media/ssf-${fingerprint(bytes)}-${bytes.length}${extension}`;
}

/**
 * Copy one carried part and everything it reaches, and answer its new name.
 *
 * Depth-first through the part's own relationships, so a chart's workbook is in
 * the package before the chart's rels part is written to name it. Already-copied
 * parts answer their existing name, so an element using one picture twice
 * copies it once — and a picture the package already holds from an EARLIER
 * insert is not copied at all, because its name is its content.
 */
async function copyPart(request: CarryRequest, carried: Carried, part: string): Promise<string> {
  const already = carried.parts.get(part);
  if (already !== undefined) return already;

  const bytes = await request.store(part);
  if (bytes === undefined) {
    // Named, and refused. A missing part is the one failure that must not be
    // survived quietly: the alternative is a package whose markup points at a
    // relationship pointing at nothing, which is exactly what PowerPoint calls
    // a damaged file and what `scripts/package-integrity.mjs` was written to
    // catch before a user ever sees it.
    throw new Error(`ssf-slide-elements: the catalogue has no part "${part}", which this element needs`);
  }

  const shareable = typeof bytes !== "string" && isMedia(part);
  const name = shareable ? mediaName(part, bytes) : freeName(request.pkg, part);
  // Reserved before the recursion below, so a part that reaches itself through
  // its own relationships cannot loop.
  carried.parts.set(part, name);

  // A picture this package already holds under its own content's name is the
  // same picture, and writing it again would be writing it twice. The content
  // type is still declared: `addContentTypeOverride` answers at once when the
  // declaration is already there, and skipping it would leave a package whose
  // part is undeclared if the earlier copy ever arrived another way.
  const shared = shareable && request.pkg.has(name);
  if (!shared) {
    if (typeof bytes === "string") {
      const text = request.transform && part.endsWith(".xml") ? request.transform(part, bytes) : bytes;
      request.pkg.setText(name, text);
    } else request.pkg.setBytes(name, bytes);
  }

  const type = request.types[part];
  if (type !== undefined) await request.pkg.addContentTypeOverride(`/${name}`, type);
  if (shared) return name;

  // The part's own relationships, if the library kept any for it.
  const sourceRels = Pkg.relsPathFor(part);
  const relsText = await request.store(sourceRels);
  if (relsText === undefined) return name;
  const doc = parseXml(typeof relsText === "string" ? relsText : new TextDecoder().decode(relsText));
  for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const reached = resolveFrom(part, target);
    const copiedTo = await copyPart(request, carried, reached);
    rel.setAttribute("Target", targetFrom(name, copiedTo));
  }
  request.pkg.setText(Pkg.relsPathFor(name), serializeXml(doc));
  return name;
}

/**
 * A relationship target resolved against the part that owns it.
 *
 * The engine's own resolver lives on `Pkg` and answers what THIS package holds;
 * this one is asked about the LIBRARY's package, which is not open here — only
 * its parts are, one fetch at a time. So it is the plain arithmetic, with no
 * package to ask, and it handles the two spellings a rels part actually uses: a
 * relative target with `..` segments, and an absolute one with a leading slash.
 */
function resolveFrom(owner: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segments = directoryOf(owner);
  for (const step of target.split("/")) {
    if (step === "." || step === "") continue;
    if (step === "..") segments.pop();
    else segments.push(step);
  }
  return segments.join("/");
}

/**
 * Copy everything an element carries into the package, and answer how its
 * relationship ids were renamed.
 *
 * The map is what `repoint` in `shapes.ts` needs: the element's markup names
 * `rId5`, the destination slide's rels part now names the same picture
 * something else, and every reference has to move together.
 */
export async function carry(request: CarryRequest): Promise<Carried> {
  const carried: Carried = { ids: new Map(), parts: new Map() };
  for (const rel of request.rels) {
    if (rel.external) {
      const id = await request.pkg.addRel(request.owner, rel.type, rel.target, true);
      carried.ids.set(rel.id, id);
      continue;
    }
    const name = await copyPart(request, carried, rel.target);
    const id = await request.pkg.addRel(request.owner, rel.type, targetFrom(request.owner, name));
    carried.ids.set(rel.id, id);
  }
  return carried;
}
