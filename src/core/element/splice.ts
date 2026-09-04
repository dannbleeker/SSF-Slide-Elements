/**
 * Put an element onto a slide, in the file.
 *
 * This is the whole product. Office.js has no call that inserts arbitrary
 * markup onto a slide — `setSelectedDataAsync` will not take OOXML in
 * PowerPoint, and the shape collection can only add geometry it has a method
 * for — so an element with a gradient, a table, a chart or an icon cannot be
 * reproduced through the API without losing something. The route that loses
 * nothing is to edit the file: read the user's deck, splice the element's
 * markup into one slide's shape tree, and hand the result back.
 *
 * The result is inserted as a NEW slide and the original is removed. That is
 * not a detail to gloss over — it is why `src/host/undo.ts` exists and why the
 * pane says what it is about to do. What makes it safe rather than reckless is
 * that source and destination are the SAME DECK: the spliced slide carries the
 * deck's own layout, master and theme, so the element adopts the user's brand
 * exactly as it would if they had pasted it, with no theme drift to reconcile.
 *
 * Four things must be rewritten on the way in, and every one of them has a
 * failure that opens as "PowerPoint found a problem with this file":
 *
 * 1. **Part names.** The element's `ppt/media/image3.png` and the target's are
 *    different pictures with the same name. Every incoming part is renamed to
 *    a free one.
 * 2. **Nested relationship targets.** A chart's `.rels` points at its workbook
 *    by name, so renaming the workbook and not the pointer leaves a chart with
 *    no data.
 * 3. **Relationship ids.** `rId3` in the library slide and `rId3` in the
 *    target slide name different things. The ids in the element's markup are
 *    rewritten to the ones the target slide hands out.
 * 4. **Shape ids.** `<p:cNvPr id>` must be unique within a slide. A collision
 *    does not always break the file, which is worse: it breaks selection,
 *    animation and grouping in ways that look like PowerPoint misbehaving.
 */
import { Pkg, extensionOf, resolveTarget } from "../pptx/pkg.js";
import { A_NS, P_NS, PKG_REL_NS, R_NS, child, element, elements, parseXml, serializeXml } from "../pptx/xml.js";
import { EXTERNAL } from "../pptx/parts.js";
import type { Box, ElementPayload } from "../catalogue/types.js";

/** Where the element should land, in EMU, top-left corner. */
export interface Placement {
  x: number;
  y: number;
  /**
   * Uniform scale applied about the element's own top-left. 1 leaves it as
   * authored. Used when a 16:9 element goes into a 4:3 deck, never to make an
   * element "fit" — a squashed element is worse than one that overhangs.
   */
  scale?: number;
}

export interface SpliceReport {
  /** How many top-level shapes were added. */
  shapes: number;
  /** Parts copied in, by their new names. */
  parts: string[];
  /** Relationship ids the target slide handed out. */
  rels: string[];
  /** Where it actually landed, after placement and scaling. */
  bounds: Box;
  /**
   * Relationship ids the element named that nothing carried, and which were
   * therefore removed from its markup. Always empty for a library harvested by
   * this repo's own harvest; non-empty means a payload and an engine that
   * disagree about what travels, which is worth surfacing rather than swallowing.
   */
  droppedRefs: string[];
}

/**
 * Decide a free part name in the target package that keeps the family and the
 * extension of the incoming one.
 *
 * `ppt/media/image3.png` becomes `ppt/media/image<next>.png`;
 * `ppt/embeddings/Microsoft_Excel_Worksheet.xlsx` has no number to bump, so it
 * gets one. Names outside those shapes fall back to a numbered name in the same
 * folder rather than being refused: the folder is already allowlisted, and a
 * name this does not recognise is a producer's choice, not an attack.
 */
export function freeName(pkg: Pkg, incoming: string): string {
  const cut = incoming.lastIndexOf("/");
  const folder = incoming.slice(0, cut);
  const file = incoming.slice(cut + 1);
  const ext = extensionOf(file);
  const dotted = ext ? `.${ext}` : "";
  const stem = ext ? file.slice(0, file.length - dotted.length) : file;
  const base = /^(.*?)(\d+)$/.exec(stem)?.[1] ?? stem;
  const prefix = `${folder}/${base}`;
  return `${prefix}${pkg.nextNumber(prefix, dotted)}${dotted}`;
}

/**
 * Every `<p:cNvPr id>` in a subtree, renumbered from `start`.
 *
 * Walks the whole subtree rather than the top level, because a group's children
 * carry ids in the same slide-wide space as the group. Answers the next free
 * value so a caller renumbering several shapes does not have to track it.
 *
 * **The same old id always becomes the same new id**, and that is not an
 * optimisation — it is the `<mc:AlternateContent>` rule. A shape that has a
 * modern form and a fallback carries BOTH, in `<mc:Choice>` and
 * `<mc:Fallback>`, and the two branches deliberately share one `<p:cNvPr id>`
 * because only one of them is ever live. Handing them different ids splits one
 * shape into two that PowerPoint then disagrees with itself about. The cover
 * slide of the shipped library is exactly this case, so a renumbering that got
 * it wrong would have been wrong on the very first thing anybody inserted onto.
 *
 * Within one top-level shape there is no other way for an id to repeat: ids are
 * unique per slide, so a repeat inside a single harvested shape is an
 * AlternateContent pair by construction.
 */
export function renumberShapeIds(root: Element, start: number): number {
  let next = start;
  const assigned = new Map<string, number>();
  const walk = (node: Element): void => {
    if (node.localName === "cNvPr" && node.namespaceURI === P_NS) {
      const old = node.getAttribute("id") ?? "";
      let fresh = assigned.get(old);
      if (fresh === undefined) {
        fresh = next;
        next += 1;
        assigned.set(old, fresh);
      }
      node.setAttribute("id", String(fresh));
    }
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) walk(c as Element);
    }
  };
  walk(root);
  return next;
}

/** The highest `<p:cNvPr id>` already on a slide. Ids must not collide with it. */
export function highestShapeId(tree: Element): number {
  let max = 1;
  for (const nv of elements(tree, P_NS, "cNvPr")) {
    const n = Number(nv.getAttribute("id"));
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max;
}

/**
 * Rewrite every relationship-namespaced attribute in a subtree through a map,
 * and remove the ones the map does not answer.
 *
 * The breadth matches `relationshipIdsIn`, for the same reason: an `r:embed` on
 * a blip, an `r:link` on a linked picture and an `r:embed` on the SVG beside an
 * icon are all references, and a rewrite that knows about one of them leaves
 * the others pointing at whatever the target slide's `rId3` happens to be. That
 * is not a crash — it is somebody else's logo inside your element.
 *
 * **An unmapped id is REMOVED, not left alone.** An earlier version left them,
 * on the reasoning that an uncarried relationship is inert; it is not. A slide
 * naming a relationship its `.rels` does not define is precisely what
 * PowerPoint calls a damaged file, and the shipped library produced 41 of them
 * on one slide. The harvest now carries everything a shape can name, so this
 * should never fire — which is exactly why it is here. It is the invariant, not
 * the plan: a library deck can contain a relationship type nobody has thought
 * of yet, and the failure mode for that must be a missing decoration rather
 * than a file that will not open.
 *
 * Where the reference IS the element — `<p:tags r:id>` carries nothing else —
 * the element goes with it, so no empty husk is left behind.
 */
export function rewriteRelIds(root: Element, map: Map<string, string>): string[] {
  const dropped: string[] = [];
  const orphans: Element[] = [];
  const walk = (node: Element): void => {
    const attrs = node.attributes;
    const stale: string[] = [];
    for (let i = 0; i < (attrs?.length ?? 0); i++) {
      const attr = attrs?.item(i);
      if (!attr) continue;
      if (attr.namespaceURI !== R_NS && !attr.name.startsWith("r:")) continue;
      const mapped = map.get(attr.value);
      if (mapped !== undefined) {
        attr.value = mapped;
        continue;
      }
      stale.push(attr.name);
      dropped.push(attr.value);
    }
    if (stale.length > 0) {
      // A `<p:tags>` is nothing but its reference; an attribute on a shape is
      // one decoration among many. Remove the smallest thing that leaves valid
      // markup behind.
      if (node.localName === "tags") orphans.push(node);
      else for (const name of stale) node.removeAttribute(name);
    }
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) walk(c as Element);
    }
  };
  walk(root);
  for (const node of orphans) node.parentNode?.removeChild(node);
  return dropped;
}

/**
 * Move and optionally scale a top-level shape.
 *
 * Only the shape's OWN `<a:xfrm>` is touched. A group's children are in the
 * group's child coordinate space and move with it for free; rewriting them too
 * would move them twice. A `<p:graphicFrame>` keeps its transform in an
 * `<p:xfrm>` in the PresentationML namespace rather than a DrawingML one, which
 * is the kind of difference that makes a table the one shape that does not
 * move.
 */
export function placeShape(shape: Element, dx: number, dy: number, scale: number, origin: Box): void {
  const holder = child(shape, P_NS, "spPr") ?? child(shape, P_NS, "grpSpPr") ?? child(shape, P_NS, "xfrm") ?? undefined;
  const xfrm = holder ? (holder.localName === "xfrm" ? holder : child(holder, A_NS, "xfrm")) : undefined;
  if (!xfrm) return;
  const off = child(xfrm, A_NS, "off");
  const ext = child(xfrm, A_NS, "ext");
  if (off) {
    const x = Number(off.getAttribute("x"));
    const y = Number(off.getAttribute("y"));
    if (Number.isFinite(x)) off.setAttribute("x", String(Math.round(origin.x + (x - origin.x) * scale + dx)));
    if (Number.isFinite(y)) off.setAttribute("y", String(Math.round(origin.y + (y - origin.y) * scale + dy)));
  }
  if (ext && scale !== 1) {
    const cx = Number(ext.getAttribute("cx"));
    const cy = Number(ext.getAttribute("cy"));
    if (Number.isFinite(cx)) ext.setAttribute("cx", String(Math.max(1, Math.round(cx * scale))));
    if (Number.isFinite(cy)) ext.setAttribute("cy", String(Math.max(1, Math.round(cy * scale))));
  }
  // A group states its CHILD coordinate space as well as its own box. Scaling
  // the outer box without the child extent stretches the group's contents away
  // from it, which reads as the element coming apart.
  if (scale !== 1) {
    const chExt = child(xfrm, A_NS, "chExt");
    if (chExt) {
      const cx = Number(chExt.getAttribute("cx"));
      const cy = Number(chExt.getAttribute("cy"));
      if (Number.isFinite(cx)) chExt.setAttribute("cx", String(Math.max(1, Math.round(cx * scale))));
      if (Number.isFinite(cy)) chExt.setAttribute("cy", String(Math.max(1, Math.round(cy * scale))));
    }
  }
}

/**
 * Copy the payload's parts into the package under free names, and answer the
 * mapping from the name the payload used to the name it now has.
 *
 * The nested `.rels` are copied too and then REWRITTEN: a chart's relationships
 * name its workbook by a path relative to the chart, and that path is stale the
 * moment the workbook is renamed. This is step 2 of the four in the file
 * header, and it is the one with no visible symptom until somebody opens the
 * chart's data.
 */
async function copyParts(pkg: Pkg, payload: ElementPayload): Promise<Map<string, string>> {
  const renamed = new Map<string, string>();

  // Names first, for every non-`.rels` part, so the rewrite below can resolve
  // any pointer regardless of the order the parts happen to be listed in.
  for (const part of payload.parts) {
    if (part.path.includes("/_rels/")) continue;
    renamed.set(part.path, freeName(pkg, part.path));
  }

  for (const part of payload.parts) {
    if (part.path.includes("/_rels/")) continue;
    const to = renamed.get(part.path);
    if (to === undefined) continue;
    if (part.text !== undefined) {
      pkg.setText(to, part.text);
      await pkg.addContentTypeOverride(to, part.contentType);
    } else if (part.base64 !== undefined) {
      pkg.setBytes(to, fromBase64(part.base64));
      // Media is declared by EXTENSION, and a package that already has PNGs
      // has the default. One that does not gets it — an undeclared extension
      // is a part PowerPoint refuses to load.
      const ext = extensionOf(to);
      if (ext) await pkg.addContentTypeDefault(ext, part.contentType);
    }
  }

  // Now the relationship parts, with their targets pointed at the new names.
  for (const part of payload.parts) {
    if (!part.path.includes("/_rels/") || part.text === undefined) continue;
    const owner = ownerOfRels(part.path);
    const newOwner = renamed.get(owner);
    if (newOwner === undefined) continue;
    const doc = parseXml(part.text);
    for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
      if (rel.getAttribute("TargetMode") === EXTERNAL) continue;
      const target = rel.getAttribute("Target") ?? "";
      const absolute = resolveTarget(owner, target);
      const mapped = renamed.get(absolute);
      if (mapped === undefined) continue;
      rel.setAttribute("Target", relativeTo(newOwner, mapped));
    }
    pkg.setText(Pkg.relsPathFor(newOwner), serializeXml(doc));
  }

  return renamed;
}

/** `ppt/charts/_rels/chart1.xml.rels` describes `ppt/charts/chart1.xml`. */
export function ownerOfRels(relsPath: string): string {
  return relsPath.replace("/_rels/", "/").replace(/\.rels$/, "");
}

/** A package path expressed relative to the part that points at it. */
export function relativeTo(owner: string, target: string): string {
  const from = owner.split("/").slice(0, -1);
  const to = target.split("/");
  const file = to.pop() ?? "";
  let i = 0;
  while (i < from.length && i < to.length && from[i] === to[i]) i += 1;
  const up = from.length - i;
  return [...Array<string>(up).fill(".."), ...to.slice(i), file].join("/");
}

/** Base64 to bytes without Node's Buffer, so the same code runs in the pane. */
export function fromBase64(b64: string): Uint8Array {
  const binary = atob(b64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/**
 * Splice an element into one slide of a package.
 *
 * Mutates `pkg`. The caller decides what to do with the result — the pane hands
 * it to PowerPoint, the suite reads it back and checks it.
 */
export async function splice(
  pkg: Pkg,
  slidePath: string,
  payload: ElementPayload,
  at: Placement,
): Promise<SpliceReport> {
  const doc = await pkg.doc(slidePath);
  const tree = element(doc, P_NS, "spTree");
  if (!tree) throw new Error(`ssf-slide-elements: ${slidePath} has no shape tree to add to`);

  const renamed = await copyParts(pkg, payload);

  // Relationships the SLIDE needs, in the ids the slide hands out.
  const relMap = new Map<string, string>();
  const newRels: string[] = [];
  for (const rel of payload.rels) {
    const target = rel.external ? rel.path : renamed.get(rel.path);
    if (target === undefined) continue;
    const rId = await pkg.addRel(
      slidePath,
      rel.type,
      rel.external ? target : relativeTo(slidePath, target),
      rel.external ? EXTERNAL : undefined,
    );
    relMap.set(rel.rId, rId);
    newRels.push(rId);
  }

  const droppedRefs: string[] = [];
  const scale = at.scale ?? 1;
  const dx = at.x - payload.bounds.x;
  const dy = at.y - payload.bounds.y;
  let nextId = highestShapeId(tree) + 1;

  for (const xml of payload.shapes) {
    // Each shape is parsed as its own document and then imported, rather than
    // being appended across documents. `@xmldom/xmldom` will let a foreign node
    // be appended and the namespaces it inherited from the old root do not come
    // with it — producing markup that serialises, and that PowerPoint reads as
    // a shape with no properties.
    const fragment = parseXml(withNamespaces(xml));
    const root = fragment.documentElement;
    if (!root) continue;
    droppedRefs.push(...rewriteRelIds(root, relMap));
    nextId = renumberShapeIds(root, nextId);
    placeShape(root, dx, dy, scale, payload.bounds);
    tree.appendChild(doc.importNode(root, true));
  }

  pkg.write(slidePath, doc);

  return {
    shapes: payload.shapes.length,
    parts: [...renamed.values()],
    rels: newRels,
    bounds: {
      x: at.x,
      y: at.y,
      cx: Math.round(payload.bounds.cx * scale),
      cy: Math.round(payload.bounds.cy * scale),
    },
    droppedRefs,
  };
}

/**
 * The namespace declarations a harvested shape needs to parse on its own.
 *
 * A shape serialised out of a slide carries only the prefixes it uses; the
 * DECLARATIONS were on `<p:sld>` and stayed behind. Parsing the fragment alone
 * then fails, or worse, succeeds with the prefixes unbound and every `p:` and
 * `a:` element in no namespace at all — which serialises back into something
 * that looks right in a diff and is not a shape.
 *
 * Declared on a wrapper only when the fragment does not already declare them,
 * so a shape harvested by a producer that wrote its own declarations is left
 * exactly as it was.
 */
export function withNamespaces(xml: string): string {
  const need: [string, string][] = [
    ["p", P_NS],
    ["a", A_NS],
    ["r", R_NS],
    ["mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"],
  ];
  const head = xml.slice(0, xml.indexOf(">") + 1);
  const missing = need.filter(([prefix]) => !head.includes(`xmlns:${prefix}=`));
  if (missing.length === 0) return xml;
  const decls = missing.map(([prefix, ns]) => ` xmlns:${prefix}="${ns}"`).join("");
  const cut = xml.indexOf(">");
  // A self-closing root — `<p:sp/>` — has its slash immediately before the `>`,
  // and inserting after it produces `<p:sp/ xmlns:p="…">`, which is not XML.
  const selfClosing = xml[cut - 1] === "/";
  const insertAt = selfClosing ? cut - 1 : cut;
  return xml.slice(0, insertAt) + decls + xml.slice(insertAt);
}
