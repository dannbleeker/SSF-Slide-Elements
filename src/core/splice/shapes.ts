/**
 * An element's shapes, made safe to live in somebody else's slide.
 *
 * The catalogue hands over one string: the element's top-level shapes as they
 * were serialised out of the library deck. Four things in that string are true
 * of the library and not of the destination, and every one of them is a file
 * PowerPoint refuses to open or opens wrong:
 *
 * - **shape ids** are unique within a slide, and the destination slide has its
 *   own;
 * - **relationship ids** (`r:embed`, `r:id`, `r:link`) name entries in the
 *   LIBRARY slide's rels part, and mean something else in the destination's;
 * - **coordinates** are where the owner put the element on a library slide;
 * - **a table's real size** is its rows and columns, not the frame around them,
 *   so scaling the frame moves a table without resizing it.
 *
 * Nothing here opens a package or writes a part. It takes XML and answers XML,
 * which is what lets the suite check every rewrite against a real element out
 * of the real library with no PowerPoint anywhere.
 */
import { A_NS, P_NS, R_NS, child, children, parseXml } from "../pptx/xml.js";
import type { Move, Rect } from "./landing.js";
import { isIdentity } from "./landing.js";

/**
 * The root a fragment is parsed under.
 *
 * A parser needs one root and an element's markup is a RUN of top-level shapes,
 * so they are wrapped. The wrapper declares nothing, and it does not have to:
 * every prefix each shape uses is declared inside that shape. That is a
 * property of the harvest rather than a hope — it serialises each top-level
 * shape standalone, so the serialiser emits every declaration the fragment
 * needs, and `test/splice-shapes.test.ts` checks it holds for all 234 elements
 * in the committed library rather than for the one this comment was written
 * against.
 */
const FRAGMENT_ROOT = "ssf-fragment";

/** An element's shapes, parsed, with the wrapper still around them. */
export function parseFragment(xml: string): Document {
  return parseXml(`<${FRAGMENT_ROOT}>${xml}</${FRAGMENT_ROOT}>`);
}

/** The top-level shapes of a parsed fragment, in slide order. */
export function topLevel(fragment: Document): Element[] {
  const out: Element[] = [];
  const root = fragment.documentElement;
  for (const node of Array.from(root.childNodes)) {
    if (node.nodeType === 1) out.push(node as Element);
  }
  return out;
}

/** Every descendant element of `root`, and `root` itself. */
function everyElement(root: Element): Element[] {
  const out: Element[] = [root];
  for (let i = 0; i < out.length; i++) {
    const node = out[i];
    if (!node) continue;
    for (const kid of Array.from(node.childNodes)) {
      if (kid.nodeType === 1) out.push(kid as Element);
    }
  }
  return out;
}

/**
 * The highest `<p:cNvPr id>` anywhere in a shape tree.
 *
 * Read from the whole tree rather than from the top-level shapes: ids are
 * unique across the SLIDE, and a group's children carry them too. A renumber
 * that started above the top-level maximum alone would collide with the inside
 * of a group the user already had.
 */
export function highestShapeId(spTree: Element): number {
  let max = 0;
  for (const el of everyElement(spTree)) {
    if (el.namespaceURI !== P_NS || el.localName !== "cNvPr") continue;
    const n = Number(el.getAttribute("id") ?? 0);
    if (Number.isSafeInteger(n) && n > max) max = n;
  }
  return max;
}

/**
 * Give every shape in the fragment an id the destination slide is not using.
 *
 * Returns the first id still free afterwards, so a caller splicing several
 * elements onto one slide can carry on from it.
 *
 * Every `<p:cNvPr>` in the fragment, not only the top-level ones — a group's
 * children, and the `<p:pic>` inside an embedded object's `<mc:Fallback>`,
 * carry ids of their own and collide just as loudly.
 */
export function renumber(fragment: Document, from: number): number {
  let next = from;
  for (const shape of topLevel(fragment)) {
    for (const el of everyElement(shape)) {
      if (el.namespaceURI !== P_NS || el.localName !== "cNvPr") continue;
      el.setAttribute("id", String(next));
      next += 1;
    }
  }
  return next;
}

/**
 * Repoint every relationship reference in the fragment.
 *
 * Anchored on the NAMESPACE, not on a list of attribute names. Every reference
 * to a relationship in OOXML is an attribute in the relationships namespace —
 * `r:embed` on a blip, `r:id` on a chart, an embedded object or a tag list,
 * `r:link` on a linked picture, `r:dm`/`r:lo`/`r:qs`/`r:cs` on SmartArt — and a
 * rewriter that lists the ones it has seen is a rewriter that silently skips
 * the first one it has not. The sibling project's integrity checker exists
 * because a pass that "knew one of the ways a slide names them" deleted the
 * others.
 *
 * A reference with no entry in the map is left exactly as it was, and that is
 * not an oversight: an element whose markup names a relationship the catalogue
 * did not carry is a defect in the harvest, and leaving the reference alone
 * makes it visible to the integrity check as an unresolvable reference instead
 * of hiding it behind a plausible-looking id.
 */
export function repoint(fragment: Document, map: Map<string, string>): number {
  let changed = 0;
  for (const shape of topLevel(fragment)) {
    for (const el of everyElement(shape)) {
      for (const attr of Array.from(el.attributes)) {
        if (attr.namespaceURI !== R_NS) continue;
        const to = map.get(attr.value);
        if (to === undefined) continue;
        el.setAttributeNS(R_NS, attr.name, to);
        changed += 1;
      }
    }
  }
  return changed;
}

/** Every relationship id the fragment names, in the order it names them. */
export function relIdsIn(fragment: Document): string[] {
  const seen = new Set<string>();
  for (const shape of topLevel(fragment)) {
    for (const el of everyElement(shape)) {
      for (const attr of Array.from(el.attributes)) {
        if (attr.namespaceURI === R_NS) seen.add(attr.value);
      }
    }
  }
  return [...seen];
}

/**
 * The `<a:xfrm>` that positions a top-level shape, whatever kind it is.
 *
 * Three spellings: an ordinary shape and a picture keep it under `<p:spPr>`, a
 * group under `<p:grpSpPr>`, and a graphic frame — a table, a chart, SmartArt,
 * an embedded object — carries `<p:xfrm>` directly, in the PresentationML
 * namespace rather than DrawingML. The last one is the one that gets missed,
 * and 58 of the 117 elements in the 16:9 library are graphic frames.
 */
function frameOf(shape: Element): Element | undefined {
  const direct = child(shape, P_NS, "xfrm");
  if (direct) return direct;
  for (const holder of ["spPr", "grpSpPr"]) {
    const props = child(shape, P_NS, holder);
    const xfrm = props ? child(props, A_NS, "xfrm") : undefined;
    if (xfrm) return xfrm;
  }
  return undefined;
}

/** The rectangle a top-level shape occupies, or undefined when it does not say. */
export function rectOf(shape: Element): Rect | undefined {
  const xfrm = frameOf(shape);
  const off = xfrm ? child(xfrm, A_NS, "off") : undefined;
  const ext = xfrm ? child(xfrm, A_NS, "ext") : undefined;
  if (!off || !ext) return undefined;
  const x = Number(off.getAttribute("x"));
  const y = Number(off.getAttribute("y"));
  const cx = Number(ext.getAttribute("cx"));
  const cy = Number(ext.getAttribute("cy"));
  if (![x, y, cx, cy].every((n) => Number.isFinite(n))) return undefined;
  return { x, y, cx, cy };
}

/** The union of every top-level shape's rectangle: the element's own frame. */
export function unionOf(shapes: Element[]): Rect | undefined {
  let out: Rect | undefined;
  for (const shape of shapes) {
    const rect = rectOf(shape);
    if (!rect) continue;
    if (!out) {
      out = { ...rect };
      continue;
    }
    const right = Math.max(out.x + out.cx, rect.x + rect.cx);
    const bottom = Math.max(out.y + out.cy, rect.y + rect.cy);
    out.x = Math.min(out.x, rect.x);
    out.y = Math.min(out.y, rect.y);
    out.cx = right - out.x;
    out.cy = bottom - out.y;
  }
  return out;
}

/**
 * Scale a table's own grid, because PowerPoint draws a table from its rows and
 * columns and ignores the frame around them.
 *
 * `docs/DESIGN.md` section 3 records this from the other end — a table's BOX is
 * the sum of its columns and rows, "not its frame's `ext`, which PowerPoint
 * ignores when it draws the table". The same fact decides what resizing means:
 * a whole-slide element scaled down to clear a taller title would move its
 * table and leave it at its authored size, overlapping whatever it was scaled
 * away from. Half the 16:9 library is tables, so this is the common case rather
 * than an exotic one.
 *
 * Row heights are a MINIMUM in PowerPoint — a row grows to fit its text — so
 * scaling them down does not guarantee the drawn table shrinks by as much. It
 * is still the right instruction to write: the alternative is not scaling at
 * all.
 */
function scaleTable(shape: Element, sx: number, sy: number): void {
  for (const el of everyElement(shape)) {
    if (el.namespaceURI !== A_NS) continue;
    if (el.localName === "gridCol") {
      const w = Number(el.getAttribute("w"));
      if (Number.isFinite(w)) el.setAttribute("w", String(Math.max(1, Math.round(w * sx))));
    } else if (el.localName === "tr") {
      const h = Number(el.getAttribute("h"));
      if (Number.isFinite(h)) el.setAttribute("h", String(Math.max(1, Math.round(h * sy))));
    }
  }
}

/**
 * Move and scale the fragment's top-level shapes.
 *
 * Each shape is repositioned RELATIVE to the element's authored frame, so the
 * shapes keep their arrangement: a diagram whose arrow sits between two boxes
 * still has it between them afterwards.
 *
 * A GROUP is scaled by its frame alone. Its `<a:chOff>` and `<a:chExt>` are the
 * coordinate space its children are drawn in, and leaving those untouched while
 * the frame changes is exactly how PowerPoint scales a group's contents —
 * rewriting the children as well would apply the scale twice.
 *
 * An identity move returns without touching a byte, so an element that lands
 * where it was authored is spliced in exactly as the owner drew it.
 */
export function applyMove(shapes: Element[], from: Rect, move: Move): void {
  if (isIdentity(move)) return;
  for (const shape of shapes) {
    const xfrm = frameOf(shape);
    if (!xfrm) continue;
    const off = child(xfrm, A_NS, "off");
    const ext = child(xfrm, A_NS, "ext");
    if (off) {
      const x = Number(off.getAttribute("x"));
      const y = Number(off.getAttribute("y"));
      if (Number.isFinite(x)) off.setAttribute("x", String(Math.round(from.x + move.dx + (x - from.x) * move.sx)));
      if (Number.isFinite(y)) off.setAttribute("y", String(Math.round(from.y + move.dy + (y - from.y) * move.sy)));
    }
    if (ext) {
      const cx = Number(ext.getAttribute("cx"));
      const cy = Number(ext.getAttribute("cy"));
      if (Number.isFinite(cx)) ext.setAttribute("cx", String(Math.max(1, Math.round(cx * move.sx))));
      if (Number.isFinite(cy)) ext.setAttribute("cy", String(Math.max(1, Math.round(cy * move.sy))));
    }
    if (move.sx !== 1 || move.sy !== 1) scaleTable(shape, move.sx, move.sy);
  }
}

/**
 * Whether these shapes may be wrapped in one group.
 *
 * `docs/DESIGN.md` section 6 makes "as one group" the default, so the user
 * moves an element as one thing. A graphic frame is the exception, and PowerPoint
 * itself is the evidence: it refuses to group a table through its own UI. A
 * table inside a `<p:grpSp>` is markup PowerPoint has no way to have produced,
 * and this add-in does not ship the first one.
 *
 * Refusing is the right shape for this rather than failing: the element still
 * lands, it lands loose, and the pane says so.
 */
export function groupable(shapes: Element[]): boolean {
  if (shapes.length < 2) return false;
  return shapes.every((s) => !(s.namespaceURI === P_NS && s.localName === "graphicFrame"));
}

/**
 * Wrap the fragment's shapes in one group, in place.
 *
 * `<a:chOff>` and `<a:chExt>` are set EQUAL to the group's own offset and
 * extent, which makes the child coordinate space the same as the slide's — so
 * every shape inside keeps the coordinates it already has. Any other pair of
 * values would require rewriting all of them, and getting that arithmetic
 * subtly wrong is a group whose contents are offset by a few millimetres in a
 * way nobody notices until the deck is printed.
 */
export function groupShapes(fragment: Document, shapes: Element[], name: string, id: number): Element {
  const frame = unionOf(shapes) ?? { x: 0, y: 0, cx: 0, cy: 0 };
  const doc = fragment;
  const grpSp = doc.createElementNS(P_NS, "p:grpSp");

  const nvGrpSpPr = doc.createElementNS(P_NS, "p:nvGrpSpPr");
  const cNvPr = doc.createElementNS(P_NS, "p:cNvPr");
  cNvPr.setAttribute("id", String(id));
  cNvPr.setAttribute("name", name);
  nvGrpSpPr.appendChild(cNvPr);
  nvGrpSpPr.appendChild(doc.createElementNS(P_NS, "p:cNvGrpSpPr"));
  nvGrpSpPr.appendChild(doc.createElementNS(P_NS, "p:nvPr"));
  grpSp.appendChild(nvGrpSpPr);

  const grpSpPr = doc.createElementNS(P_NS, "p:grpSpPr");
  const xfrm = doc.createElementNS(A_NS, "a:xfrm");
  const off = doc.createElementNS(A_NS, "a:off");
  off.setAttribute("x", String(frame.x));
  off.setAttribute("y", String(frame.y));
  const ext = doc.createElementNS(A_NS, "a:ext");
  ext.setAttribute("cx", String(frame.cx));
  ext.setAttribute("cy", String(frame.cy));
  const chOff = doc.createElementNS(A_NS, "a:chOff");
  chOff.setAttribute("x", String(frame.x));
  chOff.setAttribute("y", String(frame.y));
  const chExt = doc.createElementNS(A_NS, "a:chExt");
  chExt.setAttribute("cx", String(frame.cx));
  chExt.setAttribute("cy", String(frame.cy));
  // `appendChild` four times, not `append`: `@xmldom/xmldom` implements the DOM
  // Level 3 surface and has no DOM4 `ParentNode.append`, so the tidier spelling
  // throws at run time on markup that typechecks perfectly.
  for (const part of [off, ext, chOff, chExt]) xfrm.appendChild(part);
  grpSpPr.appendChild(xfrm);
  grpSp.appendChild(grpSpPr);

  const root = doc.documentElement;
  // The group takes the first shape's place in the run, so an element that was
  // drawn behind another stays behind it.
  root.insertBefore(grpSp, shapes[0] ?? null);
  for (const shape of shapes) grpSp.appendChild(shape);
  return grpSp;
}

/**
 * Strip a shape's placeholder claim.
 *
 * A library shape can carry `<p:ph>`, which says "I am this slide's body". Put
 * onto a destination slide that already has a body placeholder, two shapes then
 * claim the same index: PowerPoint resolves that by inheriting geometry from
 * the layout, and the element jumps to wherever the layout's placeholder is,
 * losing the landing this splice just computed.
 *
 * The shape keeps every visual property it had. What it loses is the CLAIM,
 * which was never true of the destination.
 */
export function unplaceholder(shapes: Element[]): number {
  let stripped = 0;
  for (const shape of shapes) {
    for (const el of everyElement(shape)) {
      if (el.namespaceURI !== P_NS || el.localName !== "ph") continue;
      el.parentNode?.removeChild(el);
      stripped += 1;
    }
  }
  return stripped;
}

/** Direct children of a shape tree that are shapes rather than its properties. */
export function slideShapes(spTree: Element): Element[] {
  const out: Element[] = [];
  for (const node of Array.from(spTree.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.namespaceURI === P_NS && (el.localName === "nvGrpSpPr" || el.localName === "grpSpPr")) continue;
    out.push(el);
  }
  return out;
}

/**
 * The empty content placeholders on a slide.
 *
 * `docs/DESIGN.md` section 6: when a whole-slide element lands, the "Click to
 * add text" ghosts behind it are removed, and the TITLE placeholder stays. A
 * placeholder is empty when it has no text in it at all; one the user has typed
 * into is content, and content is never removed by an insert.
 */
export function emptyBodyPlaceholders(spTree: Element): Element[] {
  const out: Element[] = [];
  for (const shape of slideShapes(spTree)) {
    if (shape.namespaceURI !== P_NS || shape.localName !== "sp") continue;
    const nvSpPr = child(shape, P_NS, "nvSpPr");
    const nvPr = nvSpPr ? child(nvSpPr, P_NS, "nvPr") : undefined;
    const ph = nvPr ? child(nvPr, P_NS, "ph") : undefined;
    if (!ph) continue;
    const type = ph.getAttribute("type") ?? "body";
    if (type === "title" || type === "ctrTitle") continue;
    const txBody = child(shape, P_NS, "txBody");
    if (!txBody) continue;
    let text = "";
    for (const p of children(txBody, A_NS, "p")) {
      for (const el of everyElement(p)) {
        if (el.namespaceURI === A_NS && el.localName === "t") text += el.textContent ?? "";
      }
    }
    if (text.trim() === "") out.push(shape);
  }
  return out;
}
