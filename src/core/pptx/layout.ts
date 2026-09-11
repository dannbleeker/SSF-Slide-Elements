/**
 * Where the destination slide's own placeholders are.
 *
 * A whole-slide element lands below the destination's title and inside its body
 * area (`docs/DESIGN.md` section 5), and neither of those is written on the
 * slide. PowerPoint's placeholders INHERIT: a slide's title placeholder
 * routinely carries no `<a:xfrm>` at all and takes its geometry from the
 * layout, which takes its own from the master. A reader that only looks at the
 * slide finds nothing on almost every real deck and concludes the slide has no
 * title — so a customer deck with a two-line title gets an element straight
 * through it, which is the exact failure section 5 exists to prevent.
 *
 * So the chain is walked: slide, then its layout, then that layout's master.
 * Nothing here decides anything about landing; it answers rectangles, and
 * `splice/landing.ts` decides.
 */
import { Pkg } from "./pkg.js";
import { REL_TYPE } from "./parts.js";
import { A_NS, P_NS, PKG_REL_NS, child, elements } from "./xml.js";

/** A rectangle in EMU. The same shape `splice/landing.ts` works in. */
export interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/**
 * Which placeholder is wanted.
 *
 * Two, because two is what the landing rule asks about. A title is `title` or
 * `ctrTitle`; everything that holds content is a body, and PowerPoint spells
 * that `body`, `obj`, `subTitle`, `tbl`, `chart`, `pic` or nothing at all —
 * an omitted `type` means body, which is the default the schema gives it and
 * the case a reader keying on the attribute misses entirely.
 */
export type Want = "title" | "body";

const TITLES = new Set(["title", "ctrTitle"]);
/** Never content: the running furniture along the bottom of a slide. */
const CHROME = new Set(["ftr", "sldNum", "dt"]);

/** The `<p:ph>` of a shape, if it claims to be a placeholder. */
function placeholderOf(shape: Element): Element | undefined {
  for (const node of Array.from(shape.childNodes)) {
    if (node.nodeType !== 1) continue;
    const nvPr = child(node as Element, P_NS, "nvPr");
    if (nvPr) return child(nvPr, P_NS, "ph");
  }
  return undefined;
}

/** Whether a placeholder is the kind being looked for. */
function wants(ph: Element, want: Want): boolean {
  const type = ph.getAttribute("type") ?? "body";
  if (want === "title") return TITLES.has(type);
  return !TITLES.has(type) && !CHROME.has(type);
}

/** A shape's own rectangle, when it carries one. */
function rectOf(shape: Element): Rect | undefined {
  const spPr = child(shape, P_NS, "spPr");
  const xfrm = spPr ? child(spPr, A_NS, "xfrm") : child(shape, P_NS, "xfrm");
  const off = xfrm ? child(xfrm, A_NS, "off") : undefined;
  const ext = xfrm ? child(xfrm, A_NS, "ext") : undefined;
  if (!off || !ext) return undefined;
  const rect = {
    x: Number(off.getAttribute("x")),
    y: Number(off.getAttribute("y")),
    cx: Number(ext.getAttribute("cx")),
    cy: Number(ext.getAttribute("cy")),
  };
  return Object.values(rect).every((n) => Number.isFinite(n)) ? rect : undefined;
}

/** The shape tree of a slide, layout or master part. */
async function spTreeOf(pkg: Pkg, part: string): Promise<Element | undefined> {
  if (!pkg.has(part)) return undefined;
  const doc = await pkg.doc(part);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  return cSld ? child(cSld, P_NS, "spTree") : undefined;
}

/** The one part of a given relationship type that a part points at, if any. */
async function relatedOfType(pkg: Pkg, owner: string, type: string): Promise<string | undefined> {
  const relsPath = Pkg.relsPathFor(owner);
  if (!pkg.has(relsPath)) return undefined;
  const rels = await pkg.doc(relsPath);
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    if (rel.getAttribute("Type") !== type) continue;
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const path = pkg.resolved(owner, target);
    if (pkg.has(path)) return path;
  }
  return undefined;
}

/** The layout a slide is built on. */
export function layoutOf(pkg: Pkg, slidePath: string): Promise<string | undefined> {
  return relatedOfType(pkg, slidePath, REL_TYPE.slideLayout);
}

/** The master a layout belongs to. */
export function masterOf(pkg: Pkg, layoutPath: string): Promise<string | undefined> {
  return relatedOfType(pkg, layoutPath, REL_TYPE.slideMaster);
}

/**
 * A placeholder's rectangle, following the inheritance chain.
 *
 * The slide first, because a slide that positions its own title is the one deck
 * that would be wrong to overrule. Then the layout, then the master, which is
 * the order PowerPoint resolves them in.
 *
 * A placeholder matched by `idx` where the slide gives one, because a layout
 * with two content placeholders side by side distinguishes them only that way,
 * and taking the first would land an element in the left column of a two-column
 * layout whatever the user had selected. A title has no `idx`, and matching on
 * type alone is right for it.
 */
export async function placeholderRect(pkg: Pkg, slidePath: string, want: Want): Promise<Rect | undefined> {
  const slideTree = await spTreeOf(pkg, slidePath);
  let idx: string | undefined;
  if (slideTree) {
    for (const node of Array.from(slideTree.childNodes)) {
      if (node.nodeType !== 1) continue;
      const shape = node as Element;
      const ph = placeholderOf(shape);
      if (!ph || !wants(ph, want)) continue;
      const own = rectOf(shape);
      if (own) return own;
      idx = ph.getAttribute("idx") ?? undefined;
      break;
    }
  }

  const layout = await layoutOf(pkg, slidePath);
  const master = layout ? await masterOf(pkg, layout) : undefined;
  for (const part of [layout, master]) {
    if (!part) continue;
    const tree = await spTreeOf(pkg, part);
    if (!tree) continue;
    for (const node of Array.from(tree.childNodes)) {
      if (node.nodeType !== 1) continue;
      const shape = node as Element;
      const ph = placeholderOf(shape);
      if (!ph || !wants(ph, want)) continue;
      // The idx the slide named, when it named one. A layout's own placeholders
      // carry the same indices the slide's do, which is what makes them the
      // same placeholder rather than merely the same kind.
      if (idx !== undefined && (ph.getAttribute("idx") ?? undefined) !== idx) continue;
      const rect = rectOf(shape);
      if (rect) return rect;
    }
  }
  return undefined;
}

/** The title and body rectangles of a slide, as `splice/landing.ts` wants them. */
export async function framesOf(pkg: Pkg, slidePath: string): Promise<{ title?: Rect; body?: Rect }> {
  const title = await placeholderRect(pkg, slidePath, "title");
  const body = await placeholderRect(pkg, slidePath, "body");
  return { ...(title ? { title } : {}), ...(body ? { body } : {}) };
}

/** The deck's slide size in EMU, from `<p:sldSz>`. */
export async function slideSize(pkg: Pkg): Promise<{ width: number; height: number }> {
  const pres = await pkg.doc("ppt/presentation.xml");
  const sz = child(pres.documentElement, P_NS, "sldSz");
  const width = Number(sz?.getAttribute("cx") ?? 0);
  const height = Number(sz?.getAttribute("cy") ?? 0);
  // A deck with no `<p:sldSz>` is not a deck PowerPoint wrote, and every
  // landing rule divides by these. The 4:3 default is what the format itself
  // falls back to, and it is better than an element positioned at NaN.
  if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0) {
    return { width: 9144000, height: 6858000 };
  }
  return { width, height };
}
