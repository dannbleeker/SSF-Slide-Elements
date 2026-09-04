/**
 * A thumbnail for the picker, drawn from the element's own geometry.
 *
 * Deliberately not a rendering. Rasterising a slide properly means LibreOffice
 * or PowerPoint itself: a build dependency that has to be installed in CI, that
 * produces a binary nobody can review in a diff, and that turns a one-line
 * change to the library into a pile of changed PNGs. What a picker actually
 * needs is a shape the eye can match against the thing it is looking for — is
 * this the three-box one or the four-box one, is the emphasis left or right —
 * and that is geometry, which is already in the file.
 *
 * So this draws boxes where the boxes are, in the colours the file states, with
 * a hint of where the text sits. It is honest about being a sketch: no
 * gradients, no shadows, no fonts.
 *
 * The one thing it must get RIGHT is the group coordinate transform, because
 * roughly half the library's elements are groups and a preview that ignores it
 * draws every group's contents in the wrong place — usually piled in the
 * top-left corner, which looks like a rendering bug rather than a maths one.
 */
import { A_NS, P_NS, child, childElements, elements } from "../pptx/xml.js";
import type { Box } from "../catalogue/types.js";

/** How a child coordinate becomes a slide coordinate inside a group. */
interface Frame {
  /** Slide-space origin the child space maps onto. */
  x: number;
  y: number;
  /** Child-space origin, subtracted before scaling. */
  chX: number;
  chY: number;
  sx: number;
  sy: number;
}

const IDENTITY: Frame = { x: 0, y: 0, chX: 0, chY: 0, sx: 1, sy: 1 };

function apply(f: Frame, x: number, y: number): { x: number; y: number } {
  return { x: f.x + (x - f.chX) * f.sx, y: f.y + (y - f.chY) * f.sy };
}

function xfrmOf(shape: Element): Element | undefined {
  const holder = child(shape, P_NS, "spPr") ?? child(shape, P_NS, "grpSpPr") ?? child(shape, P_NS, "xfrm");
  if (!holder) return undefined;
  return holder.localName === "xfrm" ? holder : child(holder, A_NS, "xfrm");
}

function boxOf(shape: Element): Box | undefined {
  const xfrm = xfrmOf(shape);
  if (!xfrm) return undefined;
  const off = child(xfrm, A_NS, "off");
  const ext = child(xfrm, A_NS, "ext");
  const x = Number(off?.getAttribute("x"));
  const y = Number(off?.getAttribute("y"));
  const cx = Number(ext?.getAttribute("cx"));
  const cy = Number(ext?.getAttribute("cy"));
  if (![x, y, cx, cy].every(Number.isFinite)) return undefined;
  return { x, y, cx, cy };
}

/**
 * The frame a group establishes for its children.
 *
 * `chOff`/`chExt` are the child coordinate space; `off`/`ext` are where that
 * space lands on the slide. They are routinely different by a factor, which is
 * what makes a group scalable, and are occasionally absent — a group with no
 * stated child space uses its own, which is the identity case.
 */
function frameOf(group: Element, outer: Frame): Frame {
  const box = boxOf(group);
  const xfrm = xfrmOf(group);
  if (!box || !xfrm) return outer;
  const chOff = child(xfrm, A_NS, "chOff");
  const chExt = child(xfrm, A_NS, "chExt");
  const chX = Number(chOff?.getAttribute("x"));
  const chY = Number(chOff?.getAttribute("y"));
  const chCx = Number(chExt?.getAttribute("cx"));
  const chCy = Number(chExt?.getAttribute("cy"));
  const placed = apply(outer, box.x, box.y);
  if (![chX, chY, chCx, chCy].every(Number.isFinite) || chCx === 0 || chCy === 0) {
    return { x: placed.x, y: placed.y, chX: box.x, chY: box.y, sx: outer.sx, sy: outer.sy };
  }
  return {
    x: placed.x,
    y: placed.y,
    chX,
    chY,
    sx: (box.cx / chCx) * outer.sx,
    sy: (box.cy / chCy) * outer.sy,
  };
}

/**
 * The first explicit colour a shape states for its fill.
 *
 * `srgbClr` only. A `schemeClr` resolves through the theme, and resolving a
 * theme properly means the colour map, the master, the layout and any override
 * on the way — a lot of code for a thumbnail, and a WRONG answer from a partial
 * implementation is worse than an honest neutral, because it looks deliberate.
 * Scheme colours fall through to the neutral below and the sketch reads as a
 * sketch.
 */
function fillOf(shape: Element): string | undefined {
  const spPr = child(shape, P_NS, "spPr");
  if (!spPr) return undefined;
  if (child(spPr, A_NS, "noFill")) return "none";
  const solid = child(spPr, A_NS, "solidFill");
  if (!solid) return undefined;
  const srgb = child(solid, A_NS, "srgbClr");
  const val = srgb?.getAttribute("val");
  return val && /^[0-9a-fA-F]{6}$/.test(val) ? `#${val}` : undefined;
}

function hasText(shape: Element): boolean {
  return elements(shape, A_NS, "t").some((t) => (t.textContent ?? "").trim().length > 0);
}

interface Mark {
  box: Box;
  fill: string;
  text: boolean;
  round: boolean;
}

const SHAPE_TAGS = new Set(["sp", "grpSp", "graphicFrame", "pic", "cxnSp"]);

/** Walk a shape tree into flat, slide-space marks. Groups recurse; leaves draw. */
function collect(nodes: Element[], frame: Frame, out: Mark[], depth = 0): void {
  // A library element is a handful of shapes a few levels deep. The cap is for
  // a pathological file rather than for anything in the shipped library, and it
  // stops a cyclic or absurdly nested document from hanging the build.
  if (depth > 12) return;
  for (const node of nodes) {
    const local = node.localName ?? "";
    if (!SHAPE_TAGS.has(local)) continue;
    if (local === "grpSp") {
      collect(childElements(node), frameOf(node, frame), out, depth + 1);
      continue;
    }
    const box = boxOf(node);
    if (!box) continue;
    const tl = apply(frame, box.x, box.y);
    const geom = child(child(node, P_NS, "spPr") ?? node, A_NS, "prstGeom");
    const preset = geom?.getAttribute("prst") ?? "";
    out.push({
      box: { x: tl.x, y: tl.y, cx: box.cx * frame.sx, cy: box.cy * frame.sy },
      fill: fillOf(node) ?? (local === "pic" ? "#c8ccd4" : "#e6e9ef"),
      text: hasText(node),
      round: preset.startsWith("ellipse") || preset.startsWith("round"),
    });
  }
}

/**
 * An inline SVG sketch of an element, scaled to fit a box `w` by `h`.
 *
 * Inline, and stored in the index rather than as a file per element: a hundred
 * separate thumbnail requests when the pane opens is a hundred round trips, and
 * these are a few hundred bytes each.
 */
export function previewSvg(shapes: Element[], bounds: Box, w = 240, h = 135): string {
  const marks: Mark[] = [];
  collect(shapes, IDENTITY, marks);
  if (marks.length === 0 || bounds.cx <= 0 || bounds.cy <= 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" role="img"></svg>`;
  }
  // Fit the element's own bounds, not the slide's: the picker shows the
  // element, and an element occupying a third of a slide would otherwise be a
  // third of a thumbnail with nothing around it.
  const k = Math.min(w / bounds.cx, h / bounds.cy);
  const ox = (w - bounds.cx * k) / 2;
  const oy = (h - bounds.cy * k) / 2;
  const px = (v: number): string => (Math.round(v * 100) / 100).toString();

  const body = marks
    .map((m) => {
      const x = ox + (m.box.x - bounds.x) * k;
      const y = oy + (m.box.y - bounds.y) * k;
      const cx = Math.max(1, m.box.cx * k);
      const cy = Math.max(1, m.box.cy * k);
      const rx = m.round ? Math.min(cx, cy) / 2 : 1.5;
      const rect = `<rect x="${px(x)}" y="${px(y)}" width="${px(cx)}" height="${px(cy)}" rx="${px(rx)}" fill="${m.fill}"/>`;
      if (!m.text || cy < 8) return rect;
      // Two bars where the text is. Not the text itself: at thumbnail size a
      // real string is a grey smear, and a wrong-font smear at that.
      const bx = px(x + cx * 0.12);
      const bw = px(cx * 0.5);
      const bw2 = px(cx * 0.66);
      return (
        rect +
        `<rect x="${bx}" y="${px(y + cy * 0.3)}" width="${bw2}" height="2" rx="1" fill="#9aa3b0"/>` +
        `<rect x="${bx}" y="${px(y + cy * 0.3 + 4)}" width="${bw}" height="2" rx="1" fill="#bcc3ce"/>`
      );
    })
    .join("");

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" role="img">${body}</svg>`;
}
