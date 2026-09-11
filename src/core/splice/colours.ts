/**
 * "As in the library": every scheme colour pinned to the value the library's
 * own theme gives it.
 *
 * `docs/DESIGN.md` section 7 offers two positions, and only one of them costs
 * anything. "This deck's theme" is what happens by itself — a shape written
 * with `<a:schemeClr val="accent1"/>` resolves against whatever theme it lands
 * in, so an element dropped into a customer's deck comes out in the customer's
 * blue. "As in the library" is this pass: the same shape leaves here written
 * `<a:srgbClr val="5B9BD5"/>`, the value the library's theme resolved it to,
 * and it looks the same in every deck it is ever inserted into.
 *
 * What is deliberately NOT touched:
 *
 * - **An explicit colour.** `<a:srgbClr>` already says what it means, under
 *   either setting. Section 7 says so and this does nothing to them.
 * - **`phClr`.** The placeholder colour is not a theme slot; it is the colour
 *   whatever is applying a style is passing in, and it appears inside the
 *   theme's own format schemes rather than on a slide. There is nothing to
 *   resolve it to here, and a pin that resolved it to something would be
 *   inventing a colour. The library's markup uses it nowhere — measured on the
 *   committed catalogue, 0 of 8,000-odd scheme colours across the two sizes.
 * - **A name the theme has no colour for.** Left as a scheme colour, which
 *   still draws: it follows the destination's theme as it always did. Better a
 *   colour that moves than a colour invented here.
 *
 * The transforms stay. `<a:schemeClr val="tx1"><a:lumMod val="75000"/></…>` is
 * that colour lightened, and `<a:lumMod>` is as legal a child of `<a:srgbClr>`
 * as of `<a:schemeClr>` — both take the same `EG_ColorTransform` group — so the
 * children are carried across rather than dropped. Dropping them would pin the
 * colour and lose the shade, which is a worse answer than not pinning at all.
 *
 * Pure: a node in, a count out, and the node rewritten in place.
 */
import type { ThemeColours } from "../pptx/theme.js";
import { A_NS, parseXml, serializeXml } from "../pptx/xml.js";

/** What a pass over some markup did, in numbers a caller can report and a test can assert. */
export interface Pinned {
  /** Scheme colours rewritten to explicit values. */
  pinned: number;
  /** Scheme colours left alone: `phClr`, and any name this theme does not carry. */
  left: number;
}

/** The placeholder colour: a style's argument, not a theme slot. Never pinned. */
const PLACEHOLDER = "phClr";

/**
 * Rewrite every scheme colour under a node to the theme's explicit value.
 *
 * Depth-first over a snapshot of each parent's children, because the walk
 * replaces nodes as it goes and a live `childNodes` read mid-replacement skips
 * the sibling after every hit.
 */
export function pinSchemeColours(root: Element | Document, theme: ThemeColours): Pinned {
  const out: Pinned = { pinned: 0, left: 0 };
  const walk = (node: Element): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      if (el.namespaceURI === A_NS && el.localName === "schemeClr") {
        const value = el.getAttribute("val") ?? "";
        const colour = value === PLACEHOLDER ? undefined : theme[value];
        if (colour === undefined) {
          out.left += 1;
          walk(el);
          continue;
        }
        const doc = el.ownerDocument;
        if (!doc) continue;
        const name = el.prefix ? `${el.prefix}:srgbClr` : "srgbClr";
        const pinnedEl = doc.createElementNS(A_NS, name);
        pinnedEl.setAttribute("val", colour);
        // The transforms come across unchanged, in order: the same lumMod that
        // shaded the scheme colour shades the explicit one.
        for (const transform of Array.from(el.childNodes)) pinnedEl.appendChild(transform.cloneNode(true));
        node.replaceChild(pinnedEl, el);
        out.pinned += 1;
        continue;
      }
      walk(el);
    }
  };
  const start = "documentElement" in root ? root.documentElement : root;
  if (start) walk(start);
  return out;
}

/**
 * The same pass over a whole XML part, as text.
 *
 * For a CARRIED part — the library's one chart, which states 17 scheme colours
 * of its own. A pin that only rewrote the slide markup would leave the chart's
 * series following the destination's theme while the shapes around it did not,
 * which is a switch that does half of what it says.
 *
 * Answers the text unchanged when nothing was pinned, so a part that needed no
 * work is copied byte for byte rather than round-tripped through a serialiser.
 */
export function pinColoursInXml(xml: string, theme: ThemeColours): { xml: string; result: Pinned } {
  if (!xml.includes("schemeClr")) return { xml, result: { pinned: 0, left: 0 } };
  const doc = parseXml(xml);
  const result = pinSchemeColours(doc, theme);
  if (result.pinned === 0) return { xml, result };
  return { xml: serializeXml(doc), result };
}
