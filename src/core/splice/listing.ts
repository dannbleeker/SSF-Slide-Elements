/**
 * Leaving exactly one slide listed in a deck.
 *
 * Its own file because three callers share it and none of them owns it: the
 * splice hands back the user's deck with one REBUILT slide listed, `onlySlide`
 * hands back the same deck with one of the user's OWN slides listed for an
 * undo, and the removal hands back a slide with an element taken off it. A
 * shared step that lives inside one of its callers is a step the other two
 * cannot use without importing that caller — which is how a test that stubbed
 * the splice broke the removal.
 */
import { Pkg } from "../pptx/pkg.js";
import { P_NS, R_NS, element, elements } from "../pptx/xml.js";

const PRESENTATION = "ppt/presentation.xml";

/**
 * Leave exactly one slide listed in the deck's own order.
 *
 * The `<p:sldId>` entries go and the relationships stay, which is the "unlisted"
 * arm of probe question 1 rather than the "pruned" one — both landed a single
 * slide on the web, and this is the half that touches least.
 */
export async function keepOnly(pkg: Pkg, slidePath: string): Promise<number> {
  const pres = await pkg.doc(PRESENTATION);
  const list = element(pres, P_NS, "sldIdLst");
  if (!list) throw new Error("ssf-slide-elements: this deck's presentation.xml has no <p:sldIdLst>");
  let removed = 0;
  for (const sldId of elements(list, P_NS, "sldId")) {
    const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
    const target = rId ? await pkg.relTarget(PRESENTATION, rId) : undefined;
    if (target === slidePath) continue;
    sldId.parentNode?.removeChild(sldId);
    removed += 1;
  }
  // The one slide that survives is the one the caller named. A package listing
  // nothing at all would be handed to `insertSlidesFromBase64` as a deck with
  // no slides in it, and the host's answer to that is not something any round
  // has measured — so it is refused here, where the cause is still nameable.
  if (!element(pres, P_NS, "sldIdLst")?.getElementsByTagNameNS(P_NS, "sldId").length) {
    throw new Error(`ssf-slide-elements: ${slidePath} is not in this deck's slide order, so nothing would be inserted`);
  }
  return removed;
}
