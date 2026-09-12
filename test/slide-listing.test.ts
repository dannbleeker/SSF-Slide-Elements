import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { keepOnly } from "../src/core/splice/listing.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * Leaving exactly one slide listed in a deck (`src/core/splice/listing.ts`).
 *
 * Three callers share it — the splice, `onlySlide` for an undo, and the removal
 * — and until now it was reached only through them, so its own refusal had
 * never run. That refusal is the difference between a nameable error and a
 * package handed to `insertSlidesFromBase64` with no slides in it, which is a
 * thing no round has measured the host's answer to.
 *
 * `test/listing.test.ts` is a different topic: that one is the store listing.
 */
describe("leaving one slide listed", () => {
  const three = (): Promise<Uint8Array> =>
    makeDeck([{ paragraphs: [["One"]] }, { paragraphs: [["Two"]] }, { paragraphs: [["Three"]] }]);

  it("unlists every other slide and says how many it dropped", async () => {
    const pkg = await Pkg.open(await three());
    expect(await keepOnly(pkg, "ppt/slides/slide2.xml")).toBe(2);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });

  it("leaves the parts and the relationships where they are", async () => {
    // The "unlisted" arm of probe question 1 rather than the "pruned" one: the
    // `<p:sldId>` entries go and nothing else does, which is the half that
    // touches least. Both landed a single slide on the web on 2026-09-10.
    const pkg = await Pkg.open(await three());
    await keepOnly(pkg, "ppt/slides/slide2.xml");
    expect(pkg.has("ppt/slides/slide1.xml")).toBe(true);
    expect(pkg.has("ppt/slides/slide3.xml")).toBe(true);
  });

  it("refuses, by name, a slide that is not in the deck's order", async () => {
    // The package would otherwise go to the host listing nothing at all. The
    // message names the part, because the caller that got here has a bug about
    // WHICH slide it asked for and the path is the only thing that says which.
    const pkg = await Pkg.open(await three());
    await expect(keepOnly(pkg, "ppt/slides/slide9.xml")).rejects.toThrow(/slide9\.xml is not in this deck/);
  });

  it("refuses a deck whose presentation.xml has no slide list", async () => {
    const pkg = await Pkg.open(await three());
    const pres = await pkg.doc("ppt/presentation.xml");
    const list = pres.getElementsByTagName("p:sldIdLst")[0];
    list?.parentNode?.removeChild(list);
    await expect(keepOnly(pkg, "ppt/slides/slide1.xml")).rejects.toThrow(/no <p:sldIdLst>/);
  });

  it("reads the relationship id in either spelling, and skips an entry carrying neither", async () => {
    // `getAttributeNS` is the correct read and `getAttribute("r:id")` is the
    // fallback, because a document parsed without the `r` namespace bound
    // answers only the second — and a `<p:sldId>` read as having no
    // relationship at all resolves to no target, which is not the slide being
    // kept, so it is unlisted. That is the right answer: an entry pointing at
    // nothing cannot be the one slide the package is for.
    const pkg = await Pkg.open(await three());
    const pres = await pkg.doc("ppt/presentation.xml");
    const ids = pres.getElementsByTagName("p:sldId");
    const first = ids[0] as Element;
    first.removeAttribute("r:id");
    // Two real entries left, and the id-less one goes with the other slide.
    expect(await keepOnly(pkg, "ppt/slides/slide2.xml")).toBe(2);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });

  it("drops nothing, and refuses nothing, on a deck of one slide", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Only"]] }]));
    expect(await keepOnly(pkg, "ppt/slides/slide1.xml")).toBe(0);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide1.xml"]);
  });
});
