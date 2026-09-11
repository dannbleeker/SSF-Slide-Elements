import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { cloneSlide, creationIdOf, notesPathFor, setCreationId } from "../src/core/pptx/clone.js";
import { P_NS, R_NS, child, element, elements, parseXml } from "../src/core/pptx/xml.js";
import { COMMENT_REL_TYPES, REL_TYPE } from "../src/core/pptx/parts.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

/**
 * Duplicating a slide inside the package, on bytes the fixture built.
 *
 * `clone.ts` was ported from SSF-Merge with its comments intact, and each of
 * those comments names a deck PowerPoint refused to open or a merged deck that
 * opened and was quietly wrong. The tests below are shaped as one case per
 * named incident, because that is the only shape in which they can go red for
 * the reason they were written: a clone that is structurally plausible and
 * semantically wrong — two slides sharing a notes page, a copy stamped with an
 * id the deck already holds, a vendor's shape pointing at our metadata — is
 * precisely what a coarser test passes.
 *
 * Three habits run through the file, and they are all about vacuity.
 *
 * The negative assertions are paired with a positive one that the thing being
 * ruled out was really there to begin with: a test that a stray creation id is
 * ignored proves nothing on a fixture that never had one.
 *
 * The reads go through `pkg.text` and are parsed again rather than through
 * `pkg.doc`, so what is asserted is what the package would SHIP. The duplicate
 * `xmlns:p14` incident lives in the serialiser, and a live document hides it.
 *
 * Where a case needs a package shape `makeDeck` has no option for — a slide
 * with its own tag reference, a comment part, a notes page under a
 * percent-encoded name — the fixture bytes are edited through JSZip, the way
 * `test/pptx.test.ts` does it, rather than by inventing a second deck builder.
 */

const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const CREATION_ID_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";
const MODERN_COMMENT = COMMENT_REL_TYPES[1] ?? "";

async function deck(...args: Parameters<typeof makeDeck>): Promise<Pkg> {
  return Pkg.open(await makeDeck(...args));
}

/** The fixture's bytes, with the zip edited before the package is opened. */
async function editedDeck(slides: SlideSpec[], edit: (zip: JSZip) => Promise<void> | void): Promise<Pkg> {
  const zip = await JSZip.loadAsync(await makeDeck(slides));
  await edit(zip);
  return Pkg.open(await zip.generateAsync({ type: "uint8array" }));
}

/** Append markup to a part already in the zip, by splicing it in front of a closing tag. */
async function spliceInto(zip: JSZip, path: string, before: string, markup: string): Promise<void> {
  const file = zip.file(path);
  if (!file) throw new Error(`the fixture has no ${path}, so this test proves nothing`);
  const text = await file.async("string");
  if (!text.includes(before)) throw new Error(`${path} no longer contains ${before}`);
  zip.file(path, text.replace(before, `${markup}${before}`));
}

/** A part as the package would write it out, parsed again. Never the live document. */
async function shipped(pkg: Pkg, path: string): Promise<Document> {
  return parseXml(await pkg.text(path));
}

/** Every `Target` in a part's relationships, by relationship type. */
async function targetsByType(pkg: Pkg, part: string): Promise<Map<string, string[]>> {
  const rels = await shipped(pkg, Pkg.relsPathFor(part));
  const out = new Map<string, string[]>();
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    const type = rel.getAttribute("Type") ?? "";
    out.set(type, [...(out.get(type) ?? []), rel.getAttribute("Target") ?? ""]);
  }
  return out;
}

/** The `val` of every `<p14:creationId>` anywhere in a part, in document order. */
async function everyCreationId(pkg: Pkg, path: string): Promise<string[]> {
  const doc = await shipped(pkg, path);
  return elements(doc, P14_NS, "creationId").map((el) => el.getAttribute("val") ?? "");
}

/** A generator that hands out the values given, then repeats the last one for ever. */
function draws(...values: number[]): { next: () => number; count: () => number } {
  let at = 0;
  return {
    next: () => values[Math.min(at++, values.length - 1)] ?? 0,
    count: () => at,
  };
}

describe("a cloned slide is a slide the deck agrees is there", () => {
  it("adds the part, its relationships, the override, the presentation relationship and the id list entry", async () => {
    /**
     * The six things `clone.ts` opens by naming. PowerPoint's answer to any one
     * of them missing is the same — it declines to open the file and does not
     * say which — so all six are asserted together, in one place, rather than
     * left to whichever later test happens to notice.
     *
     * `slidePaths` is what proves the last two AGREE rather than merely both
     * exist: it walks the id list, resolves each `r:id` through the
     * presentation's relationships, and answers with part paths. A `<p:sldId>`
     * naming a relationship that points somewhere else drops out of that list.
     */
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    expect(target).toBe("ppt/slides/slide2.xml");
    expect(pkg.has(target), "the part itself").toBe(true);
    expect(pkg.has("ppt/slides/_rels/slide2.xml.rels"), "the copy's own relationships").toBe(true);
    expect(await pkg.text("[Content_Types].xml")).toContain('PartName="/ppt/slides/slide2.xml"');
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);

    // And the id list entry is a NEW id rather than a second use of the
    // template's, which is the half `slidePaths` cannot see: it resolves by
    // `r:id`, so two entries sharing one `id` attribute would still answer with
    // two paths.
    const pres = await shipped(pkg, "ppt/presentation.xml");
    const ids = elements(pres, P_NS, "sldId").map((el) => el.getAttribute("id"));
    expect(new Set(ids).size, "two slides were given one <p:sldId> id").toBe(ids.length);
    const added = elements(pres, P_NS, "sldId")[1];
    expect(added?.getAttributeNS(R_NS, "id") ?? added?.getAttribute("r:id")).toMatch(/^rId\d+$/);
  });

  it("copies the source's relationships and leaves the source pointing where it did", async () => {
    // The layout is deliberately SHARED — it is read-only as far as a clone is
    // concerned and duplicating it would bloat the deck for nothing — so the
    // copy naming the same layout is the contract, not an oversight.
    const pkg = await deck([{ paragraphs: [["Hello"]], chart: "Sales" }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    const copy = await targetsByType(pkg, target);
    const source = await targetsByType(pkg, "ppt/slides/slide1.xml");
    expect(copy.get(REL_TYPE.slideLayout)).toEqual(["../slideLayouts/slideLayout1.xml"]);
    expect(copy.get(REL_TYPE.chart), "the chart is shared until the splice says otherwise").toEqual([
      "../charts/chart1.xml",
    ]);
    expect(source.get(REL_TYPE.slideLayout)).toEqual(["../slideLayouts/slideLayout1.xml"]);
    expect(source.get(REL_TYPE.chart)).toEqual(["../charts/chart1.xml"]);
  });

  it("clones a slide that has no relationships part at all", async () => {
    /**
     * A slide with no `.rels` is legal — it relates to nothing, not even a
     * layout — and every one of the three passes below reaches for that part.
     * The fixture always writes one, so this branch is only reachable by taking
     * it away, and it is the branch where an unguarded `pkg.doc` throws and
     * takes the whole clone with it.
     */
    const pkg = await editedDeck([{ paragraphs: [["Hello"]] }], (zip) => {
      zip.remove("ppt/slides/_rels/slide1.xml.rels");
    });
    expect(pkg.has("ppt/slides/_rels/slide1.xml.rels"), "the fixture still wrote one").toBe(false);

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect(pkg.has(target)).toBe(true);
    expect(pkg.has(Pkg.relsPathFor(target)), "a copy of nothing is still nothing").toBe(false);
    expect(await creationIdOf(pkg, target)).toBe(222);
  });
});

describe("the creation id a copy is given", () => {
  /**
   * Office.js reports a slide as `256#3561048925` — the `<p:sldId id>` joined
   * to the slide's own `<p14:creationId val>` — so re-inserting a slide
   * carrying an id the deck already holds asks the host to hold one identity
   * twice. office-js#6105 reports exactly that failing with `InvalidArgument`
   * on Windows desktop.
   */
  it("gives the copy an id, drawn without help, that the deck was not already using", async () => {
    /**
     * No generator injected, so the value comes from the module's own draw.
     * This is the case the previous version of this suite could not reach: it
     * injected a counter and asserted the counter's own uniqueness, so the
     * generator a real run uses was never in an assertion at all.
     */
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml");

    const id = await creationIdOf(pkg, target);
    expect(id, "a copy with no creation id is a copy PowerPoint invents one for").toBeDefined();
    expect(Number.isInteger(id)).toBe(true);
    // The format's range, and not the template's value.
    expect(id).toBeGreaterThan(0);
    expect(id).toBeLessThanOrEqual(0xffff_ffff);
    expect(id).not.toBe(111);
  });

  it("redraws a candidate the deck is already using rather than honouring it", async () => {
    // The freshness rule applies to an INJECTED generator too, which is what
    // makes it checkable at all: hand over a value the deck already holds and
    // watch it be refused.
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const generator = draws(111, 222);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: generator.next });

    expect(await creationIdOf(pkg, target)).toBe(222);
    expect(generator.count(), "the colliding first answer was taken").toBe(2);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml"), "the template kept its own").toBe(111);
  });

  it("stops redrawing rather than spinning on a generator that never moves", async () => {
    /**
     * Bounded rather than looped, deliberately. A generator that keeps
     * answering the same number is a caller being deliberate, and spinning for
     * ever on it would be worse than honouring it.
     *
     * The upper bound is loose on purpose: the property is "this terminates",
     * not "this draws exactly eight times", and pinning the constant would make
     * an ordinary tuning change look like a regression. The lower bound is the
     * half that carries the meaning — it refused at least once before giving in.
     */
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const generator = draws(111);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: generator.next });

    expect(generator.count()).toBeGreaterThan(1);
    expect(generator.count(), "a bounded redraw, not a loop").toBeLessThan(64);
    expect(await creationIdOf(pkg, target), "honoured rather than spun on").toBe(111);
  });

  it("does not hand a second copy the id the first one took", async () => {
    /**
     * The collision that is neither the template's nor the user's: an earlier
     * copy made in the SAME run. The set of ids in use is gathered once per
     * package and each drawn id is added to it, so a second clone can see what
     * the first one took — nothing re-reads the slides.
     */
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const generator = draws(900, 900, 901);
    const first = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: generator.next });
    const second = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: generator.next });

    expect(await creationIdOf(pkg, first)).toBe(900);
    expect(await creationIdOf(pkg, second), "the second copy reused the first copy's id").toBe(901);
  });

  it("gives a slide that carried no creation id one of its own", async () => {
    /**
     * A slide with no creation id at all is legal and PowerPoint invents one on
     * open — which would make two copies indistinguishable at exactly the wrong
     * moment. The extension slot is written rather than found.
     *
     * The `xmlns:p14` count is the second half, and it is a serialiser
     * question, not a DOM one. Declaring the prefix by hand beside
     * `createElementNS` produced `<p14:creationId xmlns:p14="…" val="…"
     * xmlns:p14="…"/>` — a duplicate attribute, which XML forbids outright
     * (WFC: Unique Att Spec) and PowerPoint rejects the whole package for,
     * saying nothing about which part.
     */
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml"), "the fixture stamped one after all").toBeUndefined();

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect(await creationIdOf(pkg, target)).toBe(222);

    const xml = await pkg.text(target);
    expect(xml).toContain(CREATION_ID_URI);
    expect(xml.match(/xmlns:p14=/g) ?? [], "a duplicate attribute is not well-formed XML").toHaveLength(1);
  });

  it("rewrites an existing id in place rather than adding a second one", async () => {
    // Two `<p14:creationId>` elements in one slide is a deck whose identity
    // depends on which one a reader happens to reach first.
    const pkg = await deck([{ paragraphs: [["Hello"]], creationId: 111 }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    expect(await everyCreationId(pkg, target)).toEqual(["222"]);
    expect(await everyCreationId(pkg, "ppt/slides/slide1.xml")).toEqual(["111"]);
  });

  it("adds itself to an extension list <p:cSld> already had, rather than starting a second one", async () => {
    /**
     * `<p:cSld>` takes at most ONE `<p:extLst>`, so a slide that already
     * carries an extension — and any slide PowerPoint has drawn a guide on
     * does — must have the creation id appended to that list. A second list
     * beside it is schema-invalid, and PowerPoint's answer to schema-invalid is
     * to decline the file without naming the part.
     *
     * The extension spliced in below is a stand-in; which one it is decides
     * nothing. That the list is already there is the whole case.
     */
    const OTHER = `<p:ext uri="{00000000-0000-0000-0000-00000000CAFE}"><p14:other xmlns:p14="${P14_NS}"/></p:ext>`;
    const pkg = await editedDeck([{ paragraphs: [["Hello"]] }], (zip) =>
      spliceInto(zip, "ppt/slides/slide1.xml", "</p:cSld>", `<p:extLst>${OTHER}</p:extLst>`),
    );

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    const doc = await shipped(pkg, target);
    const cSld = element(doc, P_NS, "cSld")!;
    const lists = elements(cSld, P_NS, "extLst").filter((el) => el.parentNode === cSld);
    expect(lists, "two <p:extLst> children is a deck PowerPoint declines").toHaveLength(1);
    expect(elements(lists[0]!, P_NS, "ext")).toHaveLength(2);
    expect(await creationIdOf(pkg, target)).toBe(222);
  });

  it("refuses a part that has no <p:cSld> instead of stamping nothing", async () => {
    // Silence here is the expensive answer: a part that took the stamp and did
    // not keep it reads, to `creationIdOf`, exactly like a slide that has none.
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    await expect(setCreationId(pkg, "ppt/theme/theme1.xml", 222)).rejects.toThrow(/has no <p:cSld>/);
  });
});

describe("a creation id hiding in the shape tree", () => {
  /**
   * An older version of this code appended the id to whatever `<p:extLst>` it
   * found first, and `element` walks DESCENDANTS — so a slide whose shape tree
   * ends in its own extension list had the id put THERE, where PowerPoint does
   * not look. It invented one on open and two copies were indistinguishable.
   *
   * Which makes such a deck an ordinary input rather than a hypothetical: a
   * deck merged by that version carries one, and using a merged deck as a
   * template is a normal thing to do. The write and the read must BOTH be
   * scoped to `<p:cSld>`'s own children, and they failed from opposite sides —
   * the write put the id somewhere PowerPoint ignores, the read found a stray
   * one and reported that the stamp had worked.
   */
  const STRAY =
    `<p:extLst><p:ext uri="${CREATION_ID_URI}">` +
    `<p14:creationId xmlns:p14="${P14_NS}" val="424242"/>` +
    `</p:ext></p:extLst>`;

  it("is not what the slide's own creation id means", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]], shapes: [STRAY] }]);
    // The stray really is in the part, or the assertion below is about nothing.
    expect(await pkg.text("ppt/slides/slide1.xml")).toContain("424242");
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBeUndefined();
  });

  it("is left alone, and the copy's own id is written as a direct child of <p:cSld>", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]], shapes: [STRAY] }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    const doc = await shipped(pkg, target);
    const cSld = element(doc, P_NS, "cSld");
    const own = cSld ? child(cSld, P_NS, "extLst") : undefined;
    expect(own, "the id went somewhere other than <p:cSld>'s own children").toBeTruthy();
    expect(elements(own!, P14_NS, "creationId")[0]?.getAttribute("val")).toBe("222");
    // And the shape tree's own list still holds what it held. Updating THAT is
    // the other half of the same defect: a diagnostic agreeing with a stamp
    // that never landed.
    expect(await everyCreationId(pkg, target)).toEqual(["424242", "222"]);
    expect(await creationIdOf(pkg, target)).toBe(222);
  });

  it("reads a creation id with no value at all as no creation id", async () => {
    // `Number(null)` is 0, and 0 is a creation id — so an id-less element read
    // without this guard puts a value into the used set that no slide has, and
    // reports one for a slide that has none.
    const pkg = await editedDeck([{ paragraphs: [["Hello"]] }], (zip) =>
      spliceInto(
        zip,
        "ppt/slides/slide1.xml",
        "</p:cSld>",
        `<p:extLst><p:ext uri="${CREATION_ID_URI}"><p14:creationId xmlns:p14="${P14_NS}"/></p:ext></p:extLst>`,
      ),
    );
    expect(await pkg.text("ppt/slides/slide1.xml")).toContain("creationId");
    expect(await creationIdOf(pkg, "ppt/slides/slide1.xml")).toBeUndefined();
  });

  it("reads no creation id from a part that is not a slide, rather than throwing", async () => {
    // The write refuses such a part outright, because a stamp that does not
    // land is worse than an error. The READ is the pane's diagnostic and is
    // asked about whatever it is pointed at, so its answer is "no id" — an
    // exception out of a diagnostic takes down the thing it was reporting on.
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    expect(await creationIdOf(pkg, "ppt/theme/theme1.xml")).toBeUndefined();
  });
});

describe("the copy's notes page", () => {
  /**
   * The one part a clone does NOT share. A notes page is per-slide content, so
   * a shared one means two slides editing the same notes — and PowerPoint then
   * shows one slide's notes on another, because the notes page names the slide
   * it belongs to and the copy inherited a name pointing at the template.
   */
  it("is the copy's own, with its back-reference repointed at the copy", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]], notes: true }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    const copy = await targetsByType(pkg, target);
    expect(copy.get(REL_TYPE.notesSlide)).toEqual(["../notesSlides/notesSlide2.xml"]);
    expect(pkg.has("ppt/notesSlides/notesSlide2.xml")).toBe(true);
    expect(await pkg.text("[Content_Types].xml")).toContain('PartName="/ppt/notesSlides/notesSlide2.xml"');

    // The back-reference, which is the half that decides whose notes the user
    // is shown.
    const back = await targetsByType(pkg, "ppt/notesSlides/notesSlide2.xml");
    expect(back.get(REL_TYPE.slide)).toEqual(["../slides/slide2.xml"]);

    // The template is untouched on both sides of the pair.
    const source = await targetsByType(pkg, "ppt/slides/slide1.xml");
    expect(source.get(REL_TYPE.notesSlide)).toEqual(["../notesSlides/notesSlide1.xml"]);
    expect((await targetsByType(pkg, "ppt/notesSlides/notesSlide1.xml")).get(REL_TYPE.slide)).toEqual([
      "../slides/slide1.xml",
    ]);
  });

  it("is numbered from the notes parts, never from the slide", async () => {
    /**
     * The two sequences are independent — part names are arbitrary and drift
     * apart the moment a slide is deleted — so naming a clone's notes after the
     * SLIDE number lands on a part that is already there. The deck below is
     * built to make that land: the copy is `slide2.xml`, and a
     * `notesSlide2.xml` is already in the package.
     *
     * Nothing about the failure is loud. `copyPart` overwrites silently and
     * `addContentTypeOverride` no-ops on an override that exists, so the
     * package stays structurally valid while the clone shares the template's
     * notes page and the decoy's owner loses its content.
     */
    const pkg = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], (zip) => {
      zip.file(
        "ppt/notesSlides/notesSlide2.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
          `<p:notes xmlns:p="${P_NS}"><p:cSld><p:spTree/></p:cSld></p:notes>`,
      );
    });

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect(target, "the copy is the slide number that collides").toBe("ppt/slides/slide2.xml");
    expect((await targetsByType(pkg, target)).get(REL_TYPE.notesSlide)).toEqual(["../notesSlides/notesSlide3.xml"]);
    expect(await pkg.text("ppt/notesSlides/notesSlide2.xml"), "the decoy was overwritten").toContain("<p:spTree/>");
  });

  it("is found under a percent-encoded name, because the package's own resolver is asked", async () => {
    /**
     * OPC maps a part name to a zip item by stripping the leading `/` and
     * nothing else, so a notes page stored as `notes%20slide1.xml` is stored
     * under exactly that name. A resolver that percent-DECODES answers a path
     * `has` says no to — and this function reads that as "no notes page", so
     * every copy kept the template's own `Target`.
     *
     * Which was not the end of it: `removeSlide` DOES ask the package, so it
     * found the part and deleted it on the way out, leaving the copies pointing
     * at nothing. Two dangling relationships and no content type, which is a
     * deck PowerPoint opens as "repaired".
     */
    const encoded = "ppt/notesSlides/notes%20slide1.xml";
    const pkg = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], async (zip) => {
      const notes = await zip.file("ppt/notesSlides/notesSlide1.xml")!.async("string");
      const rels = await zip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels")!.async("string");
      zip.remove("ppt/notesSlides/notesSlide1.xml");
      zip.remove("ppt/notesSlides/_rels/notesSlide1.xml.rels");
      zip.file(encoded, notes);
      zip.file("ppt/notesSlides/_rels/notes%20slide1.xml.rels", rels);
      const slideRels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
      zip.file(
        "ppt/slides/_rels/slide1.xml.rels",
        slideRels.replace("../notesSlides/notesSlide1.xml", "../notesSlides/notes%20slide1.xml"),
      );
    });
    expect(pkg.has(encoded), "the package holds the name as written").toBe(true);

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    const copy = await targetsByType(pkg, target);
    expect(copy.get(REL_TYPE.notesSlide), "the copy kept the template's own notes page").toEqual([
      "../notesSlides/notesSlide1.xml",
    ]);
    expect((await targetsByType(pkg, "ppt/notesSlides/notesSlide1.xml")).get(REL_TYPE.slide)).toEqual([
      "../slides/slide2.xml",
    ]);
    // The template still names the part it always named.
    expect((await targetsByType(pkg, "ppt/slides/slide1.xml")).get(REL_TYPE.notesSlide)).toEqual([
      "../notesSlides/notes%20slide1.xml",
    ]);
  });

  it("is not invented for a slide that had none", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    expect((await targetsByType(pkg, target)).get(REL_TYPE.notesSlide)).toBeUndefined();
    expect(pkg.partNames().filter((p) => p.startsWith("ppt/notesSlides/"))).toEqual([]);
  });

  it("is not copied when the relationship names a part the package does not hold", async () => {
    // A dangling notes relationship is a deck already in trouble, and the clone
    // is not the place to find out: copying a part that is not there throws,
    // and takes the whole insert with it.
    const pkg = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], (zip) => {
      zip.remove("ppt/notesSlides/notesSlide1.xml");
      zip.remove("ppt/notesSlides/_rels/notesSlide1.xml.rels");
    });

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect(await creationIdOf(pkg, target)).toBe(222);
    expect(pkg.partNames().filter((p) => p.startsWith("ppt/notesSlides/"))).toEqual([]);
  });

  it("keeps the notes page's other relationships pointing where they did", async () => {
    /**
     * A real notes page relates to its notes MASTER as well as to its slide,
     * and a page with a picture on it relates to that too. Only the
     * back-reference is the copy's business; a repoint that rewrote every
     * target in the part would leave the copy's notes page claiming the notes
     * master lives at `../slides/slide2.xml`.
     */
    const NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
    const pkg = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], (zip) =>
      spliceInto(
        zip,
        "ppt/notesSlides/_rels/notesSlide1.xml.rels",
        "</Relationships>",
        `<Relationship Id="rId2" Type="${NOTES_MASTER}" Target="../notesMasters/notesMaster1.xml"/>`,
      ),
    );

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect((await targetsByType(pkg, target)).get(REL_TYPE.notesSlide)).toEqual(["../notesSlides/notesSlide2.xml"]);
    const back = await targetsByType(pkg, "ppt/notesSlides/notesSlide2.xml");
    expect(back.get(REL_TYPE.slide)).toEqual(["../slides/slide2.xml"]);
    expect(back.get(NOTES_MASTER)).toEqual(["../notesMasters/notesMaster1.xml"]);
  });

  it("keeps its own relationships when the template's notes page had none", async () => {
    // The notes page's `.rels` is what carries the back-reference, and a notes
    // part without one is legal. The repoint pass has to be skipped rather than
    // reach for a part that is not there.
    const pkg = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], (zip) => {
      zip.remove("ppt/notesSlides/_rels/notesSlide1.xml.rels");
    });

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect((await targetsByType(pkg, target)).get(REL_TYPE.notesSlide)).toEqual(["../notesSlides/notesSlide2.xml"]);
    expect(pkg.has("ppt/notesSlides/_rels/notesSlide2.xml.rels")).toBe(false);
  });
});

describe("what a copy does not inherit", () => {
  /**
   * The `.rels` are copied verbatim, which is right for a layout and wrong for
   * anything that is ABOUT the template rather than produced by it.
   *
   * The decks below carry both kinds of tag reference at once, because the
   * whole difficulty is telling them apart: the slide's own reference lives in
   * `<p:cSld><p:custDataLst>` and must go, while a reference from inside a
   * SHAPE is another add-in's bookkeeping and must not.
   */
  const SLIDE_TAG_REL = `<Relationship Id="rId31" Type="${REL_TYPE.tags}" Target="../tags/tag8.xml"/>`;
  const SLIDE_TAGS = `<p:custDataLst><p:tags r:id="rId31"/></p:custDataLst>`;

  async function taggedDeck(extraCustData = ""): Promise<Pkg> {
    return editedDeck([{ paragraphs: [["Hello"]], shapeTags: true }], async (zip) => {
      zip.file(
        "ppt/tags/tag8.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
          `<p:tagLst xmlns:p="${P_NS}"><p:tag name="OURS" val="slide level"/></p:tagLst>`,
      );
      await spliceInto(
        zip,
        "ppt/slides/slide1.xml",
        "</p:cSld>",
        SLIDE_TAGS.replace("</p:custDataLst>", extraCustData + "</p:custDataLst>"),
      );
      await spliceInto(zip, "ppt/slides/_rels/slide1.xml.rels", "</Relationships>", SLIDE_TAG_REL);
      await spliceInto(
        zip,
        "[Content_Types].xml",
        "</Types>",
        `<Override PartName="/ppt/tags/tag8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`,
      );
    });
  }

  it("takes the slide's own tag reference, and the relationship that named it, off the copy", async () => {
    /**
     * A template that already carries a tag part hands the clone a relationship
     * pointing at the TEMPLATE's `ppt/tags/tagN.xml`, and a tag writer then
     * appends the copy's metadata there — because from its side the slide
     * plainly has a tag reference. Every copy ends up sharing one part, so all
     * but the last are overwritten, and the user's own template is stamped as
     * this add-in's output.
     */
    const pkg = await taggedDeck();
    const source = await shipped(pkg, "ppt/slides/slide1.xml");
    const sourceCSld = element(source, P_NS, "cSld");
    expect(
      sourceCSld ? child(sourceCSld, P_NS, "custDataLst") : undefined,
      "the fixture edit did not land",
    ).toBeTruthy();

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    const doc = await shipped(pkg, target);
    const cSld = element(doc, P_NS, "cSld");
    // The empty list goes too: `<p:custDataLst>` requires at least one child,
    // so leaving it behind trades a wrong reference for an invalid one.
    expect(cSld ? child(cSld, P_NS, "custDataLst") : undefined).toBeUndefined();
    expect((await targetsByType(pkg, target)).get(REL_TYPE.tags)).toEqual(["../tags/tag9.xml"]);

    // The template's own tag part and reference belong to the template.
    expect(pkg.has("ppt/tags/tag8.xml")).toBe(true);
    expect((await targetsByType(pkg, "ppt/slides/slide1.xml")).get(REL_TYPE.tags)).toEqual([
      "../tags/tag9.xml",
      "../tags/tag8.xml",
    ]);
  });

  it("keeps a tag part a SHAPE names, because that is somebody else's bookkeeping", async () => {
    /**
     * A deck touched by think-cell carries `<p:nvPr><p:custDataLst><p:tags>` on
     * a hidden shape in every slide it has seen. Removing every tag
     * relationship left that shape naming one that was gone — and that was the
     * visible half. Deleting a relationship FREES ITS ID, and the next tag
     * written takes the next free one, so the vendor's shape came out of the
     * clone pointing at OUR metadata: a reference that still resolves, to
     * somebody else's data.
     */
    const pkg = await taggedDeck();
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    const doc = await shipped(pkg, target);
    const shapeTag = elements(doc, P_NS, "tags").map((el) => el.getAttributeNS(R_NS, "id") ?? el.getAttribute("r:id"));
    expect(shapeTag, "the shape's reference went with the slide's").toEqual(["rId30"]);
    expect((await targetsByType(pkg, target)).get(REL_TYPE.tags)).toEqual(["../tags/tag9.xml"]);
    expect(await pkg.text("ppt/tags/tag9.xml")).toContain("VENDOR");
  });

  it("keeps a custom data list that held more than the tag reference", async () => {
    // `<p:custDataLst>` takes `<p:custData>` as well, and that is not this
    // add-in's to drop. Only the list emptied BY the drop goes.
    const pkg = await taggedDeck(`<p:custData r:id="rId1"/>`);
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });

    const doc = await shipped(pkg, target);
    const cSld = element(doc, P_NS, "cSld");
    const custData = cSld ? child(cSld, P_NS, "custDataLst") : undefined;
    expect(custData, "the whole list went with the tag reference").toBeTruthy();
    expect(child(custData!, P_NS, "tags")).toBeUndefined();
    expect(child(custData!, P_NS, "custData")).toBeTruthy();
  });

  it("drops a comment the template carried, rather than sharing it or copying it", async () => {
    /**
     * A comment hangs off the SLIDE, so the wholesale rels copy hands every
     * clone a relationship to the template's comment part — a reviewer's "check
     * this with Legal" appearing on every copy, as one shared thread. Copying
     * the part per clone would be worse rather than better: the same note, many
     * times, deliberately.
     *
     * Dropping them is also what makes the two template routes AGREE. On a 1.10
     * host `exportAsBase64Presentation` drops comments and `ppt/authors.xml`
     * outright (office-js#6867, measured on that host 2026-08-28), so the
     * subset route was already producing comment-free clones while the file
     * route produced shared ones.
     */
    const part = "ppt/comments/modernComment_101_AEAB9DA1.xml";
    const pkg = await editedDeck([{ paragraphs: [["Hello"]] }], async (zip) => {
      zip.file(part, '<?xml version="1.0"?><cm/>');
      await spliceInto(
        zip,
        "ppt/slides/_rels/slide1.xml.rels",
        "</Relationships>",
        `<Relationship Id="rId9" Type="${MODERN_COMMENT}" Target="../comments/modernComment_101_AEAB9DA1.xml"/>`,
      );
    });
    expect(
      (await targetsByType(pkg, "ppt/slides/slide1.xml")).get(MODERN_COMMENT),
      "the fixture edit did not land",
    ).toBeTruthy();

    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect((await targetsByType(pkg, target)).get(MODERN_COMMENT)).toBeUndefined();
    // The part itself stays: it is the template's, and the template still
    // points at it.
    expect(pkg.has(part)).toBe(true);
    expect((await targetsByType(pkg, "ppt/slides/slide1.xml")).get(MODERN_COMMENT)).toEqual([
      "../comments/modernComment_101_AEAB9DA1.xml",
    ]);
  });
});

describe("the notes page a slide owns", () => {
  /**
   * Exported because a copy gets its own notes page precisely so the copies can
   * differ, and that only pays if whatever fills a slide can reach the notes
   * too. Every "no" it can answer is a case where a caller must not be handed a
   * path, and three of the four are shapes a real deck arrives in.
   */
  it("is the part the notes relationship names", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]], notes: true }]);
    expect(await notesPathFor(pkg, "ppt/slides/slide1.xml")).toBe("ppt/notesSlides/notesSlide1.xml");

    // And after a clone it is the COPY's own page, which is the whole reason
    // the notes are duplicated at all.
    const target = await cloneSlide(pkg, "ppt/slides/slide1.xml", { creationId: () => 222 });
    expect(await notesPathFor(pkg, target)).toBe("ppt/notesSlides/notesSlide2.xml");
  });

  it("is nothing when the slide has no notes relationship", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    expect(await notesPathFor(pkg, "ppt/slides/slide1.xml")).toBeUndefined();
  });

  it("is nothing for a part that has no relationships at all", async () => {
    const pkg = await deck([{ paragraphs: [["Hello"]] }]);
    expect(pkg.has(Pkg.relsPathFor("ppt/theme/theme1.xml")), "the fixture gave the theme a rels part").toBe(false);
    expect(await notesPathFor(pkg, "ppt/theme/theme1.xml")).toBeUndefined();
  });

  it("is nothing when the relationship leads nowhere", async () => {
    // Two ways to lead nowhere and the same answer for both: a `Target` naming
    // a part the package does not hold, and a `Target` that is empty. A path
    // handed back for either is a path a caller then reads.
    const gone = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], (zip) => {
      zip.remove("ppt/notesSlides/notesSlide1.xml");
    });
    expect(await notesPathFor(gone, "ppt/slides/slide1.xml")).toBeUndefined();

    const empty = await editedDeck([{ paragraphs: [["Hello"]], notes: true }], async (zip) => {
      const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
      zip.file("ppt/slides/_rels/slide1.xml.rels", rels.replace("../notesSlides/notesSlide1.xml", ""));
    });
    expect(await notesPathFor(empty, "ppt/slides/slide1.xml")).toBeUndefined();
  });
});
