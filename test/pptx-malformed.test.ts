import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Pkg } from "../src/core/pptx/pkg.js";
import { REL_TYPE } from "../src/core/pptx/parts.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * Packages whose XML is not shaped the way the reader assumes.
 *
 * `pkg.ts` is full of guards for attributes that ought to be there and are
 * not — `getAttribute("Target")` answering null, a `<p:sldId>` with no `r:id`,
 * a `<Default>` with no `Extension`. Every one of them is a decision about what
 * to do with a deck somebody else's software wrote, and on 2026-09-12 the file
 * measured **87.11% of its branches**, the lowest in `src/core/pptx`: the
 * guards were there and the cases that reach them were not.
 *
 * They are not hypothetical. `CLAUDE.md` is a list of the times a real host or
 * a real deck did something the format permits and nobody expected, and every
 * deck this add-in touches is a user's own — read through `getFileAsync`, which
 * hands back whatever they have, written by whatever wrote it. A guard that has
 * never been executed is a guess about what it does.
 *
 * The file gets bytes rather than a hand-built object throughout: these are
 * REAL packages, built by the fixture and then damaged in one specific way, so
 * a case cannot pass because the fixture happens to be missing something else.
 *
 * **Two kinds of case live here, and they are labelled.** Writing them meant
 * removing each guard to watch the case go red, and seven guards did not make
 * anything go red: a second line further down already produced the same answer.
 * Those are marked BELT AND BRACES with what actually catches it, measured on
 * 2026-09-12. They stay — the guards and the cases both. A guard whose intent
 * is written out is worth keeping in a reader of other people's files, and a
 * case that pins the BEHAVIOUR ("a malformed entry is skipped") is worth
 * keeping whichever line delivers it. What is not worth keeping is a comment
 * claiming a case proves something it does not, which is what all seven of
 * these said before they were tried.
 */

const TWO = [{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }];
const CONTENT_TYPES = "[Content_Types].xml";
const PRESENTATION = "ppt/presentation.xml";
const PRES_RELS = "ppt/_rels/presentation.xml.rels";
const SLIDE1 = "ppt/slides/slide1.xml";
const SLIDE1_RELS = "ppt/slides/_rels/slide1.xml.rels";
const SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const REL_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/**
 * A real fixture deck with named parts replaced, or removed with `null`.
 *
 * Rebuilt as a zip and reopened, rather than `setText` on an open package: a
 * part REMOVED is the case several of these need, and `Pkg` has no way to
 * remove one from the outside — which is correct, and is why this goes through
 * the bytes.
 */
async function craft(edits: Record<string, string | null>): Promise<Pkg> {
  const source = await JSZip.loadAsync(await makeDeck(TWO));
  const out = new JSZip();
  for (const [name, file] of Object.entries(source.files)) {
    if (file.dir) continue;
    if (!(name in edits)) {
      out.file(name, await file.async("uint8array"));
      continue;
    }
    const replacement = edits[name];
    if (replacement !== null && replacement !== undefined) out.file(name, replacement);
  }
  // Parts the fixture has no copy of, appended in the order given — which two
  // of the cases below depend on, because a zip is read back in the order it
  // was written and one counter walks it.
  for (const [name, replacement] of Object.entries(edits)) {
    if (replacement !== null && !(name in source.files)) out.file(name, replacement);
  }
  return Pkg.open(await out.generateAsync({ type: "uint8array" }));
}

/** A `.rels` part holding exactly these `<Relationship>` elements. */
function rels(...entries: string[]): string {
  return `${REL_HEAD}\r\n<Relationships xmlns="${RELS_NS}">${entries.join("")}</Relationships>`;
}

/** A `[Content_Types].xml` holding exactly these children. */
function types(...entries: string[]): string {
  return `${REL_HEAD}\r\n<Types xmlns="${CT_NS}">${entries.join("")}</Types>`;
}

describe("a content types part that does not say everything it should", () => {
  it("finds the Default for an extension past one that names no extension", async () => {
    // `<Default>` without `Extension` is schema-invalid and a reader still has
    // to survive it. The guard is `(d.getAttribute("Extension") ?? "")`, and
    // without the `?? ""` this throws on the null rather than skipping the
    // element — taking every caller that asks a content type with it.
    const pkg = await craft({
      [CONTENT_TYPES]: types(
        '<Default ContentType="application/octet-stream"/>',
        '<Default Extension="png" ContentType="image/png"/>',
      ),
    });
    expect(await pkg.contentTypeOf("ppt/media/image1.png")).toBe("image/png");
  });

  it("adds a Default past one that names no extension, rather than thinking it is already there", async () => {
    const pkg = await craft({
      [CONTENT_TYPES]: types('<Default ContentType="application/octet-stream"/>'),
    });
    await pkg.addContentTypeDefault("emf", "image/x-emf");
    expect(await pkg.contentTypeOf("ppt/media/image1.emf")).toBe("image/x-emf");
  });

  it("keeps the FIRST of two Overrides naming the same part", async () => {
    // Two Overrides for one part is schema-invalid, and a reader still has to
    // pick one. `!index.has(name)` makes it the first, which is the same one a
    // linear scan would have found — so the index that replaced the scan did
    // not quietly change the answer. Nothing was holding it to that.
    const pkg = await craft({
      [CONTENT_TYPES]: types(
        `<Override PartName="/${SLIDE1}" ContentType="application/x-first"/>`,
        `<Override PartName="/${SLIDE1}" ContentType="application/x-second"/>`,
      ),
    });
    expect(await pkg.contentTypeOf(SLIDE1)).toBe("application/x-first");
  });

  it("says nothing for an Override that names no content type, rather than answering null", async () => {
    // The `?? undefined` turns the DOM's null into the absence this function's
    // signature promises. A null reaching a caller that has already checked for
    // undefined is the shape of bug that survives typechecking.
    const pkg = await craft({
      [CONTENT_TYPES]: types(`<Override PartName="/${SLIDE1}"/>`),
    });
    expect(await pkg.contentTypeOf(SLIDE1)).toBeUndefined();
  });

  it("forgets its index of Overrides when the whole part is written over", async () => {
    /**
     * The cache invalidation, which nothing was exercising.
     *
     * `overrideIndex` builds a map once and keeps it. `addContentTypeOverride`
     * goes through the parsed document and updates the map with it, so the two
     * stay together — but a caller that replaces the whole part with `setText`
     * leaves an index built from bytes that are gone. `noteWritten` drops it
     * for exactly that reason, and until now no case wrote that part.
     *
     * The failure it prevents is silent and one-directional: the STALE answer
     * is returned, so a part declared as one thing is copied as another, and
     * the deck opens damaged rather than the run raising.
     */
    const pkg = await craft({});
    expect(await pkg.contentTypeOf(SLIDE1), "the index is built by this first read").toBe(SLIDE_TYPE);
    pkg.setText(CONTENT_TYPES, types(`<Override PartName="/${SLIDE1}" ContentType="application/x-rewritten"/>`));
    expect(await pkg.contentTypeOf(SLIDE1)).toBe("application/x-rewritten");
  });
});

describe("relationships that are missing the attributes they are named by", () => {
  it("numbers a new relationship past ids it cannot read as rIdN", async () => {
    // Two ways a `.rels` in the wild is not a run of `rId1..rIdN`: a
    // relationship with no `Id` at all, and one whose id another tool chose.
    //
    // BELT AND BRACES. The `?? ""` and `?? 0` around the id read are not what
    // saves this: `exec(null)` coerces to the string "null", which matches
    // nothing, so the result is `NaN` — and `NaN > max` is false, so the
    // maximum survives on its own. Removing both leaves this green. What the
    // case pins is the numbering itself.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Type="${REL_TYPE.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`,
        `<Relationship Id="docRel" Type="http://example/t" Target="a.xml"/>`,
        `<Relationship Id="rId5" Type="http://example/t" Target="b.xml"/>`,
      ),
    });
    expect(await pkg.addRel(SLIDE1, "http://example/t", "c.xml")).toBe("rId6");
  });

  it("skips a relationship with no target when listing what a part points at", async () => {
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Id="rId1" Type="${REL_TYPE.slideLayout}"/>`,
        `<Relationship Id="rId2" Type="http://example/t" Target="../media/image1.png"/>`,
      ),
    });
    expect(await pkg.relatedParts(SLIDE1)).toEqual(["ppt/media/image1.png"]);
  });

  it("names a part once when two relationships point at it", async () => {
    // `!out.includes(resolved)` is the dedupe, and it decides what
    // `orphanedParts` counts: a slide that reaches its chart twice must not
    // report the chart twice, or the removal sweep visits it twice.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Id="rId1" Type="http://example/t" Target="../media/image1.png"/>`,
        `<Relationship Id="rId2" Type="http://example/t" Target="../media/image1.png"/>`,
      ),
    });
    expect(await pkg.relatedParts(SLIDE1)).toEqual(["ppt/media/image1.png"]);
  });

  it("lists no slides at all when the presentation has no relationships part", async () => {
    // Not a deck PowerPoint would write, and the answer still has to be a list.
    // `slidePaths` resolves every `<p:sldId>` through the presentation's own
    // `.rels`; with the part gone there is nothing to resolve any of them to,
    // and an empty list is the honest answer rather than a raise.
    const pkg = await craft({ [PRES_RELS]: null });
    expect(await pkg.slidePaths()).toEqual([]);
  });

  it("skips a presentation relationship with no id when resolving the slide list", async () => {
    // BELT AND BRACES. `!id` is not what skips it: an entry keyed by the empty
    // string matches no `r:id` any `<p:sldId>` can carry, so it is unreachable
    // either way.
    const pkg = await craft({
      [PRES_RELS]: rels(
        `<Relationship Type="${REL_TYPE.slide}" Target="slides/slide1.xml"/>`,
        `<Relationship Id="rId3" Type="${REL_TYPE.slide}" Target="slides/slide2.xml"/>`,
      ),
    });
    // `rId2` names slide 1 in the fixture's presentation and now resolves to
    // nothing, so only slide 2 is listed — from the entry that kept its id.
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });
});

describe("a slide list that is not a list of resolvable slides", () => {
  /** The fixture's presentation, with its `<p:sldIdLst>` replaced by `body`. */
  async function presentationWith(body: string): Promise<Pkg> {
    const source = await JSZip.loadAsync(await makeDeck(TWO));
    const xml = await (source.file(PRESENTATION) as JSZip.JSZipObject).async("string");
    return craft({ [PRESENTATION]: xml.replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/, body) });
  }

  it("skips a slide entry carrying no relationship id", async () => {
    // BELT AND BRACES. `if (!rId) continue` is not what skips it —
    // `targets.get(null)` is undefined and the `if (target)` below drops it
    // anyway. The behaviour is the point: one unusable entry costs its own
    // slide and not the list.
    const pkg = await presentationWith('<p:sldIdLst><p:sldId id="256"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>');
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });

  it("removes a slide from a presentation that has no slide list at all", async () => {
    // Nothing to unlink, and the part and its relationships still have to go.
    // The `list ? … : []` is what keeps this from throwing on the null.
    const pkg = await presentationWith("");
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    expect(pkg.has(SLIDE1_RELS)).toBe(false);
  });

  it("removes a slide past an entry carrying no relationship id", async () => {
    const pkg = await presentationWith('<p:sldIdLst><p:sldId id="256"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>');
    // `rId2` is slide 1 in the fixture, so the second entry is the one that
    // matches and the first is skipped rather than read as a match.
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    expect(await pkg.text(PRESENTATION)).not.toContain('r:id="rId2"');
  });

  it("gives a new slide an id past an entry that carries no id", async () => {
    // BELT AND BRACES, confirmed rather than assumed: `Number(null)` is 0 and
    // `NaN > max` is false, so the maximum survives with or without the `?? 0`.
    // Removing it leaves this green. The case pins that a deck carrying a
    // malformed entry still gets a unique, in-range id for its new slide.
    const pkg = await presentationWith(
      '<p:sldIdLst><p:sldId r:id="rId2"/><p:sldId id="300" r:id="rId3"/></p:sldIdLst>',
    );
    expect(await pkg.appendSldId("rId9")).toBe(301);
  });
});

describe("removing a slide whose own relationships are damaged", () => {
  it("removes a slide that has no relationships part at all", async () => {
    // `if (this.has(relsPath))` guards the whole notes-and-orphans sweep. A
    // slide with no `.rels` is unusual and legal — nothing it points at — and
    // the part itself still has to go, along with its entry in the slide list.
    const pkg = await craft({ [SLIDE1_RELS]: null });
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });

  it("skips a relationship that declares no type when deciding what goes with the slide", async () => {
    // The sweep DELETES what a notes or comment relationship points at, so
    // "what type is this" decides whether a part is destroyed, and an entry
    // with no `Type` must read as "not one of those".
    //
    // BELT AND BRACES. The `?? ""` is not what does it: `null` is not equal to
    // the notes type and `COMMENT_REL_TYPES.includes(null)` is false, so the
    // entry is skipped without it. The behaviour is what matters — one
    // malformed entry must not take the whole removal with it.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Id="rId1" Target="../media/image1.png"/>`,
        `<Relationship Id="rId2" Type="${REL_TYPE.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`,
      ),
    });
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    expect(pkg.has("ppt/slideLayouts/slideLayout1.xml"), "the shared layout is not the slide's to take").toBe(true);
  });

  it("leaves a notes relationship with no target alone rather than resolving nothing", async () => {
    // A notes relationship is one of the few this sweep DELETES what it points
    // at, so a missing target is the case where it must do nothing at all — the
    // alternative being `resolved(slidePath, null)` and a removal aimed at
    // whatever that produces.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Id="rId1" Type="${REL_TYPE.notesSlide}"/>`,
        `<Relationship Id="rId2" Type="${REL_TYPE.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`,
      ),
    });
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    // The layout is shared with slide 2 and stays, which is the check that the
    // sweep did not simply give up when it met the damaged entry.
    expect(pkg.has("ppt/slideLayouts/slideLayout1.xml")).toBe(true);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });
});

describe("parts named like something they are not", () => {
  it("counts the highest media number whatever order the zip lists them in", async () => {
    // `nextFree` walks a set in the order the zip yielded it, so the maximum is
    // only right if a SMALLER number after a bigger one leaves it alone. A zip
    // is not sorted, and PowerPoint has no reason to write media in order.
    const pkg = await craft({
      "ppt/media/image10.png": "not really a png",
      "ppt/media/image2.png": "nor is this",
    });
    expect(pkg.nextMediaNumber()).toBe(11);
  });

  it("ignores a part merely named .rels when looking for what refers to what", async () => {
    /**
     * The orphan sweep reads every `.rels` in the package to find out what
     * still points at a part. It finds them by name, and the two conditions are
     * not the same question: `endsWith(".rels")` says it looks like one, and
     * `includes("/_rels/")` says it IS one.
     *
     * A part called `ppt/notes.rels` sitting outside a `_rels/` folder is legal
     * in a zip and is not a relationships part.
     *
     * BELT AND BRACES, and the reason is worth knowing before trusting the
     * check for anything: the scan never reads the file at the path it
     * collected. It derives an OWNER from the path and then asks that owner for
     * its own relationships, so a path with no `/_rels/` in it resolves to the
     * package root — whose real `_rels/.rels` is already in the list. Removing
     * the condition leaves this green because the extra entry is a duplicate,
     * not because the crafted content is refused. Measured 2026-09-12, with a
     * decoy carrying valid relationship XML that claims the slide's picture.
     */
    // The slide is given a picture of its own, because the scan only runs when
    // there is something whose orphanhood is in question — `orphanedParts`
    // returns early otherwise, and the first version of this case did nothing.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(`<Relationship Id="rId1" Type="${REL_TYPE.image}" Target="../media/image1.png"/>`),
      "ppt/media/image1.png": "not really a png",
      // Valid relationship XML, deliberately: a decoy of arbitrary text proves
      // nothing, because a part yielding no <Relationship> contributes no
      // referrers and the sweep reaches the same answer either way. This one
      // CLAIMS the slide's picture, so reading it as a referrer keeps a part
      // alive that nothing real points at.
      "ppt/notes.rels": rels(`<Relationship Id="rId1" Type="http://example/t" Target="ppt/media/image1.png"/>`),
    });
    await pkg.removeSlide(SLIDE1);
    expect(pkg.has(SLIDE1)).toBe(false);
    // The picture is the slide's own and nothing else refers to it, so the scan
    // ran and reached a conclusion rather than being skipped.
    expect(pkg.has("ppt/media/image1.png")).toBe(false);
    expect(pkg.has("ppt/notes.rels"), "not the slide's, and not read as a referrer either").toBe(true);
  });

  it("leaves a chart's workbook alone when the package does not actually hold it", async () => {
    // A chart owns its embedded workbook, so the sweep follows one hop out and
    // takes it. A DANGLING relationship — a target the package never had — must
    // not turn into anything.
    //
    // BELT AND BRACES. `this.has(child)` keeps it out of the owned list, and
    // without that check the outcome is identical anyway: the path reaches
    // `removePart`, which finds nothing to remove. The check earns its place by
    // keeping the list honest rather than by changing what happens.
    const pkg = await craft({
      [SLIDE1_RELS]: rels(`<Relationship Id="rId1" Type="${REL_TYPE.chart}" Target="../charts/chart1.xml"/>`),
      "ppt/charts/chart1.xml": '<?xml version="1.0"?><c:chartSpace xmlns:c="http://example/c"/>',
      "ppt/charts/_rels/chart1.xml.rels": rels(
        `<Relationship Id="rId1" Type="http://example/t" Target="../embeddings/Workbook.xlsx"/>`,
      ),
    });
    await pkg.removeSlide(SLIDE1);
    // The chart itself is the slide's and goes; the workbook was never there.
    expect(pkg.has("ppt/charts/chart1.xml")).toBe(false);
    expect(pkg.has("ppt/embeddings/Workbook.xlsx")).toBe(false);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });
});
