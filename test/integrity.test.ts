/**
 * Whether the packages this engine produces are internally consistent.
 *
 * Ported from SSF-Merge's `test/integrity.test.ts` on 2026-09-08, minus the
 * blocks that ran its merge. SSF-Merge found two defects on 2026-08-29 that
 * were the same mistake in the two places that delete a relationship, and no
 * gate it had could see either. The rule those defects broke is not about
 * charts or tags: the parts, the relationships and the markup have to agree.
 * So the check is structural, knows nothing about what a deck is for, and is
 * pointed at everything this suite can build.
 *
 * `scripts/package-integrity.mjs` holds the rules, so a script and this file
 * cannot read different ones.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
// @ts-expect-error — a plain .mjs tool with no types. The rules live THERE so
// the suite and any script share one copy.
import { packageProblems, relPrefixesIn, resolvePart } from "../scripts/package-integrity.mjs";
import { Pkg, resolveTarget } from "../src/core/pptx/pkg.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

type Parts = Map<string, string | Uint8Array>;

async function partsOf(bytes: Uint8Array): Promise<Parts> {
  const zip = await JSZip.loadAsync(bytes);
  const parts: Parts = new Map();
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name);
    if (!file || file.dir) continue;
    const xml = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, xml ? await file.async("string") : await file.async("uint8array"));
  }
  return parts;
}

const problems = (parts: Parts): string[] => packageProblems(parts) as string[];

/** Everything the fixture can put on one slide, which no other test combines. */
const EVERYTHING: SlideSpec = {
  paragraphs: [["Name"]],
  notes: "note",
  chart: {
    title: "Title",
    categories: ["a", "b"],
    workbook: ["a"],
    values: ["1", "2"],
    // A callout drawn ON the chart, which lives in a drawing part the CHART
    // points at.
    callout: "callout",
  },
  modernChart: { title: "Title", categories: ["a"], series: "s", workbook: ["a"] },
  smartArt: ["one", "second"],
  smartArtDrawingOn: "slide",
  icons: true,
  shapeTags: true,
};

describe("a deck the fixture builds", () => {
  it("is consistent before anything touches it", async () => {
    // The base every case below breaks one way. If this is not clean, every
    // `toEqual([])` below is asserting something about the fixture instead.
    expect(problems(await partsOf(await makeDeck([EVERYTHING, { paragraphs: [["after"]] }])))).toEqual([]);
  });
});

describe("the checker itself", () => {
  /**
   * A checker that has never failed is a checker nobody has tested.
   *
   * Each case breaks a good package one way and asserts the problem is named.
   * The cases are the ones SSF-Merge's engine could produce, and two of them
   * are exactly what it shipped (its #124 and #126).
   */
  async function goodParts(): Promise<Parts> {
    return partsOf(await makeDeck([EVERYTHING, { paragraphs: [["after"]] }]));
  }

  const slideOf = (parts: Parts): string =>
    [...parts.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()[0] ?? "";

  it("names a relationship whose target is not in the package", async () => {
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    parts.set(rels, (parts.get(rels) as string).replace(/Target="\.\.\/charts\/[^"]*"/, 'Target="../charts/gone.xml"'));
    expect(problems(parts).join("\n")).toContain("which is not in the package");
  });

  it("names markup that references a relationship the part does not have", async () => {
    // SSF-Merge #124: the relationship deleted out from under live markup.
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    const body = parts.get(rels) as string;
    const id = /Id="([^"]*)"[^>]*relationships\/image/.exec(body)?.[1] ?? "";
    expect(id, "the fixture drew no image relationship, so this proves nothing").not.toBe("");
    parts.set(rels, body.replace(new RegExp(`<Relationship Id="${id}"[^>]*/>`), ""));
    expect(problems(parts).join("\n")).toContain("names a relationship the part does not have");
  });

  it("says nothing about an EMPTY relationship id, which names nothing", async () => {
    // PowerPoint's own SmartArt markup, and the false alarm of 2026-08-30.
    //
    // `r:blip=""` on a `<dgm:shape>` means "no picture here". The layout part
    // carrying it has no `.rels` beside it and correctly needs none, so reading
    // the empty string as a relationship id reported four problems per layout
    // part on a sound package — sixteen on that round's deck, every one of them
    // this, none of them real, on a deck PowerPoint opened with no repair.
    //
    // No fixture caught it: this suite's SmartArt is built by the fixture, and
    // the fixture writes no layout part. PowerPoint writes one. That gap is the
    // reason the real-host round exists, so the shape is pinned here as raw
    // markup rather than waiting for the next deck a human makes by hand.
    const parts = await goodParts();
    parts.set(
      "ppt/diagrams/layout9.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<dgm:shape r:blip=""/><dgm:shape r:blip=""/>' +
        "</dgm:layoutDef>",
    );
    expect(problems(parts).join("\n")).not.toContain("names a relationship the part does not have");
  });

  it("says nothing about a target given from the package ROOT", async () => {
    // Legal OOXML and the second false alarm of this kind: a `Target` that
    // begins with `/` is already a part name. Resolved relatively it became
    // `ppt/slides/ppt/media/image1.png`, a part no package holds, and the
    // checker reported a sound deck as pointing at something not in it —
    // `rId2 points at /ppt/slides/slide1.xml, which is not in the package`, on
    // a deck plainly holding that slide.
    //
    // Nothing this engine writes uses the spelling; a deck somebody is sent
    // may, which is exactly the population this checker exists for.
    const parts = await goodParts();
    const rels = "ppt/_rels/presentation.xml.rels";
    const body = parts.get(rels) as string;
    const absolute = body.replace(/Target="slides\/([^"]*)"/, 'Target="/ppt/slides/$1"');
    expect(absolute, "the fixture had no relative slide target to make absolute").not.toBe(body);
    parts.set(rels, absolute);
    expect(problems(parts)).toEqual([]);
  });

  it("says nothing about a part name written with a percent escape", async () => {
    // A relationship Target is a URI reference, so a part name holding a space
    // is written `my%20photo.png` while the zip entry is `my photo.png`.
    // Compared raw it was "not in the package" on a package that holds it —
    // `rId4 points at ../media/my%20photo.png` — which is the third invented
    // problem of this family and the same population as the other two: a deck
    // written by something other than PowerPoint.
    const parts = await goodParts();
    const media = [...parts.keys()].find((n) => /^ppt\/media\/.+/.test(n));
    expect(media, "the fixture drew no media part, so this proves nothing").toBeTruthy();
    const base = (media as string).split("/").pop() as string;
    parts.set(`ppt/media/my photo.png`, parts.get(media as string) as Uint8Array);
    parts.delete(media as string);
    let rewritten = 0;
    for (const [name, body] of [...parts]) {
      if (!name.endsWith(".rels") || typeof body !== "string" || !body.includes(base)) continue;
      parts.set(name, body.split(base).join("my%20photo.png"));
      rewritten++;
    }
    expect(rewritten, "no relationship named the media part, so this proves nothing").toBeGreaterThan(0);
    expect(problems(parts)).toEqual([]);
  });

  it("reads a reference through the prefix the part BINDS, not through `r`", async () => {
    // `r` is only a name. A part may bind the relationships namespace to any
    // prefix, and binding `r` to something else is equally legal — in which
    // case `r:embed` is a different attribute that happens to share a spelling.
    //
    // Both directions were wrong, and this file's own docstring claimed the
    // second could not happen: a part using `rel:` had every reference skipped,
    // so a genuinely dangling one went unreported, and a part binding `r:`
    // elsewhere was reported as naming relationships it does not have.
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    expect(relPrefixesIn(`<p:sld xmlns:rel="${REL}"/>`)).toEqual(["rel"]);
    expect(relPrefixesIn(`<p:sld xmlns:r="${REL}"/>`)).toEqual(["r"]);
    // Declares nothing: a fragment, or a part shown without the parent it
    // inherits the binding from.
    expect(relPrefixesIn(`<a:blip r:embed="rId9"/>`)).toEqual(["r"]);
    // Binds `r` to something else, so `r:` names no relationship here.
    expect(relPrefixesIn(`<p:sld xmlns:r="urn:something-else"/>`)).toEqual([]);

    // And behaviourally, in the direction a unit assertion cannot reach: a
    // slide that names a relationship through `rel:` is checked like any other.
    // Under the prefix-literal reader this reference was invisible, so a
    // dangling one was reported as nothing at all — the check was not failing,
    // it was not running.
    const parts = await goodParts();
    const slide = slideOf(parts);
    const body = parts.get(slide) as string;
    parts.set(
      slide,
      body
        .replace("<p:sld ", `<p:sld xmlns:rel="${REL}" `)
        .replace("<p:cSld>", `<p:cSld><p:custom rel:embed="rIdNope"/>`),
    );
    expect(problems(parts).join("\n")).toContain("names a relationship the part does not have");
  });

  it("names a reference that leads to the wrong KIND of part", async () => {
    // SSF-Merge #126: the id freed by a delete and taken by the next thing that needed
    // one, so the reference still resolves — to somebody else's data. Both
    // checks above pass on this.
    const parts = await goodParts();
    const slide = slideOf(parts);
    const rels = `ppt/slides/_rels/${slide.split("/").pop()}.rels`;
    parts.set(
      rels,
      (parts.get(rels) as string).replace(
        /(<Relationship Id="[^"]*" Type="[^"]*)relationships\/tags("[^>]*Target=")[^"]*"/,
        '$1relationships/image$2../media/whatever.png"',
      ),
    );
    const said = problems(parts).join("\n");
    expect(said).toContain("wants tags and leads to image");
  });

  it("names a part no content type covers", async () => {
    // An EXTENSION with no Default, which is the shape this can actually take.
    // A first attempt added a stray `.xml` part and the check said nothing —
    // rightly: every deck declares `Default Extension="xml"`, so that one is
    // covered by definition. The real case is an embedding or a media file
    // whose extension nothing declares, which is what stripping a part out of a
    // package leaves behind if its Default goes with it.
    const parts = await goodParts();
    parts.set("ppt/embeddings/oleObject1.bin", new Uint8Array([1, 2, 3]));
    expect(problems(parts).join("\n")).toContain("no content type covers it");
  });

  it("names a part holding a character XML cannot carry", async () => {
    // The one defect a round trip through this repo's own parser cannot show:
    // `@xmldom/xmldom` writes such a character out and reads it back, so the
    // document is fine at both ends and the FILE is not.
    for (const code of [0x00, 0x0b, 0x1f, 0xd800]) {
      const parts = await goodParts();
      const slide = slideOf(parts);
      parts.set(slide, (parts.get(slide) as string).replace("</a:t>", String.fromCharCode(code) + "</a:t>"));
      expect(problems(parts).join("\n")).toContain("which XML cannot carry");
    }
  });

  it("says nothing about an emoji, which is a legal surrogate PAIR", async () => {
    // The false alarm the `u` flag exists to avoid: matched code UNIT by code
    // unit, every astral character in every deck is two "surrogates".
    const parts = await goodParts();
    const slide = slideOf(parts);
    parts.set(slide, (parts.get(slide) as string).replace("</a:t>", "\u{1F600}</a:t>"));
    expect(problems(parts)).toEqual([]);
  });
});

describe("the engine and the checker resolve a relationship the same way", () => {
  // Owner PART and target, never a `.rels` path: both functions are given the
  // part that OWNS the relationship, and a rels path would be a misuse that
  // says nothing about either.
  const pairs: [string, string][] = [
    ["ppt/slides/slide1.xml", "../charts/chart1.xml"],
    ["ppt/slides/slide1.xml", "../media/image1.png"],
    ["ppt/slides/slide1.xml", "chart1.xml"],
    ["ppt/slides/slide1.xml", "./chart1.xml"],
    ["ppt/charts/chart1.xml", "../embeddings/Microsoft_Excel_Worksheet.xlsx"],
    ["ppt/presentation.xml", "slides/slide1.xml"],
    ["ppt/presentation.xml", "../ppt/slides/slide1.xml"],
    ["[Content_Types].xml", "ppt/presentation.xml"],
    // The one they disagreed on.
    ["ppt/slides/slide1.xml", "/ppt/media/image1.png"],
    ["ppt/charts/chart1.xml", "/ppt/embeddings/wb.xlsx"],
    // A Target is a URI reference and a part name is not, so a part called
    // `my chart.xml` is written escaped. Both branches decode.
    ["ppt/slides/slide1.xml", "../charts/my%20chart.xml"],
    ["ppt/slides/slide1.xml", "/ppt/embeddings/Sales%20Data.xlsx"],
    ["ppt/slides/slide1.xml", "../media/100%.png"],
    ["ppt/slides/slide1.xml", "../charts/%2E%2E"],
  ];

  it.each(pairs)("agrees on %s + %s", (owner, target) => {
    expect(resolvePart(owner, target)).toBe(resolveTarget(owner, target));
  });

  it("resolves a percent escape to the name the package holds", () => {
    // Stated as a VALUE, not only as an agreement: a zip entry name is the
    // literal name, so an escaped answer matches nothing. `pkg.has` then says
    // no and `cloneSlideGraphics` skips a chart it cannot find — every merged
    // copy keeps pointing at the template's, and the whole deck shows the last
    // record's data with nothing said.
    for (const resolve of [resolveTarget, resolvePart]) {
      expect(resolve("ppt/slides/slide1.xml", "../charts/my%20chart.xml")).toBe("ppt/charts/my chart.xml");
      expect(resolve("ppt/slides/slide1.xml", "/ppt/embeddings/Sales%20Data.xlsx")).toBe(
        "ppt/embeddings/Sales Data.xlsx",
      );
      // Not valid encoding: kept exactly as it stands, because a part really
      // called `100%.png` is a better answer than a throw out of a merge.
      expect(resolve("ppt/slides/slide1.xml", "../media/100%.png")).toBe("ppt/media/100%.png");
      // Decoded AFTER the walk, so this is a segment named two dots and not a
      // step upward — the package's own sweep deletes what it is pointed at.
      expect(resolve("ppt/slides/slide1.xml", "../charts/%2E%2E")).toBe("ppt/charts/..");
    }
  });

  it("resolves a root-relative target to the part it names", () => {
    // The case that was wrong, stated as a value rather than as an agreement:
    // two functions can agree and both be wrong, and this is what the answer is.
    expect(resolveTarget("ppt/slides/slide1.xml", "/ppt/media/image1.png")).toBe("ppt/media/image1.png");
    expect(resolvePart("ppt/slides/slide1.xml", "/ppt/media/image1.png")).toBe("ppt/media/image1.png");
  });
});

describe("removing one of two slides that share a chart", () => {
  /**
   * `orphanedParts` decides which of a removed slide's parts nothing else needs
   * — and it used to skip the `.rels` of every CANDIDATE while the scan was
   * running, which is the answer the scan is computing. So a chart another
   * slide still references was kept, its own relationships were skipped all the
   * same, and its embedded workbook finished the scan with no referrer at all
   * and was swept.
   *
   * What comes out is a surviving chart pointing at a part that is not in the
   * package, which is exactly what PowerPoint calls damaged — from a merge that
   * reported success. Two slides sharing one chart is all it takes, and a deck
   * where somebody duplicated a slide by hand has them.
   */
  it("keeps the workbook the surviving slide's chart still needs", async () => {
    const zip = await JSZip.loadAsync(
      await makeDeck([
        { paragraphs: [["a"]], chart: { title: "shared", workbook: ["shared"] } },
        { paragraphs: [["b"]] },
      ]),
    );
    const rels = zip.file("ppt/slides/_rels/slide2.xml.rels");
    if (!rels) throw new Error("the fixture wrote no rels for slide 2");
    zip.file(
      "ppt/slides/_rels/slide2.xml.rels",
      (await rels.async("string")).replace(
        "</Relationships>",
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"` +
          ` Target="../charts/chart1.xml"/></Relationships>`,
      ),
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    expect(problems(await partsOf(bytes)), "the deck going in").toEqual([]);

    const pkg = await Pkg.open(bytes);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    const parts = await partsOf(await pkg.toBytes());

    expect(problems(parts)).toEqual([]);
    // The chart survives because slide 2 names it, and so must everything it
    // needs — asserted by name, because "no problems" would also be satisfied
    // by the chart and its rels having gone too.
    expect(parts.has("ppt/charts/chart1.xml")).toBe(true);
    expect(parts.has("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")).toBe(true);
  });

  it("still takes the whole chart when the slide really was the last owner", async () => {
    // The other direction, because a fix that keeps everything is not a fix.
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["a"]], chart: { title: "own", workbook: ["own"] } }, { paragraphs: [["b"]] }]),
    );
    await pkg.removeSlide("ppt/slides/slide1.xml");
    const parts = await partsOf(await pkg.toBytes());
    expect(problems(parts)).toEqual([]);
    expect([...parts.keys()].filter((n) => n.includes("chart") || n.includes("embeddings"))).toEqual([]);
  });
});
