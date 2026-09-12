import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { HarvestError, Pkg, harvest, parseXml, serializeXml } from "../src/core/index.js";
import type { Names } from "../src/core/index.js";
import { coloursOf, themeColoursFor, themeOf } from "../src/core/pptx/theme.js";
import { pinColoursInXml, pinSchemeColours } from "../src/core/splice/colours.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The colour switch: `docs/DESIGN.md` section 7's second position.
 *
 * "This deck's theme" is the absence of work — a scheme colour resolves against
 * whatever deck it lands in, which is what makes an inserted element look like
 * the customer's. "As in the library" is two pieces, and this file is both:
 * READING the library theme's colours at harvest time, and PINNING an element's
 * scheme colours to them at insert time.
 *
 * The library decks themselves are the last word. They are harvested here, not
 * described: the 16:9 deck states `dk1` as a SYSTEM colour and the 4:3 deck as
 * an explicit one, and a reader that handles only the second loses black out of
 * the map without failing.
 */

/** The stock Office palette, in the spelling PowerPoint writes it. */
const OFFICE = {
  dk1: "sys:windowText:000000",
  lt1: "sys:window:FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "5B9BD5",
  accent2: "ED7D31",
};

/** A shape fragment, parsed the way the splice parses an element's markup. */
function fragment(inner: string): Element {
  const doc = parseXml(
    `<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${inner}</p:sp>`,
  );
  return doc.documentElement;
}

describe("reading a deck's theme colours", () => {
  it("reads both spellings a slot is written in", async () => {
    // `<a:srgbClr>` is the simple one. `<a:sysClr val="windowText"
    // lastClr="000000"/>` is what the stock Office theme uses for dk1 and lt1,
    // and `lastClr` is the only concrete value in the file — a reader that
    // ignores it drops black and white, the two commonest colours in the
    // library, and reports no failure at all.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const colours = await coloursOf(pkg, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml");
    expect(colours["dk1"]).toBe("000000");
    expect(colours["lt1"]).toBe("FFFFFF");
    expect(colours["accent1"]).toBe("5B9BD5");
  });

  it("resolves the four names a slide writes through the master's colour map", async () => {
    // The map is the hop a reader forgets. A slide writes `tx1`, which is not
    // a theme slot at all; the master says which slot it means, and on a DARK
    // master it means the LIGHT one. Assuming tx1 is dk1 is right until the
    // first dark design and then every text colour comes out white on white.
    const dark = await Pkg.open(
      await makeDeck([{ paragraphs: [["One"]] }], {
        scheme: OFFICE,
        clrMap: { bg1: "dk1", tx1: "lt1", bg2: "dk2", tx2: "lt2" },
      }),
    );
    const colours = await coloursOf(dark, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml");
    expect(colours["tx1"]).toBe("FFFFFF");
    expect(colours["bg1"]).toBe("000000");
    expect(colours["tx2"]).toBe("E7E6E6");
  });

  it("walks slide, layout, master, theme", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    expect((await themeColoursFor(pkg, "ppt/slides/slide1.xml"))["accent2"]).toBe("ED7D31");
  });

  it("answers nothing for a theme that states no colours, rather than inventing a palette", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }]));
    expect(await themeColoursFor(pkg, "ppt/slides/slide1.xml")).toEqual({});
  });

  it("reads the colours the committed library decks actually carry, and they are the SAME colours", async () => {
    // The two decks disagreed until 2026-09-11: the 16:9 deck was on the stock
    // Office palette and the 4:3 deck on "07 Blå", a palette it had carried
    // since it was made at another company in 2013 — so the same element came
    // out Office orange in one size and light blue in the other, in the library
    // itself, and "as in the library" pinned it to two different colours. The
    // 4:3 deck was re-themed to the 16:9 deck's scheme.
    //
    // Asserted as an EQUALITY between the decks rather than as two lists of
    // hexes, because the thing that must stay true is that they agree. A future
    // palette change to both keeps this green; a palette change to one does not.
    const wide = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
    const narrow = await Pkg.open(new Uint8Array(readFileSync("template/library-4x3.pptx")));
    const wideColours = await themeColoursFor(wide, (await wide.slidePaths())[1] as string);
    const narrowColours = await themeColoursFor(narrow, (await narrow.slidePaths())[1] as string);
    expect(narrowColours).toEqual(wideColours);
    // And the values themselves, so a deck re-themed to something unreadable
    // cannot pass merely by being re-themed consistently.
    expect(wideColours).toMatchObject({
      accent1: "5B9BD5",
      accent2: "ED7D31",
      tx1: "000000",
      bg1: "FFFFFF",
      tx2: "44546A",
    });
  });
});

describe("a theme slot this reader cannot make sense of", () => {
  it("leaves the slot out rather than putting something wrong in the map", async () => {
    // Neither spelling readable: a value that is not six hex digits, and a
    // system colour whose `lastClr` is not either. The rule is the same as for
    // an absent slot — leave it out, so the pin passes over it and the colour
    // keeps following the destination's theme. A slot half-read into the map
    // is a colour this add-in invented and wrote into somebody's deck.
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["One"]] }], {
        scheme: { ...OFFICE, accent3: "not a colour", accent4: "sys:windowText:zzzzzz" },
      }),
    );
    const colours = await coloursOf(pkg, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml");
    expect(colours["accent3"]).toBeUndefined();
    expect(colours["accent4"]).toBeUndefined();
    // And the readable ones beside them are still there, so this is the slot
    // being skipped rather than the read giving up.
    expect(colours["accent1"]).toBe("5B9BD5");
    expect(colours["dk1"]).toBe("000000");
  });
});

describe("a part that points at no theme", () => {
  it("answers nothing when the part has relationships but none of them is a theme", async () => {
    // A slide's own rels name its layout and nothing else. The walk has to
    // reach the end of a real list and answer undefined, rather than answering
    // the first relationship it finds — which would hand `coloursOf` a layout
    // part and get a map of nothing back, silently.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    expect(await themeOf(pkg, "ppt/slides/slide1.xml")).toBeUndefined();
    // The master, which does point at one, so the case above is the walk
    // finishing rather than the walk being broken.
    expect(await themeOf(pkg, "ppt/slideMasters/slideMaster1.xml")).toBe("ppt/theme/theme1.xml");
  });

  it("answers an empty map for a part whose chain reaches no theme", async () => {
    // `themeColoursFor` short-circuits rather than reading a theme it does not
    // have. An empty map means "pin nothing", which is the safe answer: the
    // element keeps its scheme colours and follows the destination's theme,
    // which is what it would have done under the other setting anyway.
    //
    // The theme part itself is the part with no chain at all — it has no rels
    // file, so the walk stops at the first hop. A MASTER is not: its rels name
    // its layouts, and the chain walks layout → master → theme right back
    // round to the theme it started under.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    expect(await themeColoursFor(pkg, "ppt/theme/theme1.xml")).toEqual({});
    expect(Object.keys(await themeColoursFor(pkg, "ppt/slides/slide1.xml")).length).toBeGreaterThan(4);
  });
});

describe("a theme read with no master, or a broken one", () => {
  it("reads the theme's own slots with no master, and leaves the mapped names out", async () => {
    // The four names a slide writes — tx1, bg1, tx2, bg2 — are the master's to
    // resolve, and without one they are left out rather than guessed. An
    // identity map is a guess that is right most of the time, and the times it
    // is wrong are decks with a dark master, which is exactly where an
    // inserted element would come out with black text on black.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const alone = await coloursOf(pkg, "ppt/theme/theme1.xml");
    expect(alone["accent1"]).toBe("5B9BD5");
    expect(alone["dk1"]).toBe("000000");
    expect(alone["tx1"]).toBeUndefined();
    expect(alone["bg1"]).toBeUndefined();
  });

  it("does the same for a master path that is not in the package", async () => {
    // Not a case a real deck reaches, and one a CALLER reaches by handing on a
    // chain that did not resolve. Answering the theme's own slots is better
    // than throwing: the element still gets most of its colours pinned.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const missing = await coloursOf(pkg, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster9.xml");
    expect(missing["accent1"]).toBe("5B9BD5");
    expect(missing["tx1"]).toBeUndefined();
  });

  it("answers an empty map for a theme part with no colour scheme in it", async () => {
    // `makeDeck` writes `<a:themeElements/>` and nothing else when no scheme is
    // asked for, which is a theme part PowerPoint would accept and this reader
    // has nothing to take from.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }]));
    expect(await coloursOf(pkg, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml")).toEqual({});
  });
});

describe("a relationship that does not lead to a theme", () => {
  /** The master's relationship list, live, so a case can doctor it. */
  const relsOf = async (pkg: Pkg): Promise<Document> => pkg.doc(Pkg.relsPathFor("ppt/slideMasters/slideMaster1.xml"));

  it("skips an EXTERNAL theme relationship rather than resolving it as a path", async () => {
    // `TargetMode="External"` means the target is a URL, not a part, and
    // resolving one as a package path produces a string that happens to look
    // like a part and is not one. Nothing in the library decks carries one; a
    // customer's deck can.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const rels = await relsOf(pkg);
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      if ((rel.getAttribute("Type") ?? "").endsWith("/theme")) rel.setAttribute("TargetMode", "External");
    }
    expect(await themeOf(pkg, "ppt/slideMasters/slideMaster1.xml")).toBeUndefined();
  });

  it("skips a theme relationship whose part is not in the package", async () => {
    // A dangling relationship. Answering the missing path would hand
    // `coloursOf` a part it then fails to open, turning a colour setting into
    // a failed insert; answering nothing means the element keeps its scheme
    // colours, which is what the other setting does anyway.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const rels = await relsOf(pkg);
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      if ((rel.getAttribute("Type") ?? "").endsWith("/theme")) rel.setAttribute("Target", "../theme/theme9.xml");
    }
    expect(await themeOf(pkg, "ppt/slideMasters/slideMaster1.xml")).toBeUndefined();
  });

  it("skips a theme relationship with no target at all", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    const rels = await relsOf(pkg);
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      if ((rel.getAttribute("Type") ?? "").endsWith("/theme")) rel.removeAttribute("Target");
    }
    expect(await themeOf(pkg, "ppt/slideMasters/slideMaster1.xml")).toBeUndefined();
  });

  it("answers nothing for a part with no relationship list of its own", async () => {
    // A theme part has none. The walk stops before it opens anything.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["One"]] }], { scheme: OFFICE }));
    expect(await themeOf(pkg, "ppt/theme/theme1.xml")).toBeUndefined();
  });
});

describe("a colour map naming a slot the theme does not carry", () => {
  it("leaves that mapped name out instead of writing undefined into the map", async () => {
    // The map is the master's, the slots are the theme's, and the two can
    // disagree: a master re-pointed at a leaner theme maps `tx2` to a slot
    // that is no longer there. Skipping it keeps the map to colours that
    // exist, so the pin either finds a real value or leaves the scheme colour
    // alone.
    const pkg = await Pkg.open(
      await makeDeck([{ paragraphs: [["One"]] }], { scheme: { dk1: "000000", lt1: "FFFFFF" } }),
    );
    const master = await pkg.doc("ppt/slideMasters/slideMaster1.xml");
    const map = master.getElementsByTagName("p:clrMap")[0] as Element;
    map.setAttribute("hlink", "accent6");
    const colours = await coloursOf(pkg, "ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml");
    expect(colours["hlink"]).toBeUndefined();
    // And a name mapped to a slot that IS there still resolves, so this is the
    // one attribute being skipped rather than the loop giving up.
    expect(colours["tx1"]).toBe("000000");
  });
});

describe("the harvest carrying that map", () => {
  /** A heading and one element per entry in `on`, each on the master it names. */
  const library = (extra: Parameters<typeof makeDeck>[1] = {}, on: (1 | 2)[] = [1]): Promise<Uint8Array> =>
    makeDeck(
      [
        { paragraphs: [], title: "Kasser", noBody: true },
        ...on.map((theme, i) => ({
          paragraphs: [["hello"]],
          title: `Kasse ${i + 1}`,
          box: `<a:xfrm><a:off x="1000000" y="1500000"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>`,
          theme,
        })),
      ],
      extra,
    );
  const names: Names = { categories: { Kasser: "Boxes" }, names: { "Kasse 1": "Box 1", "Kasse 2": "Box 2" } };

  it("puts the library theme's colours on the catalogue", async () => {
    const pkg = await Pkg.open(await library({ scheme: OFFICE }));
    const { catalogue } = await harvest(pkg, { size: "16:9", names });
    expect(catalogue.theme).toMatchObject({ accent1: "5B9BD5", tx1: "000000" });
  });

  it("refuses a deck whose elements are spread over two themes", async () => {
    // One colour map per size, so a deck with two themes has no single right
    // answer to pin to — and picking one would pin half the library to colours
    // no slide in it ever had, with nothing downstream able to tell.
    const pkg = await Pkg.open(await library({ scheme: OFFICE, second: { scheme: { accent1: "C83A3A" } } }, [1, 2]));
    const failure = await harvest(pkg, { size: "16:9", names }).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(HarvestError);
    expect((failure as HarvestError).problems.join("\n")).toContain("spread over 2 themes");
  });
});

describe("pinning an element's scheme colours", () => {
  const theme = { accent1: "5B9BD5", tx1: "000000" };

  it("rewrites a scheme colour to the value the library gave it", () => {
    const shape = fragment(`<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`);
    expect(pinSchemeColours(shape, theme)).toEqual({ pinned: 1, left: 0 });
    expect(serializeXml(shape.ownerDocument)).toContain(`<a:srgbClr val="5B9BD5"/>`);
    expect(serializeXml(shape.ownerDocument)).not.toContain("schemeClr");
  });

  it("carries the transforms across, because a pinned colour that loses its shade is wrong twice", () => {
    // `<a:lumMod>` is as legal a child of srgbClr as of schemeClr — the same
    // colour-transform group — so the shade survives the pin. Dropping it would
    // turn a 75% tint of the text colour into the text colour.
    const shape = fragment(
      `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:alpha val="50000"/></a:schemeClr></a:solidFill>`,
    );
    expect(pinSchemeColours(shape, theme).pinned).toBe(1);
    const xml = serializeXml(shape.ownerDocument);
    expect(xml).toContain(`<a:srgbClr val="000000"><a:lumMod val="75000"/><a:alpha val="50000"/></a:srgbClr>`);
  });

  it("leaves phClr alone, and a name this theme does not carry", () => {
    // phClr is not a theme slot: it is the colour a style is being handed, and
    // there is nothing here to resolve it to. A name the theme has no colour
    // for is left as a scheme colour, so it still draws — following the
    // destination, as it always did. Better a colour that moves than one
    // invented here.
    const shape = fragment(
      `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln><a:schemeClr val="accent6"/></a:ln>`,
    );
    expect(pinSchemeColours(shape, theme)).toEqual({ pinned: 0, left: 2 });
    expect(serializeXml(shape.ownerDocument)).toContain(`val="phClr"`);
  });

  it("reaches a scheme colour nested inside one it left alone", () => {
    const shape = fragment(
      `<a:solidFill><a:schemeClr val="phClr"><a:schemeClr val="tx1"/></a:schemeClr></a:solidFill>`,
    );
    expect(pinSchemeColours(shape, theme)).toEqual({ pinned: 1, left: 1 });
  });

  it("pins every one of them, not the first", () => {
    const shape = fragment(
      `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:ln><a:schemeClr val="accent1"/></a:ln>` +
        `<a:effectLst><a:schemeClr val="tx1"/></a:effectLst>`,
    );
    // The walk replaces nodes as it goes; a live childNodes read mid-walk skips
    // the sibling after every hit, which is a pass that silently does half.
    expect(pinSchemeColours(shape, theme).pinned).toBe(3);
    expect(serializeXml(shape.ownerDocument)).not.toContain("schemeClr");
  });

  it("hands back a carried part's text untouched when there is nothing to pin", () => {
    const xml = `<?xml version="1.0"?>\r\n<c:chartSpace xmlns:c="urn:x"><c:plotArea/></c:chartSpace>`;
    const done = pinColoursInXml(xml, theme);
    expect(done.xml).toBe(xml);
    expect(done.result).toEqual({ pinned: 0, left: 0 });
  });

  it("pins the colours inside a carried part", () => {
    const xml =
      `<c:chartSpace xmlns:c="urn:x" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<c:ser><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:ser></c:chartSpace>`;
    const done = pinColoursInXml(xml, theme);
    expect(done.result.pinned).toBe(1);
    expect(done.xml).toContain(`<a:srgbClr val="5B9BD5"/>`);
  });

  it("hands a carried part's text back UNCHANGED when it has scheme colours but none this theme knows", () => {
    // The case between the two above, and the one that decides whether a part
    // is copied byte for byte or round-tripped through a serialiser. It HAS
    // `schemeClr`, so the cheap early exit does not fire; nothing is pinned, so
    // the part must still come back as itself rather than as xmldom's idea of
    // itself — which differs in attribute order and self-closing tags, and
    // would rewrite every chart in the library on a setting that changed
    // nothing.
    const xml =
      `<c:chartSpace xmlns:c="urn:x" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<c:ser><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></c:ser></c:chartSpace>`;
    const done = pinColoursInXml(xml, theme);
    expect(done.result).toEqual({ pinned: 0, left: 1 });
    expect(done.xml).toBe(xml);
  });

  it("counts a scheme colour with no val at all as one it cannot resolve", () => {
    // Not something PowerPoint writes, and something a hand-edited or
    // third-party part can carry. The rule is the same as for a name the theme
    // does not know: leave it, and say so in `left`. Inventing a colour for it
    // would change how somebody's deck draws.
    const shape = fragment(`<a:solidFill><a:schemeClr/></a:solidFill>`);
    expect(pinSchemeColours(shape, theme)).toEqual({ pinned: 0, left: 1 });
    expect(serializeXml(shape.ownerDocument)).toContain("schemeClr");
  });

  it("takes a whole DOCUMENT as well as one shape, which is how a carried part arrives", () => {
    // Both arms of the same entry point: the splice hands it an element out of
    // a parsed fragment, and `pinColoursInXml` hands it a Document. A version
    // that only walked an Element pinned the slide's shapes and silently left
    // the chart's seventeen alone.
    const doc = parseXml(
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:theme>`,
    );
    expect(pinSchemeColours(doc, theme)).toEqual({ pinned: 1, left: 0 });
    expect(serializeXml(doc)).toContain(`<a:srgbClr val="5B9BD5"/>`);
  });
});
