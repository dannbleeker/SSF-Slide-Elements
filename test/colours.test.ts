import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { HarvestError, Pkg, harvest, parseXml, serializeXml } from "../src/core/index.js";
import type { Names } from "../src/core/index.js";
import { coloursOf, themeColoursFor } from "../src/core/pptx/theme.js";
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
});
