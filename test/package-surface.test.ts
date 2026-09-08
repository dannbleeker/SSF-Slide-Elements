import { describe, expect, it } from "vitest";
import { Pkg, extensionOf } from "../src/core/pptx/pkg.js";
import {
  A_NS,
  P_NS,
  R_NS,
  child,
  children,
  element,
  elements,
  parseXml,
  relationshipIdsIn,
} from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The rest of the package layer's surface, on bytes the fixture built.
 *
 * SSF-Merge exercised these through its clone and its merge run. Neither is
 * here yet, and a method nothing calls is a method whose next caller finds the
 * bug — so each is asked directly, for the answer the splice will rely on.
 */

async function deck(...args: Parameters<typeof makeDeck>): Promise<Pkg> {
  return Pkg.open(await makeDeck(...args));
}

const SLIDE = "ppt/slides/slide1.xml";
const TYPES = "[Content_Types].xml";
const PRESENTATION = "ppt/presentation.xml";

describe("reading and writing raw parts", () => {
  it("hands back a binary part's bytes untouched", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], icons: true }]);
    const png = await pkg.bytes("ppt/media/icon1.png");
    expect(Array.from(png.slice(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47]);
    // And a part written as bytes comes back as the same bytes, through the
    // zip, not re-encoded as text.
    pkg.setBytes("ppt/media/image7.png", new Uint8Array([0, 255, 128]));
    const again = await Pkg.open(await pkg.toBytes());
    expect(Array.from(await again.bytes("ppt/media/image7.png"))).toEqual([0, 255, 128]);
  });

  it("refuses to read a part that is not there, naming it", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    await expect(pkg.bytes("ppt/media/nothing.png")).rejects.toThrow(/no part "ppt\/media\/nothing.png"/);
    await expect(pkg.text("ppt/slides/slide9.xml")).rejects.toThrow(/no part/);
    await expect(pkg.copyPart("ppt/slides/slide9.xml", "ppt/slides/slide10.xml")).rejects.toThrow(/cannot copy/);
  });

  it("lists every part and no directory entry", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const names = pkg.partNames();
    expect(names).toContain(TYPES);
    expect(names).toContain("ppt/slides/slide2.xml");
    expect(names.some((n) => n.endsWith("/"))).toBe(false);
  });
});

describe("copying a part", () => {
  it("copies the EDITED version, never the bytes on disk", async () => {
    // Cloning a slide whose text had already been changed would otherwise copy
    // the version the package was opened with, silently.
    const pkg = await deck([{ paragraphs: [["before"]] }]);
    const doc = await pkg.doc(SLIDE);
    const t = element(doc, A_NS, "t");
    if (!t) throw new Error("the fixture changed shape");
    t.textContent = "after";
    await pkg.copyPart(SLIDE, "ppt/slides/slide2.xml");
    expect(await pkg.text("ppt/slides/slide2.xml")).toContain("after");
    expect(await pkg.text("ppt/slides/slide2.xml")).not.toContain("before");
  });

  it("gives the copy its own numbers, never one the package holds", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], notes: true }, { paragraphs: [["b"]] }]);
    expect(pkg.nextSlideNumber()).toBe(3);
    // Notes are numbered on their own: one slide has notes, so the next free
    // notes number is 2 whatever the slide count says.
    expect(pkg.nextNotesNumber()).toBe(2);
    await pkg.copyPart("ppt/notesSlides/notesSlide1.xml", "ppt/notesSlides/notesSlide2.xml");
    expect(pkg.nextNotesNumber()).toBe(3);
  });
});

describe("content types", () => {
  it("answers a part's own override before its extension's default", async () => {
    const pkg = await deck([{ paragraphs: [["a"]], icons: true }]);
    expect(await pkg.contentTypeOf(SLIDE)).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    expect(await pkg.contentTypeOf("ppt/media/icon1.png")).toBe("image/png");
    // Nothing covers it: no override, no default for the extension, or no
    // extension at all.
    expect(await pkg.contentTypeOf("ppt/embeddings/oleObject1.bin")).toBeUndefined();
    expect(await pkg.contentTypeOf("ppt/embeddings/workbook")).toBeUndefined();
  });

  it("declares an extension once, ahead of the overrides", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    await pkg.addContentTypeDefault("jpeg", "image/jpeg");
    // Case-insensitive: a second entry for the same extension is schema-invalid.
    await pkg.addContentTypeDefault("JPEG", "image/jpeg");
    const types = await pkg.doc(TYPES);
    const defaults = elements(types, "http://schemas.openxmlformats.org/package/2006/content-types", "Default");
    expect(defaults.filter((d) => d.getAttribute("Extension")?.toLowerCase() === "jpeg")).toHaveLength(1);
    expect(types.documentElement.firstChild).toBe(defaults[0]);
    expect(await pkg.contentTypeOf("ppt/media/photo.JPEG")).toBe("image/jpeg");
  });
});

describe("the slide id list", () => {
  it("appends a slide with the highest id plus one, in the format's range", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const rId = await pkg.addRel(
      PRESENTATION,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      "slides/slide3.xml",
    );
    await pkg.copyPart(SLIDE, "ppt/slides/slide3.xml");
    const first = await pkg.appendSldId(rId);
    expect(first).toBe(258);
    // Highest plus one again, and a removal does not lower it: a reused slide
    // id is a duplicate.
    await pkg.removeSlide("ppt/slides/slide2.xml");
    const second = await pkg.appendSldId(rId);
    expect(second).toBe(259);
    expect(await pkg.slidePaths()).toEqual([SLIDE, "ppt/slides/slide3.xml", "ppt/slides/slide3.xml"]);
  });

  it("refuses a deck that has run out of ids, and one with no list at all", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const pres = await pkg.doc(PRESENTATION);
    const sldId = element(pres, P_NS, "sldId");
    if (!sldId) throw new Error("the fixture changed shape");
    sldId.setAttribute("id", "2147483647");
    await expect(pkg.appendSldId("rId2")).rejects.toThrow(/run out of slide ids/);

    const bare = await deck([{ paragraphs: [["a"]] }]);
    const list = element(await bare.doc(PRESENTATION), P_NS, "sldIdLst");
    list?.parentNode?.removeChild(list);
    await expect(bare.appendSldId("rId2")).rejects.toThrow(/no <p:sldIdLst>/);
    expect(await bare.slidePaths()).toEqual([]);
  });

  it("skips a listed slide whose relationship is missing", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const rels = await pkg.doc(Pkg.relsPathFor(PRESENTATION));
    for (const rel of elements(rels, "http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")) {
      if (rel.getAttribute("Target") === "slides/slide2.xml") rel.parentNode?.removeChild(rel);
    }
    expect(await pkg.slidePaths()).toEqual([SLIDE]);
  });
});

describe("relationships, one hop out", () => {
  it("answers nothing for a part with no relationships and for an id it does not have", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    expect(await pkg.relatedParts("ppt/theme/theme1.xml")).toEqual([]);
    expect(await pkg.relTarget("ppt/theme/theme1.xml", "rId1")).toBeUndefined();
    expect(await pkg.relTarget(SLIDE, "rId99")).toBeUndefined();
  });

  it("skips an external target, which is a URL and not a part", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const rels = await pkg.doc(Pkg.relsPathFor(SLIDE));
    const link = rels.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
    link.setAttribute("Id", "rId7");
    link.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
    link.setAttribute("Target", "https://example.invalid/");
    link.setAttribute("TargetMode", "External");
    rels.documentElement.appendChild(link);
    expect(await pkg.relatedParts(SLIDE)).toEqual(["ppt/slideLayouts/slideLayout1.xml"]);
  });
});

describe("holding parsed parts", () => {
  it("peeks at a part without keeping it, and reads a held part as held", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const before = pkg.cachedParts();
    const count = await pkg.peek("ppt/slides/slide2.xml", (d) => elements(d, A_NS, "t").length);
    expect(count).toBe(1);
    expect(pkg.cachedParts(), "peek retained the part").toBe(before);
    // A part somebody has already parsed stays parsed, and peek sees its edit.
    const doc = await pkg.doc(SLIDE);
    const t = element(doc, A_NS, "t");
    if (!t) throw new Error("the fixture changed shape");
    t.textContent = "edited";
    expect(await pkg.peek(SLIDE, (d) => element(d, A_NS, "t")?.textContent)).toBe("edited");
  });

  it("releases a part back to the zip with its edit, and forgets it", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const doc = await pkg.doc(SLIDE);
    const t = element(doc, A_NS, "t");
    if (!t) throw new Error("the fixture changed shape");
    t.textContent = "kept";
    expect(pkg.cachedParts()).toBe(1);
    pkg.release(SLIDE);
    expect(pkg.cachedParts()).toBe(0);
    // Releasing a part that was never held is nothing.
    pkg.release("ppt/slides/slide1.xml");
    const again = await Pkg.open(await pkg.toBytes());
    expect(await again.text(SLIDE)).toContain("kept");
  });
});

describe("a part's extension", () => {
  it("reads the last segment, never the whole path", () => {
    expect(extensionOf("ppt/media/image1.png")).toBe("png");
    expect(extensionOf("ppt/embeddings/workbook")).toBe("");
    expect(extensionOf("ppt/my.charts/workbook")).toBe("");
    expect(extensionOf("[Content_Types].xml")).toBe("xml");
  });
});

describe("the XML helpers", () => {
  it("tells a direct child from a descendant", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }]);
    const doc = await pkg.doc(SLIDE);
    const cSld = element(doc, P_NS, "cSld");
    if (!cSld) throw new Error("the fixture changed shape");
    // `element` walks descendants and finds the shape; `child` asks what
    // `<p:cSld>` itself owns, which is the tree, not the shape.
    expect(element(cSld, P_NS, "sp")).toBeDefined();
    expect(child(cSld, P_NS, "sp")).toBeUndefined();
    expect(child(cSld, P_NS, "spTree")).toBeDefined();
    expect(children(cSld, P_NS, "spTree")).toHaveLength(1);
    expect(children(cSld, A_NS, "spTree")).toEqual([]);
  });

  it("finds every relationship id a part names, wherever it names it", async () => {
    // An icon's SVG companion (`asvg:svgBlip r:embed`), a linked picture
    // (`a:blip r:link`) and a shape's tag reference are all references, and none
    // of them is an `r:embed` on an `a:blip`.
    const pkg = await deck([{ paragraphs: [["a"]], icons: true, shapeTags: true }]);
    const ids = relationshipIdsIn(await pkg.doc(SLIDE));
    expect([...ids].sort()).toEqual(["rId20", "rId21", "rId22", "rId30"]);
    // By prefix as well as by namespace, which is the conservative direction:
    // an `r:` attribute bound to some other namespace is still counted, since
    // keeping a relationship that could have gone is cheaper than dropping one
    // the markup still names.
    const bare = parseXml('<p:sld xmlns:p="urn:p" xmlns:r="urn:other"><p:pic r:embed="rId5"/></p:sld>');
    expect([...relationshipIdsIn(bare)]).toEqual(["rId5"]);
    expect(relationshipIdsIn(parseXml("<r/>")).size).toBe(0);
    expect(R_NS).toContain("relationships");
  });
});
