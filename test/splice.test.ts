import { describe, expect, it } from "vitest";
import "./setup-base64.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { P_NS, element, elements, parseXml, serializeXml } from "../src/core/pptx/xml.js";
import {
  freeName,
  fromBase64,
  highestShapeId,
  ownerOfRels,
  placeShape,
  relativeTo,
  renumberShapeIds,
  rewriteRelIds,
  splice,
  withNamespaces,
} from "../src/core/element/splice.js";
import { verify } from "../src/core/element/verify.js";
import type { ElementPayload } from "../src/core/catalogue/types.js";
import { PNG_BASE64, buildDeck } from "./fixtures/deck.js";

const SP = `<p:sp xmlns:p="${P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="7" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="600000"/><a:ext cx="2000000" cy="800000"/></a:xfrm></p:spPr></p:sp>`;

function payload(over: Partial<ElementPayload> = {}): ElementPayload {
  return {
    version: 1,
    id: "test--box",
    shapes: [SP],
    rels: [],
    parts: [],
    bounds: { x: 500000, y: 600000, cx: 2000000, cy: 800000 },
    ...over,
  };
}

describe("path arithmetic", () => {
  it("expresses a target relative to the part that names it", () => {
    expect(relativeTo("ppt/slides/slide1.xml", "ppt/media/image3.png")).toBe("../media/image3.png");
    expect(relativeTo("ppt/charts/chart1.xml", "ppt/charts/style1.xml")).toBe("style1.xml");
  });

  it("names the part a .rels describes", () => {
    expect(ownerOfRels("ppt/charts/_rels/chart1.xml.rels")).toBe("ppt/charts/chart1.xml");
  });

  it("round-trips base64 through the DOM globals the pane uses", () => {
    expect(fromBase64(PNG_BASE64).length).toBeGreaterThan(60);
  });
});

describe("shape ids", () => {
  it("gives every distinct id a fresh one, from the given start", () => {
    const doc = parseXml(SP);
    const next = renumberShapeIds(doc.documentElement, 40);
    expect(elements(doc, P_NS, "cNvPr")[0]!.getAttribute("id")).toBe("40");
    expect(next).toBe(41);
  });

  /**
   * The AlternateContent rule, proven rather than asserted.
   *
   * Both branches of an `<mc:AlternateContent>` carry the same `<p:cNvPr id>`
   * deliberately, because only one is ever live. A renumbering that treats them
   * as two shapes splits one shape in two — and the shipped library's cover
   * slide is exactly this case, so getting it wrong is wrong on the first thing
   * anybody inserts onto.
   */
  it("keeps both branches of an AlternateContent on one id", () => {
    const alt = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="${P_NS}"><mc:Choice Requires="cx1"><p:sp><p:nvSpPr><p:cNvPr id="9" name="A"/></p:nvSpPr></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="9" name="A"/></p:nvSpPr></p:sp></mc:Fallback></mc:AlternateContent>`;
    const doc = parseXml(alt);
    const next = renumberShapeIds(doc.documentElement, 100);
    const ids = elements(doc, P_NS, "cNvPr").map((n) => n.getAttribute("id"));
    expect(ids).toEqual(["100", "100"]);
    expect(next).toBe(101);
  });

  it("reads the highest id already on a slide", async () => {
    const pkg = await Pkg.open(await buildDeck());
    const tree = element(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "spTree")!;
    expect(highestShapeId(tree)).toBe(2);
  });
});

describe("relationship rewriting", () => {
  it("maps every r:-namespaced attribute, not just r:embed", () => {
    const xml = `<p:pic xmlns:p="${P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="rId1" r:link="rId2"/></p:pic>`;
    const doc = parseXml(xml);
    rewriteRelIds(
      doc.documentElement,
      new Map([
        ["rId1", "rId9"],
        ["rId2", "rId10"],
      ]),
    );
    const out = serializeXml(doc.documentElement);
    expect(out).toContain('r:embed="rId9"');
    expect(out).toContain('r:link="rId10"');
  });

  /**
   * The invariant, not the plan. A slide naming a relationship its .rels does
   * not define is what PowerPoint calls a damaged file; the shipped library
   * produced 41 of those on one slide before shape tags were carried.
   */
  it("removes a reference nothing carried, taking an empty <p:tags> with it", () => {
    const xml = `<p:sp xmlns:p="${P_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvSpPr><p:nvPr><p:custDataLst><p:tags r:id="rId7"/></p:custDataLst></p:nvPr></p:nvSpPr></p:sp>`;
    const doc = parseXml(xml);
    const dropped = rewriteRelIds(doc.documentElement, new Map());
    expect(dropped).toEqual(["rId7"]);
    expect(serializeXml(doc.documentElement)).not.toContain("rId7");
    expect(serializeXml(doc.documentElement)).not.toContain("<p:tags");
  });
});

describe("placement", () => {
  it("moves a shape by the requested delta", () => {
    const doc = parseXml(SP);
    placeShape(doc.documentElement, 100000, 200000, 1, { x: 500000, y: 600000, cx: 1, cy: 1 });
    expect(serializeXml(doc.documentElement)).toContain('x="600000"');
    expect(serializeXml(doc.documentElement)).toContain('y="800000"');
  });

  it("scales about the element's own origin", () => {
    const doc = parseXml(SP);
    placeShape(doc.documentElement, 0, 0, 0.5, { x: 500000, y: 600000, cx: 2000000, cy: 800000 });
    const out = serializeXml(doc.documentElement);
    // The origin shape stays put and halves.
    expect(out).toContain('x="500000"');
    expect(out).toContain('cx="1000000"');
  });
});

describe("namespace repair", () => {
  it("adds only the prefixes a fragment does not already declare", () => {
    const out = withNamespaces(SP);
    // `p` and `a` were on the fragment already and must not be repeated — a
    // duplicate xmlns for the same prefix is not well-formed.
    expect(out.match(/xmlns:p=/g)).toHaveLength(1);
    expect(out.match(/xmlns:a=/g)).toHaveLength(1);
    // `r` and `mc` were not, and a harvested shape can name either.
    expect(out).toContain("xmlns:r=");
    expect(out).toContain("xmlns:mc=");
    expect(() => parseXml(out)).not.toThrow();
  });

  it("is a no-op when every prefix is already declared", () => {
    const complete = withNamespaces(SP);
    expect(withNamespaces(complete)).toBe(complete);
  });

  it("does not produce `<p:sp/ xmlns…>` on a self-closing root", () => {
    const out = withNamespaces("<p:sp/>");
    expect(out.endsWith("/>")).toBe(true);
    expect(() => parseXml(out)).not.toThrow();
  });
});

describe("splice", () => {
  it("adds the element's shapes and leaves a slide that verifies", async () => {
    const pkg = await Pkg.open(await buildDeck());
    const report = await splice(pkg, "ppt/slides/slide1.xml", payload(), { x: 0, y: 0 });
    expect(report.shapes).toBe(1);
    expect(await verify(pkg, "ppt/slides/slide1.xml")).toEqual([]);
  });

  it("never reuses a shape id already on the target slide", async () => {
    const pkg = await Pkg.open(await buildDeck());
    // The fixture's own rectangle is id 2; the incoming shape is id 7.
    await splice(pkg, "ppt/slides/slide1.xml", payload(), { x: 0, y: 0 });
    const tree = element(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "spTree")!;
    const ids = elements(tree, P_NS, "cNvPr").map((n) => n.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("renames an incoming part that collides, and points the slide at the new name", async () => {
    const pkg = await Pkg.open(await buildDeck());
    pkg.setBytes("ppt/media/image1.png", fromBase64(PNG_BASE64));
    const pic = `<p:pic xmlns:p="${P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvPicPr><p:cNvPr id="3" name="Pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>`;
    const report = await splice(
      pkg,
      "ppt/slides/slide1.xml",
      payload({
        shapes: [pic],
        rels: [
          {
            rId: "rId5",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            path: "ppt/media/image1.png",
          },
        ],
        parts: [{ path: "ppt/media/image1.png", contentType: "image/png", base64: PNG_BASE64 }],
        bounds: { x: 0, y: 0, cx: 100, cy: 100 },
      }),
      { x: 0, y: 0 },
    );
    // The existing image1.png is untouched; the incoming one got a free name.
    expect(report.parts).toEqual(["ppt/media/image2.png"]);
    expect(await verify(pkg, "ppt/slides/slide1.xml")).toEqual([]);
  });

  it("keeps only the slide it was asked to keep", async () => {
    const pkg = await Pkg.open(await buildDeck());
    expect((await pkg.slidePaths()).length).toBe(2);
    await pkg.keepOnlySlide("ppt/slides/slide2.xml");
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
  });

  it("refuses to keep a slide that is not in the deck", async () => {
    const pkg = await Pkg.open(await buildDeck());
    await expect(pkg.keepOnlySlide("ppt/slides/slide9.xml")).rejects.toThrow(/not in this presentation/);
  });
});

describe("free names", () => {
  it("keeps the family and the extension", async () => {
    const pkg = await Pkg.open(await buildDeck());
    pkg.setBytes("ppt/media/image1.png", new Uint8Array([1]));
    pkg.setBytes("ppt/media/image7.png", new Uint8Array([1]));
    expect(freeName(pkg, "ppt/media/image3.png")).toBe("ppt/media/image8.png");
  });

  it("numbers a name that had no number", async () => {
    const pkg = await Pkg.open(await buildDeck());
    expect(freeName(pkg, "ppt/embeddings/Workbook.xlsx")).toBe("ppt/embeddings/Workbook1.xlsx");
  });
});
