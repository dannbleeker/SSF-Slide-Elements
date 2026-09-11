import { readFileSync } from "node:fs";
import { beforeAll, describe, expect, it } from "vitest";
import { Pkg, harvest } from "../src/core/index.js";
import type { Element as CatalogueElement, Harvest, Names } from "../src/core/index.js";
import { REL_TYPE } from "../src/core/pptx/parts.js";
import { PKG_REL_NS, elements } from "../src/core/pptx/xml.js";
import { carry, freeName, targetFrom, type PartStore } from "../src/core/splice/carry.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The parts an element brings with it, arriving in somebody else's package.
 *
 * `carry.ts` is the one step of the splice whose failures are all SILENT. A
 * picture copied under a name the deck already holds overwrites the deck's own
 * and nothing complains; a chart copied without rewriting its rels points at a
 * workbook that is not in the package at all; a part copied twice doubles a
 * deck's size with no visible difference. Every one of those leaves a package
 * that is structurally valid, which is why the assertions here are on the
 * PACKAGE that comes out — the names it now holds, what its content types
 * declare, and what each relationship `Target` actually resolves to — rather
 * than on what the function returned.
 *
 * Two halves, the way `test/splice.test.ts` and `test/catalogue.test.ts` are
 * built. The synthetic cases drive a real in-memory .pptx from
 * `test/fixtures/deck.ts` against a fake store — a Map lookup, which is the
 * exact shape the pane has, one fetch per part — so that each rule can be put
 * under pressure on its own. The last block runs the real library's chart
 * through the same code, so the synthetic ones are not proving something about
 * XML no producer writes: the library's chart really does reach
 * `../embeddings/Microsoft_Excel_Worksheet.xlsx` through its own rels part, and
 * that name really does have no digits to extend.
 */

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

/** The part that owns the new relationships, everywhere below: the slide the element lands on. */
const SLIDE = "ppt/slides/slide1.xml";

const IMAGE = REL_TYPE.image;
/**
 * Not in `REL_TYPE`, deliberately: nothing in today's library carries a
 * hyperlink, which is exactly why the external path needs a test rather than a
 * reader's confidence.
 */
const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const OLE_TYPE = "application/vnd.openxmlformats-officedocument.oleObject";

/** Bytes that are not valid UTF-8, so a part written as TEXT comes back mangled. */
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0xff, 0xfe, 0x00, 0x01]);

/** A store the pane will be: parts by the path the library knew them by, and a note of every ask. */
interface Store {
  (path: string): Promise<Uint8Array | string | undefined>;
  /** Every path asked for, in order. A part copied twice is asked for twice. */
  asked: string[];
}

function storeOf(parts: Record<string, string | Uint8Array>): Store {
  const held = new Map(Object.entries(parts));
  const asked: string[] = [];
  const store: PartStore = (path) => {
    asked.push(path);
    return Promise.resolve(held.get(path));
  };
  return Object.assign(store, { asked });
}

/** One of the element's relationships, as the catalogue records them. */
function rel(
  id: string,
  type: string,
  target: string,
  external = false,
): {
  id: string;
  type: string;
  target: string;
  external: boolean;
} {
  return { id, type, target, external };
}

/** A `.rels` part, as the library holds one beside a carried part. */
function relsXml(rels: { id: string; type: string; target: string; external?: boolean }[]): string {
  const written = rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"` +
        `${r.external === true ? ' TargetMode="External"' : ""}/>`,
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">${written}</Relationships>`;
}

/** What a part's relationships say, in the destination package. */
async function relsOf(pkg: Pkg, part: string): Promise<{ id: string; type: string; target: string; mode: string }[]> {
  const doc = await pkg.doc(Pkg.relsPathFor(part));
  return elements(doc, PKG_REL_NS, "Relationship").map((r) => ({
    id: r.getAttribute("Id") ?? "",
    type: r.getAttribute("Type") ?? "",
    target: r.getAttribute("Target") ?? "",
    mode: r.getAttribute("TargetMode") ?? "",
  }));
}

/** A plain destination: one slide, nothing carried, so a copy's name is the module's alone. */
async function plain(): Promise<Pkg> {
  return Pkg.open(await makeDeck([{ paragraphs: [["First"]] }]));
}

/**
 * A destination that already holds a picture, a chart and a tag part.
 *
 * The names a free name has to avoid: `ppt/media/icon1.png`,
 * `ppt/charts/chart1.xml` and `ppt/tags/tag9.xml`. A test against an EMPTY deck
 * cannot tell "the next free number" from "the number the library used".
 */
async function crowded(): Promise<Pkg> {
  return Pkg.open(
    await makeDeck([{ paragraphs: [["First"]], icons: true, chart: "the destination's own", shapeTags: true }]),
  );
}

describe("the target one part writes for another", () => {
  it("spells a picture from the slide the way PowerPoint spells it", () => {
    // `../media/image3.png`, not `/ppt/media/image3.png`. Both are legal OPC
    // and PowerPoint opens either, so nothing downstream can catch a change
    // here — only the string can.
    expect(targetFrom(SLIDE, "ppt/media/image3.png")).toBe("../media/image3.png");
  });

  it("names a neighbour in the same directory without climbing to the root and back", () => {
    // A chart's styling parts sit beside it. Counting one `..` per segment of
    // the owner's path instead of per segment they do NOT share writes
    // `../../ppt/charts/colors1.xml`, which resolves to the same part — so a
    // resolver-based check passes and the package is written two ways at once.
    expect(targetFrom("ppt/charts/chart2.xml", "ppt/charts/colors1.xml")).toBe("colors1.xml");
  });

  it("climbs out of ppt/ for a part that lives outside it", () => {
    expect(targetFrom(SLIDE, "docProps/thumbnail.jpeg")).toBe("../../docProps/thumbnail.jpeg");
  });

  it("writes one spelling for every shape of pair the splice produces", () => {
    const pairs = [
      [SLIDE, "ppt/media/image3.png", "../media/image3.png"],
      [SLIDE, "ppt/tags/tag2.xml", "../tags/tag2.xml"],
      [SLIDE, "docProps/thumbnail.jpeg", "../../docProps/thumbnail.jpeg"],
      ["ppt/charts/chart1.xml", "ppt/embeddings/Worksheet1.xlsx", "../embeddings/Worksheet1.xlsx"],
      ["ppt/charts/chart1.xml", "ppt/charts/colors1.xml", "colors1.xml"],
      // The owner's directory CONTAINS the target's, which is the one case the
      // walk ends on the target running out rather than the owner: an
      // off-by-one there writes `../../ppt/somewhere.xml` or `somewhere.xml`,
      // and one of those names a part that is not the one asked for.
      [SLIDE, "ppt/somewhere.xml", "../somewhere.xml"],
      ["ppt/presentation.xml", "ppt/media/image3.png", "media/image3.png"],
    ] as const;
    for (const [owner, part, spelled] of pairs) {
      expect(targetFrom(owner, part), `${owner} → ${part}`).toBe(spelled);
    }
  });

  it("agrees with the package's own resolver, which is the reader at the other end", async () => {
    // `targetFrom` deliberately does not share an implementation with
    // `resolveTarget` in `pkg.ts` — the module comment says why — and two
    // implementations of the same arithmetic drift. Agreement over the pairs
    // the splice really produces is what keeps them honest, and it is a
    // different question from the spelling: several wrong spellings resolve.
    const pkg = await plain();
    const pairs = [
      [SLIDE, "ppt/media/image3.png"],
      [SLIDE, "ppt/tags/tag2.xml"],
      [SLIDE, "docProps/thumbnail.jpeg"],
      ["ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"],
      ["ppt/charts/chart1.xml", "ppt/charts/colors1.xml"],
    ] as const;
    for (const [owner, part] of pairs) {
      expect(pkg.resolved(owner, targetFrom(owner, part)), `${owner} → ${part}`).toBe(part);
    }
  });
});

describe("a name the destination is not already using", () => {
  it("extends the destination's own numbering rather than reusing a name it holds", async () => {
    const pkg = await crowded();
    // The library and the deck both call a picture `icon1.png`. Answering the
    // library's name overwrites the DECK's picture, and the package stays
    // valid: every relationship resolves, every content type is declared, and
    // the user's own image has silently become the element's.
    expect(pkg.has("ppt/media/icon1.png")).toBe(true);
    const name = freeName(pkg, "ppt/media/icon1.png");
    expect(name).toBe("ppt/media/icon2.png");
    expect(pkg.has(name)).toBe(false);
  });

  it("gives a part with no digits in its name one", async () => {
    // The library's workbook, which is `Microsoft_Excel_Worksheet.xlsx` with no
    // number at all. Stripping "trailing digits" from a name that has none must
    // still produce a new name rather than the same one back.
    const pkg = await plain();
    expect(freeName(pkg, "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx")).toBe(
      "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    );
  });

  it("counts each extension as its own family, so an .emf does not inherit the .png numbering", async () => {
    // `image1.png` and `image1.emf` are two different parts and the package can
    // hold both. A counter shared across extensions would skip a name for no
    // reason, and — worse in the other direction — a counter that ignored the
    // extension would hand `image1.emf` back for a deck holding exactly that.
    const pkg = await plain();
    pkg.setBytes("ppt/media/image1.png", PNG);
    expect(freeName(pkg, "ppt/media/image7.emf")).toBe("ppt/media/image1.emf");
    expect(freeName(pkg, "ppt/media/image7.png")).toBe("ppt/media/image2.png");
  });

  it("treats a dot in a directory as no extension at all", async () => {
    // `lastIndexOf(".")` finds the dot in the DIRECTORY here, and a naive
    // `split(".").pop()` would call `parts/tag` an extension of `parts/tag` —
    // producing `ppt/my1.parts/tag`, a part in a directory the package does not
    // have. The guard is the comparison against the last slash.
    const pkg = await plain();
    expect(freeName(pkg, "ppt/my.parts/tag")).toBe("ppt/my.parts/tag1");
    expect(freeName(pkg, "ppt/media/thumbnail")).toBe("ppt/media/thumbnail1");
  });

  it("only moves on once the part has been written, which is what keeps two copies apart", async () => {
    /**
     * `freeName` does not RESERVE anything — it reads `Pkg`'s counter, and the
     * counter only learns of a name when the part is written. So asking twice
     * without writing answers the same name twice.
     *
     * That is the coupling `copyPart` depends on, and it is one line away from
     * breaking: move the `setBytes` after the recursion, or skip it on a path
     * that returns early, and two carried parts get one name — the second
     * overwrites the first, and the package is still valid.
     */
    const pkg = await crowded();
    const first = freeName(pkg, "ppt/media/icon1.png");
    expect(freeName(pkg, "ppt/media/icon1.png"), "a name is not reserved by being asked for").toBe(first);
    pkg.setBytes(first, PNG);
    expect(freeName(pkg, "ppt/media/icon1.png")).not.toBe(first);
  });
});

describe("what an element carries into the deck", () => {
  it("copies the part, declares it, and points the slide's new relationship at the copy", async () => {
    const pkg = await crowded();
    const store = storeOf({ "ppt/media/icon1.png": PNG });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", IMAGE, "ppt/media/icon1.png")],
      types: { "ppt/media/icon1.png": "image/png" },
      store,
    });

    // Read from the answer rather than written here. A media part's name is
    // derived from its own bytes now, so that the NEXT insert of the same
    // picture finds it already in the package instead of copying it again —
    // what this case is about is that the copy is there, declared, and pointed
    // at, which is true whatever it ends up called.
    const name = carried.parts.get("ppt/media/icon1.png") ?? "";
    expect(name).toMatch(/^ppt\/media\/[^/]+\.png$/);
    expect(name, "the library's own name would collide in a deck that has one").not.toBe("ppt/media/icon1.png");
    expect(pkg.has(name)).toBe(true);
    const added = (await relsOf(pkg, SLIDE)).find((r) => r.id === carried.ids.get("rId5"));
    expect(added?.type).toBe(IMAGE);
    // Relative to the SLIDE, which is the part that now owns the relationship.
    expect(added?.target).toBe(`../media/${name.split("/").pop() ?? ""}`);
    expect(added?.mode).toBe("");
    expect(pkg.resolved(SLIDE, added?.target ?? "")).toBe(name);
  });

  it("renames every id the element's markup uses, and never twice to the same one", async () => {
    // `repoint` in `shapes.ts` rewrites the markup from this map, so an id
    // missing from it leaves an `r:embed` pointing at whatever the destination
    // slide happens to call `rId5` — its own layout, in a fixture deck.
    const pkg = await plain();
    const store = storeOf({
      "ppt/media/image1.png": PNG,
      "ppt/media/image2.png": PNG,
      "ppt/embeddings/oleObject1.bin": PNG,
    });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [
        rel("rId5", IMAGE, "ppt/media/image1.png"),
        rel("rId6", IMAGE, "ppt/media/image2.png"),
        rel("rId7", REL_TYPE.package, "ppt/embeddings/oleObject1.bin"),
      ],
      types: {},
      store,
    });

    expect([...carried.ids.keys()]).toEqual(["rId5", "rId6", "rId7"]);
    const given = [...carried.ids.values()];
    expect(new Set(given).size, "two of the element's relationships were given one id").toBe(3);
    // The destination's own `rId1` is its layout, and nothing here may take it.
    const slide = await relsOf(pkg, SLIDE);
    expect(slide.filter((r) => given.includes(r.id)).length).toBe(3);
    expect(given).not.toContain("rId1");
  });

  it("writes a picture as bytes, so it survives the zip unchanged", async () => {
    // JSZip stores a string as UTF-8: a PNG handed to `setText` is re-encoded,
    // and the deck opens with a broken picture on every slide it appears on.
    // Only a round trip through the zip catches it — the in-memory part looks
    // fine either way.
    const pkg = await plain();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", IMAGE, "ppt/media/image1.png")],
      types: {},
      store: storeOf({ "ppt/media/image1.png": PNG }),
    });
    const name = carried.parts.get("ppt/media/image1.png") as string;
    const again = await Pkg.open(await pkg.toBytes());
    expect(Array.from(await again.bytes(name))).toEqual(Array.from(PNG));
  });

  it("writes an XML part the store hands over as text", async () => {
    const pkg = await crowded();
    const tag = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="urn:x"><p:tag name="SSF" val="carried"/></p:tagLst>`;
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", REL_TYPE.tags, "ppt/tags/tag1.xml")],
      types: {},
      store: storeOf({ "ppt/tags/tag1.xml": tag }),
    });
    // tag9 is the destination's own, so the copy is tag10 and the deck's
    // vendor tag part is still there and still says what it said.
    expect(carried.parts.get("ppt/tags/tag1.xml")).toBe("ppt/tags/tag10.xml");
    expect(await pkg.text("ppt/tags/tag10.xml")).toContain('val="carried"');
    expect(await pkg.text("ppt/tags/tag9.xml")).toContain("do not delete");
  });

  it("declares an Override even where the destination already has a Default for the extension", async () => {
    /**
     * The module's third rule. The fixture deck declares `.png` with a Default,
     * so `contentTypeOf` answers "image/png" for the copy whether or not an
     * Override was written — which means a change that dropped the Override
     * would pass every content-type check in the suite and only fail on the
     * first extension the destination has no Default for.
     *
     * So the assertion is on the literal entry in `[Content_Types].xml`.
     */
    const pkg = await crowded();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", IMAGE, "ppt/media/icon1.png")],
      types: { "ppt/media/icon1.png": "image/png" },
      store: storeOf({ "ppt/media/icon1.png": PNG }),
    });
    const name = carried.parts.get("ppt/media/icon1.png") ?? "";
    expect(await pkg.text("[Content_Types].xml")).toContain(`PartName="/${name}"`);
  });

  it("declares nothing for a part the catalogue has no content type for", async () => {
    // `.bin` has no Default in the fixture deck, so this is the one place the
    // difference is visible: with the type, the copy is classified; without it,
    // the package has a part PowerPoint cannot classify — which is what the
    // harvest refuses to produce and why `carried` exists at all.
    const declared = await plain();
    const bare = await plain();
    const request = {
      owner: SLIDE,
      rels: [rel("rId5", REL_TYPE.package, "ppt/embeddings/oleObject1.bin")],
      store: storeOf({ "ppt/embeddings/oleObject1.bin": PNG }),
    };
    const withType = await carry({
      ...request,
      pkg: declared,
      types: { "ppt/embeddings/oleObject1.bin": OLE_TYPE },
    });
    const without = await carry({ ...request, pkg: bare, types: {} });

    expect(await declared.contentTypeOf(withType.parts.get("ppt/embeddings/oleObject1.bin") as string)).toBe(OLE_TYPE);
    expect(await bare.contentTypeOf(without.parts.get("ppt/embeddings/oleObject1.bin") as string)).toBeUndefined();
  });

  it("does not copy a picture the package already holds from an earlier insert", async () => {
    /**
     * The second insert of the same element used to copy its picture again.
     * Measured on the validators' deck before this: four inserts of
     * `markeringer-1` left four byte-identical copies of one 29 KB `.emf` and
     * cost 11.6 KB each. A user who stamps thirty slides carried thirty.
     *
     * A media part's name is its own fingerprint and length, so the second
     * carry finds it with one lookup — rather than by decompressing every
     * picture in the user's deck to compare, which on a deck full of
     * photographs is the cost this engine spent a day taking out of the base64
     * path.
     */
    const pkg = await plain();
    const before = pkg.partNames().filter((n) => n.startsWith("ppt/media/")).length;
    const twice = async (): Promise<string> => {
      const store = storeOf({ "ppt/media/image1.png": PNG });
      const carried = await carry({
        pkg,
        owner: SLIDE,
        rels: [rel("rId5", IMAGE, "ppt/media/image1.png")],
        types: { "ppt/media/image1.png": "image/png" },
        store,
      });
      return carried.parts.get("ppt/media/image1.png") ?? "";
    };

    const first = await twice();
    const second = await twice();
    expect(second, "the same bytes were carried under a second name").toBe(first);
    expect(pkg.partNames().filter((n) => n.startsWith("ppt/media/"))).toHaveLength(before + 1);

    // Both inserts point at it, and the package still declares it once.
    const declarations = (await pkg.text("[Content_Types].xml")).split(`PartName="/${first}"`).length - 1;
    expect(declarations, "the content type was declared twice").toBe(1);
  });

  it("carries a DIFFERENT picture separately, whatever it is called", async () => {
    // The other half, and the one that would make the case above vacuous if it
    // broke: sharing by content must not share by extension. Two pictures that
    // differ by a byte are two pictures.
    const pkg = await plain();
    const other = new Uint8Array([...PNG.slice(0, PNG.length - 1), 0x02]);
    const names: string[] = [];
    for (const bytes of [PNG, other]) {
      const carried = await carry({
        pkg,
        owner: SLIDE,
        rels: [rel("rId5", IMAGE, "ppt/media/image1.png")],
        types: { "ppt/media/image1.png": "image/png" },
        store: storeOf({ "ppt/media/image1.png": bytes }),
      });
      names.push(carried.parts.get("ppt/media/image1.png") ?? "");
    }
    expect(names[0]).not.toBe(names[1]);
    expect(pkg.has(names[0] ?? "")).toBe(true);
    expect(pkg.has(names[1] ?? "")).toBe(true);
  });

  it("copies a part named by two relationships once, and points both at it", async () => {
    /**
     * An element that shows one picture twice names it twice. Copying per
     * relationship rather than per part gives the deck two identical pictures
     * under two names — invisible in the slide, and the 16 MB library says what
     * that costs at scale.
     *
     * The store's ask count is the sharp end of it: `copyPart` answers an
     * already-copied part before it fetches, so a second fetch means a second
     * copy is on its way.
     */
    const pkg = await plain();
    const store = storeOf({ "ppt/media/image1.png": PNG });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", IMAGE, "ppt/media/image1.png"), rel("rId6", IMAGE, "ppt/media/image1.png")],
      types: { "ppt/media/image1.png": "image/png" },
      store,
    });

    expect(carried.parts.size).toBe(1);
    expect(store.asked.filter((p) => p === "ppt/media/image1.png")).toHaveLength(1);
    const name = carried.parts.get("ppt/media/image1.png") ?? "";
    expect(pkg.partNames().filter((n) => n.startsWith("ppt/media/"))).toEqual([name]);
    // Two relationships, two ids, one target.
    const ids = [carried.ids.get("rId5"), carried.ids.get("rId6")];
    expect(new Set(ids).size).toBe(2);
    const slide = await relsOf(pkg, SLIDE);
    const target = `../media/${name.split("/").pop() ?? ""}`;
    expect(slide.filter((r) => ids.includes(r.id)).map((r) => r.target)).toEqual([target, target]);
  });
});

describe("the parts a carried part reaches through its own relationships", () => {
  it("copies a chart's rels part and rewrites its target to the workbook's new name", async () => {
    /**
     * The case the module comment records. Copy the chart and not its rels, and
     * PowerPoint opens a chart it cannot edit; copy both and leave the Target
     * alone, and the copy points at a workbook that is not in this package at
     * all.
     *
     * The destination already holds a `chart1.xml` of its own, so a copy that
     * reused the library's name would be visible here as the destination's
     * chart quietly becoming the element's.
     */
    const pkg = await crowded();
    const chart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="urn:x"><c:marker>SSF-CARRIED-CHART</c:marker></c:chartSpace>`;
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
      types: {
        "ppt/charts/chart1.xml": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      store: storeOf({
        "ppt/charts/chart1.xml": chart,
        "ppt/charts/_rels/chart1.xml.rels": relsXml([
          { id: "rId1", type: REL_TYPE.package, target: "../embeddings/Microsoft_Excel_Worksheet.xlsx" },
        ]),
        "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx": PNG,
      }),
    });

    expect(carried.parts.get("ppt/charts/chart1.xml")).toBe("ppt/charts/chart2.xml");
    expect(carried.parts.get("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx")).toBe(
      "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    );
    // The library's own name is nowhere in this deck, so a Target left alone
    // points at nothing.
    expect(pkg.has("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx")).toBe(false);
    const nested = await relsOf(pkg, "ppt/charts/chart2.xml");
    expect(nested.map((r) => r.target)).toEqual(["../embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    expect(pkg.has(pkg.resolved("ppt/charts/chart2.xml", nested[0]?.target ?? ""))).toBe(true);
    // The destination's chart is untouched, and the element's is the new part.
    expect(await pkg.text("ppt/charts/chart2.xml")).toContain("SSF-CARRIED-CHART");
    expect(await pkg.text("ppt/charts/chart1.xml")).not.toContain("SSF-CARRIED-CHART");
  });

  it("resolves a nested target against the part that owns it, however it is spelled", async () => {
    /**
     * Four spellings, resolved against `ppt/charts/chart1.xml`. The first two
     * are what a rels part really uses — `..` out of the chart's directory, and
     * a leading `/` for the package root. The other two are what a producer
     * that joins path segments emits, and the resolver has a case for each:
     * without them a `.` step becomes a directory called "." and a doubled
     * separator becomes a segment called "", so the store is asked for a part
     * nobody holds and the copy fails by NAME rather than silently — but only
     * because the missing-part refusal is there to catch it.
     */
    const pkg = await plain();
    const store = storeOf({
      "ppt/charts/chart1.xml": "<c:chartSpace xmlns:c='urn:x'/>",
      "ppt/charts/_rels/chart1.xml.rels": relsXml([
        { id: "rId1", type: REL_TYPE.package, target: "../embeddings/book.xlsx" },
        { id: "rId2", type: IMAGE, target: "/ppt/media/absolute.png" },
        { id: "rId3", type: REL_TYPE.chartUserShapes, target: "./drawing1.xml" },
        { id: "rId4", type: IMAGE, target: "..//media/doubled.png" },
      ]),
      "ppt/embeddings/book.xlsx": PNG,
      "ppt/media/absolute.png": PNG,
      "ppt/charts/drawing1.xml": "<c:userShapes xmlns:c='urn:x'/>",
      "ppt/media/doubled.png": PNG,
    });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
      types: {},
      store,
    });

    expect([...carried.parts.keys()].sort()).toEqual([
      "ppt/charts/chart1.xml",
      "ppt/charts/drawing1.xml",
      "ppt/embeddings/book.xlsx",
      "ppt/media/absolute.png",
      "ppt/media/doubled.png",
    ]);
    // And every rewritten Target names a part this package actually holds,
    // which is the only thing PowerPoint cares about.
    const copy = carried.parts.get("ppt/charts/chart1.xml") as string;
    for (const nested of await relsOf(pkg, copy)) {
      expect(pkg.has(pkg.resolved(copy, nested.target)), `${nested.target} names no part`).toBe(true);
    }
  });

  it("reads a rels part the store hands back as bytes", async () => {
    // The pane fetches parts over the network and the store's type allows
    // either; a rels part arriving as bytes must not be parsed as "[object
    // Uint8Array]", which yields a document with no relationships and a chart
    // whose workbook is quietly left behind.
    const pkg = await plain();
    const rels = relsXml([{ id: "rId1", type: REL_TYPE.package, target: "../embeddings/book.xlsx" }]);
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
      types: {},
      store: storeOf({
        "ppt/charts/chart1.xml": "<c:chartSpace xmlns:c='urn:x'/>",
        "ppt/charts/_rels/chart1.xml.rels": new TextEncoder().encode(rels),
        "ppt/embeddings/book.xlsx": PNG,
      }),
    });
    expect(carried.parts.get("ppt/embeddings/book.xlsx")).toBe("ppt/embeddings/book1.xlsx");
    expect((await relsOf(pkg, carried.parts.get("ppt/charts/chart1.xml") as string)).map((r) => r.target)).toEqual([
      "../embeddings/book1.xlsx",
    ]);
  });

  it("gives a part with no relationships of its own no rels part", async () => {
    // A picture has none, and most of what the library carries is a picture.
    // An empty `<Relationships/>` beside every copied image is legal and
    // useless, and it is what a version that always wrote one would produce.
    const pkg = await plain();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", IMAGE, "ppt/media/image1.png")],
      types: {},
      store: storeOf({ "ppt/media/image1.png": PNG }),
    });
    const name = carried.parts.get("ppt/media/image1.png") as string;
    expect(pkg.has(Pkg.relsPathFor(name))).toBe(false);
  });

  it("leaves an external target on a carried part alone, and copies nothing for it", async () => {
    /**
     * A chart with a hyperlink on it. The URL is not a part: resolving it would
     * name `ppt/charts/https:/example.invalid/help`, the store would refuse it,
     * and an element that opens a link would fail to insert at all.
     */
    const pkg = await plain();
    const store = storeOf({
      "ppt/charts/chart1.xml": "<c:chartSpace xmlns:c='urn:x'/>",
      "ppt/charts/_rels/chart1.xml.rels": relsXml([
        { id: "rId1", type: HYPERLINK, target: "https://example.invalid/help", external: true },
        { id: "rId2", type: REL_TYPE.package, target: "../embeddings/book.xlsx" },
      ]),
      "ppt/embeddings/book.xlsx": PNG,
    });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
      types: {},
      store,
    });

    expect(carried.parts.size, "the URL was copied as if it were a part").toBe(2);
    expect(store.asked.some((p) => p.includes("example.invalid"))).toBe(false);
    const nested = await relsOf(pkg, carried.parts.get("ppt/charts/chart1.xml") as string);
    expect(nested.find((r) => r.id === "rId1")).toEqual({
      id: "rId1",
      type: HYPERLINK,
      target: "https://example.invalid/help",
      mode: "External",
    });
  });

  it("skips a relationship with no target rather than copying a part called nothing", async () => {
    // A rels part missing a `Target` is malformed, and the two ways to treat it
    // are to skip it or to resolve "" — which names the owner's own DIRECTORY
    // and asks the store for it. Both relationships stay on the copy either
    // way; only the store's asks say which happened.
    const pkg = await plain();
    const store = storeOf({
      "ppt/charts/chart1.xml": "<c:chartSpace xmlns:c='urn:x'/>",
      "ppt/charts/_rels/chart1.xml.rels":
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${IMAGE}"/><Relationship Id="rId2" Type="${IMAGE}" Target=""/>` +
        `</Relationships>`,
    });
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
      types: {},
      store,
    });

    expect(carried.parts.size).toBe(1);
    expect(store.asked).toEqual(["ppt/charts/chart1.xml", "ppt/charts/_rels/chart1.xml.rels"]);
    expect((await relsOf(pkg, carried.parts.get("ppt/charts/chart1.xml") as string)).map((r) => r.id)).toEqual([
      "rId1",
      "rId2",
    ]);
  });

  it("does not loop on two parts that reach each other", async () => {
    // The reason the new name goes into the map BEFORE the recursion. Without
    // that line this test does not fail — it never returns.
    const pkg = await crowded();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId5", REL_TYPE.tags, "ppt/tags/tag1.xml")],
      types: {},
      store: storeOf({
        "ppt/tags/tag1.xml": "<p:tagLst xmlns:p='urn:x'/>",
        "ppt/tags/_rels/tag1.xml.rels": relsXml([{ id: "rId1", type: REL_TYPE.tags, target: "tag2.xml" }]),
        "ppt/tags/tag2.xml": "<p:tagLst xmlns:p='urn:x'/>",
        "ppt/tags/_rels/tag2.xml.rels": relsXml([{ id: "rId1", type: REL_TYPE.tags, target: "tag1.xml" }]),
      }),
    });

    expect(carried.parts.get("ppt/tags/tag1.xml")).toBe("ppt/tags/tag10.xml");
    expect(carried.parts.get("ppt/tags/tag2.xml")).toBe("ppt/tags/tag11.xml");
    // Each copy points at the OTHER copy, not at the library's part.
    expect((await relsOf(pkg, "ppt/tags/tag10.xml")).map((r) => r.target)).toEqual(["tag11.xml"]);
    expect((await relsOf(pkg, "ppt/tags/tag11.xml")).map((r) => r.target)).toEqual(["tag10.xml"]);
  });
});

describe("a relationship the element itself declares external", () => {
  it("arrives on the slide as External, with nothing copied for it", async () => {
    // `MarkupRel.external` carries the flag the harvest read. Without
    // `TargetMode="External"` a consumer reads the URL as a part name, finds
    // nothing, and reports the whole file as damaged.
    const pkg = await plain();
    const store = storeOf({});
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId9", HYPERLINK, "https://example.invalid/x", true)],
      types: {},
      store,
    });

    expect(carried.parts.size).toBe(0);
    expect(store.asked, "the store was asked for a URL").toEqual([]);
    const added = (await relsOf(pkg, SLIDE)).find((r) => r.id === carried.ids.get("rId9"));
    expect(added).toEqual({
      id: carried.ids.get("rId9") as string,
      type: HYPERLINK,
      target: "https://example.invalid/x",
      mode: "External",
    });
  });

  it("still carries the internal relationships beside it", async () => {
    // A link and a picture on one element. An early `continue` that skipped the
    // rest of the loop would lose the picture, and the shape would draw empty.
    const pkg = await plain();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: [rel("rId9", HYPERLINK, "https://example.invalid/x", true), rel("rId5", IMAGE, "ppt/media/image1.png")],
      types: {},
      store: storeOf({ "ppt/media/image1.png": PNG }),
    });
    expect([...carried.ids.keys()]).toEqual(["rId9", "rId5"]);
    expect([...carried.parts.keys()]).toEqual(["ppt/media/image1.png"]);
  });
});

describe("what it refuses", () => {
  it("refuses a part the store does not hold, by name", async () => {
    // The one failure that must not be survived quietly: the alternative is
    // markup pointing at a relationship pointing at nothing, which is what
    // PowerPoint calls a damaged file.
    const pkg = await plain();
    await expect(
      carry({
        pkg,
        owner: SLIDE,
        rels: [rel("rId5", IMAGE, "ppt/media/image1.png")],
        types: {},
        store: storeOf({}),
      }),
    ).rejects.toThrow('the catalogue has no part "ppt/media/image1.png"');
    // And nothing half-written behind it.
    expect(pkg.has("ppt/media/image1.png")).toBe(false);
    expect((await relsOf(pkg, SLIDE)).map((r) => r.type)).toEqual([REL_TYPE.slideLayout]);
  });

  it("names the nested part that is missing, not the one that reached it", async () => {
    /**
     * A chart whose workbook the store does not hold. Naming the CHART here
     * sends whoever reads the message to a part that is present and fine, and
     * the error is an instrument: it should name the condition it actually
     * found.
     */
    const pkg = await plain();
    await expect(
      carry({
        pkg,
        owner: SLIDE,
        rels: [rel("rId7", REL_TYPE.chart, "ppt/charts/chart1.xml")],
        types: {},
        store: storeOf({
          "ppt/charts/chart1.xml": "<c:chartSpace xmlns:c='urn:x'/>",
          "ppt/charts/_rels/chart1.xml.rels": relsXml([
            { id: "rId1", type: REL_TYPE.package, target: "../embeddings/book.xlsx" },
          ]),
        }),
      }),
    ).rejects.toThrow('the catalogue has no part "ppt/embeddings/book.xlsx"');
  });
});

describe("the chart the library really ships", () => {
  let library: Harvest;

  beforeAll(async () => {
    const pkg = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
    library = await harvest(pkg, { size: "16:9", names: NAMES });
  }, 120_000);

  it("carries the chart and the workbook behind it, under names the destination is not using", async () => {
    /**
     * The synthetic cases above are a story about XML this test does not have
     * to trust: these are the owner's own parts, the harvest's own `rels` and
     * the catalogue's own content types, and the two names happen to be the
     * two hard ones — a `chart1.xml` the destination also has, and a workbook
     * with no digits to extend.
     */
    const el = library.catalogue.elements.find((e) => e.markup.rels.some((r) => r.type === REL_TYPE.chart));
    expect(el, "the library no longer has a chart to carry").toBeDefined();

    const pkg = await crowded();
    const carried = await carry({
      pkg,
      owner: SLIDE,
      rels: (el as CatalogueElement).markup.rels,
      types: library.catalogue.carried,
      store: (path) => Promise.resolve(library.parts.get(path)),
    });

    expect(carried.parts.get("ppt/charts/chart1.xml")).toBe("ppt/charts/chart2.xml");
    expect(carried.parts.get("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx")).toBe(
      "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    );
    expect((await relsOf(pkg, "ppt/charts/chart2.xml")).map((r) => r.target)).toEqual([
      "../embeddings/Microsoft_Excel_Worksheet1.xlsx",
    ]);
    // Every part that arrived is in the package and is classified. A part with
    // no content type is a package PowerPoint refuses outright, without saying
    // which part it could not classify.
    for (const [, name] of carried.parts) {
      expect(pkg.has(name), `${name} was not written`).toBe(true);
      expect(await pkg.contentTypeOf(name), `${name} has no content type`).toBeDefined();
    }
    expect(pkg.has("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx")).toBe(false);
  });
});
