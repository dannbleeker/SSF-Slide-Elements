/**
 * A deck whose slide ORDER is not its part numbering.
 *
 * Every fixture in the suite has `<p:sldId>` entries in the same order as
 * `slide1.xml`, `slide2.xml`, `slide3.xml`. A deck somebody has actually worked
 * in almost never does — dragging one slide in the thumbnail rail reorders the
 * list and leaves the parts where they were. Everything in this project that
 * takes an INDEX means an index into the listed order.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Pkg, harvest } from "../src/core/index.js";
import type { Names } from "../src/core/index.js";
import { readShapeTags } from "../src/core/pptx/tags.js";
import { P_NS, elements, serializeXml } from "../src/core/pptx/xml.js";
import { removeElement, slidesHolding } from "../src/core/splice/remove.js";
import { onlySlide, splice } from "../src/core/splice/splice.js";
import { makeDeck } from "./fixtures/deck.js";

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

/** The fixture's three slides, listed in the order 3, 1, 2. */
async function reordered(): Promise<string> {
  const pkg = await Pkg.open(
    await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }]),
  );
  const doc = await pkg.doc("ppt/presentation.xml");
  const list = elements(doc, P_NS, "sldIdLst")[0];
  if (!list) throw new Error("no sldIdLst");
  const ids = elements(list, P_NS, "sldId");
  // Move the last one to the front.
  list.insertBefore(ids[2] as never, ids[0] as never);
  pkg.setText("ppt/presentation.xml", serializeXml(doc));
  return pkg.toBase64();
}

describe("a deck whose slides have been dragged around", () => {
  it("lists them in the order the user sees, not the order they are numbered", async () => {
    const pkg = await Pkg.open(await reordered());
    const paths = await pkg.slidePaths();
    expect(paths).toEqual(["ppt/slides/slide3.xml", "ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
  });

  it("splices onto the slide at that INDEX, which is the one the user is on", async () => {
    const library = await harvest(await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
      size: "16:9",
      names: NAMES,
    });
    const el = library.catalogue.elements.find((e) => e.id === "markeringer-1");
    if (!el) throw new Error("no markeringer-1");
    const deck = await reordered();
    const report = await splice({
      deck,
      slide: 0,
      element: { id: el.id, name: el.name, kind: el.kind, box: el.box, landing: el.landing, markup: el.markup },
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: { version: "v1", carried: library.catalogue.carried, theme: library.catalogue.theme },
      store: (path) => Promise.resolve(library.parts.get(path)),
    });
    const out = await Pkg.open(report.base64);
    const xml = await out.text(report.slidePath);
    // Index 0 of the LISTED order is the part slide3.xml, whose text is "Third".
    expect(xml).toContain("Third");
    expect(xml).not.toContain("First");
    expect(await out.slidePaths()).toEqual([report.slidePath]);
  });

  it("puts the right slide back for an undo", async () => {
    const deck = await reordered();
    const { base64, path } = await onlySlide(deck, 0);
    const out = await Pkg.open(base64);
    expect(path).toBe("ppt/slides/slide3.xml");
    expect(await out.text(path)).toContain("Third");
  });

  it("removes from the slide the user means", async () => {
    const library = await harvest(await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
      size: "16:9",
      names: NAMES,
    });
    const el = library.catalogue.elements.find((e) => e.id === "markeringer-1");
    if (!el) throw new Error("no markeringer-1");
    const report = await splice({
      deck: await reordered(),
      slide: 0,
      element: { id: el.id, name: el.name, kind: el.kind, box: el.box, landing: el.landing, markup: el.markup },
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: { version: "v1", carried: library.catalogue.carried, theme: library.catalogue.theme },
      store: (path) => Promise.resolve(library.parts.get(path)),
    });
    const held = await slidesHolding(await Pkg.open(report.base64), el.id);
    expect(held).toEqual([0]);
    const back = await removeElement({ deck: report.base64, slide: 0, element: el.id });
    const after = await Pkg.open(back.base64);
    expect((await readShapeTags(after, back.slidePath)).filter((t) => t.element === el.id)).toEqual([]);
    expect(await after.text(back.slidePath)).toContain("Third");
  });
});
