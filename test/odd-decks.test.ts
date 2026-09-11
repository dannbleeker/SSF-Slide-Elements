/**
 * Destinations the splice had never been handed.
 *
 * Every other deck in this suite is three ordinary slides with a paragraph on
 * each. The decks people actually work in are not that: one-slide decks, empty
 * slides, slides carrying a modern chart or SmartArt, slides still stamped by a
 * catalogue two versions old, the same stamp dropped twice in a row. None of
 * those had ever reached the engine.
 *
 * They all passed the first time they were run. That is the point of writing
 * them down: the claim "it handles a deck like this" is now something the suite
 * says rather than something nobody had asked.
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs.
import { packageProblems } from "../scripts/package-integrity.mjs";
import { Pkg, harvest } from "../src/core/index.js";
import type { Names } from "../src/core/index.js";
import { TAG_CATALOGUE, TAG_ELEMENT, readShapeTags, writeShapeTags } from "../src/core/pptx/tags.js";
import { removeElement } from "../src/core/splice/remove.js";
import { onlySlide, splice } from "../src/core/splice/splice.js";
import { makeDeck } from "./fixtures/deck.js";

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;
let library: Awaited<ReturnType<typeof harvest>>;

async function partsOf(bytes: Uint8Array): Promise<Map<string, string | Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const parts = new Map<string, string | Uint8Array>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const text = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, text ? await file.async("string") : await file.async("uint8array"));
  }
  return parts;
}

async function put(deck: string | Uint8Array, id: string, slide: number, target: "onto" | "new" = "onto") {
  const el = library.catalogue.elements.find((e) => e.id === id);
  if (!el) throw new Error(`no ${id}`);
  return splice({
    deck,
    slide,
    element: { id: el.id, name: el.name, kind: el.kind, box: el.box, landing: el.landing, markup: el.markup },
    options: { target, group: true, colours: "deck" },
    catalogue: { version: "v1", carried: library.catalogue.carried, theme: library.catalogue.theme },
    store: (path) => Promise.resolve(library.parts.get(path)),
  });
}

describe("destinations nobody has handed it", () => {
  const WHOLE = "hvid-kasse-2x1-vertikale";
  const PART = "markeringer-1";

  it("harvests the library once", async () => {
    library = await harvest(await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
      size: "16:9",
      names: NAMES,
    });
    expect(library.catalogue.elements.length).toBe(117);
  }, 120000);

  it("a deck of exactly one slide", async () => {
    const deck = await makeDeck([{ paragraphs: [["Only"]] }]);
    const report = await put(deck, WHOLE, 0);
    expect(report.deckSlides).toBe(1);
    expect(packageProblems(await partsOf(await (await Pkg.open(report.base64)).toBytes()))).toEqual([]);
    // And the undo of that insert: putting the user's one slide back.
    const { base64, path } = await onlySlide(deck, 0);
    expect(path).toBe("ppt/slides/slide1.xml");
    expect((await (await Pkg.open(base64)).slidePaths()).length).toBe(1);
  });

  it("a slide with nothing on it at all", async () => {
    const deck = await makeDeck([{ paragraphs: [] }, { paragraphs: [["Second"]] }]);
    const report = await put(deck, WHOLE, 0);
    expect(packageProblems(await partsOf(await (await Pkg.open(report.base64)).toBytes()))).toEqual([]);
  });

  it("a slide carrying OUR tag from an older catalogue", async () => {
    // The deck a user built last month, opened by this month's add-in. The tag
    // is ours, the catalogue version is not this one, and the element id may
    // not even be in the library any more.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }]));
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const shape = doc.getElementsByTagName("p:sp")[0] as unknown as Element;
    await writeShapeTags(pkg, "ppt/slides/slide1.xml", shape, [
      [TAG_ELEMENT, "fortroligt"],
      [TAG_CATALOGUE, "0000deadbeef"],
    ]);
    const deck = await pkg.toBase64();
    const report = await put(deck, WHOLE, 0);
    const out = await Pkg.open(report.base64);
    expect(packageProblems(await partsOf(await out.toBytes()))).toEqual([]);
    // The old stamp survives the rebuild, which is what "Used in this deck"
    // reads back.
    const tags = await readShapeTags(out, report.slidePath);
    expect(tags.some((t) => t.element === "fortroligt")).toBe(true);
    expect(tags.some((t) => t.element === WHOLE)).toBe(true);
    // And it can be removed by name even though this catalogue has no such id.
    const back = await removeElement({ deck: report.base64, slide: 0, element: "fortroligt" });
    expect(
      (await readShapeTags(await Pkg.open(back.base64), back.slidePath)).some((t) => t.element === "fortroligt"),
    ).toBe(false);
  });

  it("a slide holding a modern chart in mc:AlternateContent", async () => {
    const deck = await makeDeck([
      { paragraphs: [["First"]] },
      { paragraphs: [["Second"]], modernChart: { title: "Funnel", categories: ["a", "b"], values: ["1", "2"] } },
    ]);
    const report = await put(deck, WHOLE, 1);
    const out = await Pkg.open(report.base64);
    expect(packageProblems(await partsOf(await out.toBytes()))).toEqual([]);
    // The user's own chart is still on the rebuilt slide.
    expect(await out.text(report.slidePath)).toContain("AlternateContent");
  });

  it("a slide holding SmartArt", async () => {
    const deck = await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]], smartArt: ["one", "two"] }]);
    const report = await put(deck, PART, 1);
    const out = await Pkg.open(report.base64);
    expect(packageProblems(await partsOf(await out.toBytes()))).toEqual([]);
  });

  it("inserting as a new slide after the LAST slide", async () => {
    const deck = await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }]);
    const report = await put(deck, WHOLE, 1, "new");
    expect(packageProblems(await partsOf(await (await Pkg.open(report.base64)).toBytes()))).toEqual([]);
  });

  it("the same element inserted twice onto one slide, then removed once", async () => {
    const first = await put(await makeDeck([{ paragraphs: [["First"]] }]), PART, 0);
    const second = await put(first.base64, PART, 0);
    const out = await Pkg.open(second.base64);
    const before = (await readShapeTags(out, second.slidePath)).filter((t) => t.element === PART);
    expect(before.length).toBe(2);
    const back = await removeElement({ deck: second.base64, slide: 0, element: PART });
    expect(back.removed).toBe(2);
    expect(
      (await readShapeTags(await Pkg.open(back.base64), back.slidePath)).filter((t) => t.element === PART),
    ).toEqual([]);
  });
});
