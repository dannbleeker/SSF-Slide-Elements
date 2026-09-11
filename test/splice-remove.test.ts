import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { packageProblems } from "../scripts/package-integrity.mjs";
import { Pkg, harvest } from "../src/core/index.js";
import type { Catalogue, Element as CatalogueElement, Names } from "../src/core/index.js";
import { readShapeTags } from "../src/core/pptx/tags.js";
import { removeElement, slidesHolding } from "../src/core/splice/remove.js";
import { splice, type SpliceElement } from "../src/core/splice/splice.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * Taking an element back off a slide.
 *
 * The only thing this project does that REMOVES something from the user's own
 * deck, so the cases below are about what it leaves alone as much as what it
 * takes: a shape the user drew, a shape another add-in tagged, and the slide's
 * own layout all have to survive it.
 *
 * Every case builds its deck by actually INSERTING first, through the same
 * splice the product uses, rather than by hand-writing a tagged shape. A
 * fixture written to match the reader would agree with a remover that read the
 * wrong thing; a deck the insert produced is the deck the feature will meet.
 */

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

let library: { catalogue: Catalogue; parts: Map<string, Uint8Array | string> };

beforeAll(async () => {
  const pkg = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
  library = await harvest(pkg, { size: "16:9", names: NAMES });
}, 120_000);

function element(id: string): CatalogueElement {
  const found = library.catalogue.elements.find((e) => e.id === id);
  if (!found) throw new Error(`the library has no element "${id}"`);
  return found;
}

function asSplice(el: CatalogueElement): SpliceElement {
  return { id: el.id, name: el.name, kind: el.kind, box: el.box, landing: el.landing, markup: el.markup };
}

const store = (path: string): Promise<Uint8Array | string | undefined> => Promise.resolve(library.parts.get(path));

/** The user's deck after an insert of `id` onto slide `slide`, as bytes. */
async function afterInserting(id: string, slide = 1, deck?: Uint8Array): Promise<string> {
  const base =
    deck ?? (await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }]));
  const report = await splice({
    deck: base,
    slide,
    element: asSplice(element(id)),
    options: { target: "onto", group: true, colours: "deck" },
    catalogue: { version: "test", carried: library.catalogue.carried, theme: library.catalogue.theme },
    store,
  });
  // The package the host would be handed lists one slide; what the USER then
  // has is that deck with the rebuild in it. For these cases the package itself
  // is the deck under test, which is what the pane hands back for the next
  // read anyway.
  return report.base64;
}

async function partsOf(bytes: Uint8Array): Promise<Map<string, string | Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const parts = new Map<string, string | Uint8Array>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const xml = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, xml ? await file.async("string") : await file.async("uint8array"));
  }
  return parts;
}

describe("finding where an element is", () => {
  it("names the slides that carry it, and no others", async () => {
    const deck = await afterInserting("hvid-kasse-2x1-vertikale");
    const pkg = await Pkg.open(deck);
    // The package the splice hands back lists one slide, which is the rebuilt
    // one — so that is where the element is.
    expect(await slidesHolding(pkg, "hvid-kasse-2x1-vertikale")).toEqual([0]);
    expect(await slidesHolding(pkg, "something-else-entirely")).toEqual([]);
  });
});

describe("taking it off", () => {
  const ID = "hvid-kasse-2x1-vertikale";

  it("removes every shape the insert tagged, and says how many", async () => {
    const deck = await afterInserting(ID);
    const before = await Pkg.open(deck);
    const beforePath = (await before.slidePaths())[0] as string;
    const tagged = await readShapeTags(before, beforePath);
    expect(tagged.filter((t) => t.element === ID).length).toBeGreaterThan(0);

    const report = await removeElement({ deck, slide: 0, element: ID });
    expect(report.removed).toBe(tagged.filter((t) => t.element === ID).length);
    expect(report.left).toBe(0);

    const after = await Pkg.open(report.base64);
    expect((await readShapeTags(after, report.slidePath)).filter((t) => t.element === ID)).toEqual([]);
  });

  it("leaves the rest of the slide exactly where it was", async () => {
    // The user's own shape is the thing this feature must never touch. The
    // fixture's slide carries one, with text in it.
    const deck = await afterInserting(ID);
    const report = await removeElement({ deck, slide: 0, element: ID });
    const after = await Pkg.open(report.base64);
    const xml = await after.text(report.slidePath);
    expect(xml).toContain("Second");
  });

  it("hands back a package the integrity checker finds nothing wrong with", async () => {
    const deck = await afterInserting(ID);
    const report = await removeElement({ deck, slide: 0, element: ID });
    const bytes = await (await Pkg.open(report.base64)).toBytes();
    expect(packageProblems(await partsOf(bytes))).toEqual([]);
  });

  it("lists exactly the rebuilt slide, like every other package this add-in hands over", async () => {
    const deck = await afterInserting(ID);
    const report = await removeElement({ deck, slide: 0, element: ID });
    const after = await Pkg.open(report.base64);
    expect(await after.slidePaths()).toEqual([report.slidePath]);
  });

  it("refuses a slide that carries nothing of ours, rather than rebuilding it for nothing", async () => {
    // A rebuild that removed nothing would still be handed to the host, which
    // is an insert and a delete against somebody's deck in exchange for no
    // change at all.
    const plain = await makeDeck([{ paragraphs: [["First"]] }]);
    await expect(removeElement({ deck: plain, slide: 0, element: ID })).rejects.toThrow(/carries no/);
  });

  it("refuses a slide the deck does not have", async () => {
    const deck = await afterInserting(ID);
    await expect(removeElement({ deck, slide: 9, element: ID })).rejects.toThrow(/is not in this deck/);
  });

  it("leaves another add-in's tagged shape alone", async () => {
    // A deck touched by think-cell carries a tagged shape on every slide it has
    // seen, and the measured deck had those tags sitting in `ppt/tags/` beside
    // ours. Keying on the tag NAME is what keeps them out of this.
    const base = await makeDeck([{ paragraphs: [["First"]], shapeTags: true }, { paragraphs: [["Second"]] }]);
    const deck = await afterInserting(ID, 0, base);
    const report = await removeElement({ deck, slide: 0, element: ID });
    const after = await Pkg.open(report.base64);
    // The vendor's own tag part is untouched, and the shape that points at it
    // is still on the slide: `readShapeTags` answers nothing for it, which is
    // exactly why the removal could not have reached it.
    expect(after.has("ppt/tags/tag9.xml")).toBe(true);
    expect(await after.text("ppt/tags/tag9.xml")).toContain("VENDOR");
    const xml = await after.text(report.slidePath);
    expect(xml).toContain("custDataLst");
  });

  it("removes one element without touching another on the same slide", async () => {
    const first = await afterInserting(ID);
    const both = await splice({
      deck: first,
      slide: 0,
      element: asSplice(element("markeringer-1")),
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: { version: "test", carried: library.catalogue.carried, theme: library.catalogue.theme },
      store,
    });
    const report = await removeElement({ deck: both.base64, slide: 0, element: ID });
    const after = await Pkg.open(report.base64);
    const left = await readShapeTags(after, report.slidePath);
    expect(left.filter((t) => t.element === ID)).toEqual([]);
    expect(left.filter((t) => t.element === "markeringer-1").length).toBeGreaterThan(0);
  });
});
