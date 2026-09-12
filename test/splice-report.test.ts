import { readFileSync } from "node:fs";
import { beforeAll, describe, expect, it } from "vitest";
import { Pkg, harvest } from "../src/core/index.js";
import type { Catalogue, Element as CatalogueElement, Names } from "../src/core/index.js";
import { contentCount, occupiedBoxes, topLevelShapes } from "../src/core/catalogue/boxes.js";
import { P_NS, child, elements } from "../src/core/pptx/xml.js";
import { onlySlide, splice, type SpliceElement, type SpliceReport } from "../src/core/splice/splice.js";
import { fractionOf } from "../src/pane/card.js";
import JSZip from "jszip";
import { makeDeck } from "./fixtures/deck.js";

/**
 * What the splice REPORTS, held against what it actually did.
 *
 * `SpliceReport` carries `shapes`, `grouped`, `placeholders`, `held`, `pinned`
 * and `landed`, and those numbers are not diagnostics — the pane shows them to
 * the user and decides on them. `held` is the whole of the "Move to a new
 * slide" offer (`docs/DESIGN.md` section 6). `landed` becomes the grey
 * rectangle the preview card draws. `placeholders` is how many of the user's
 * own boxes went.
 *
 * Nothing was checking any of them. `splice.test.ts` sweeps all 117 elements
 * through the package integrity checker, which asks whether PowerPoint can OPEN
 * the result — a real and expensive question, and a different one from whether
 * the report describes it. A report that is wrong does not produce a damaged
 * file; it produces a pane that says something untrue, which is the failure
 * nobody notices.
 *
 * And the round trip at the end is the hand check that settled this project's
 * worst bug. `CLAUDE.md` records it: an undo reported success and changed
 * nothing, and what caught it was putting a triangle on a slide and comparing
 * the shape inventory before and after, id by id and name by name. That was
 * done once, by a person, against a real PowerPoint. Here it is over every
 * element, on every run.
 */

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

let library: { catalogue: Catalogue; parts: Map<string, Uint8Array | string> };

beforeAll(async () => {
  const pkg = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
  library = await harvest(pkg, { size: "16:9", names: NAMES });
}, 120_000);

/** Three ordinary slides, the middle one the user is on. */
function destination(): Promise<Uint8Array> {
  return makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }]);
}

function asSplice(el: CatalogueElement): SpliceElement {
  return {
    id: el.id,
    name: el.name,
    kind: el.kind,
    box: el.box,
    landing: el.landing,
    markup: { xml: el.markup.xml, rels: el.markup.rels },
  };
}

const store = (path: string): Promise<Uint8Array | string | undefined> => Promise.resolve(library.parts.get(path));

async function put(
  el: CatalogueElement,
  deck: Uint8Array | string,
  slide: number,
  options: { target: "onto" | "new"; group: boolean; colours: "deck" | "library" },
): Promise<SpliceReport> {
  return splice({
    deck,
    slide,
    element: asSplice(el),
    options,
    catalogue: { version: "v1", carried: library.catalogue.carried, theme: library.catalogue.theme },
    store,
  });
}

/**
 * Every top-level shape of a slide, as `id|name`.
 *
 * The two together, because either alone is the check that misses: PowerPoint
 * renumbers ids freely and two shapes can share a name. This is the same pair
 * the round that settled the undo defect compared by hand.
 */
async function inventory(pkg: Pkg, path: string): Promise<string[]> {
  const doc = await pkg.doc(path);
  return topLevelShapes(doc).map((shape) => {
    for (const container of Array.from(shape.childNodes)) {
      if (container.nodeType !== 1) continue;
      const cNvPr = child(container as Element, P_NS, "cNvPr");
      if (cNvPr) return `${cNvPr.getAttribute("id") ?? "?"}|${cNvPr.getAttribute("name") ?? "?"}`;
    }
    return "?|?";
  });
}

/**
 * The fixture deck, with slide 2 carrying a real, EMPTY body placeholder.
 *
 * The fixture's own "Body" shape has `<p:nvPr></p:nvPr>` and no `<p:ph>` in it,
 * so it is a text box rather than a placeholder and `emptyBodyPlaceholders`
 * rightly never touches it. Which means no fixture in this suite can produce a
 * ghost, and the count that reports them had nothing to be measured against.
 */
async function withEmptyPlaceholder(): Promise<Uint8Array> {
  const source = await JSZip.loadAsync(await destination());
  const xml = await (source.file("ppt/slides/slide2.xml") as JSZip.JSZipObject).async("string");
  const ghost =
    `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Click to add text"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
  const out = new JSZip();
  for (const [name, file] of Object.entries(source.files)) {
    if (file.dir) continue;
    if (name === "ppt/slides/slide2.xml") out.file(name, xml.replace("</p:spTree>", `${ghost}</p:spTree>`));
    else out.file(name, await file.async("uint8array"));
  }
  return out.generateAsync({ type: "uint8array" });
}

/** The slide `<p:sldIdLst>` lists at `index`. */
async function slideAt(pkg: Pkg, index: number): Promise<string> {
  const path = (await pkg.slidePaths())[index];
  if (path === undefined) throw new Error(`no slide ${index}`);
  return path;
}

/** A handful of elements that between them cover every landing and both kinds. */
function sample(): CatalogueElement[] {
  const all = library.catalogue.elements;
  const byLanding = new Map<string, CatalogueElement>();
  for (const el of all) {
    const key = `${el.landing}:${el.kind}`;
    if (!byLanding.has(key)) byLanding.set(key, el);
  }
  return [...byLanding.values()];
}

describe("what the report says about the shapes", () => {
  it("covers every landing the library has, so the sweep is not one element four times", () => {
    const kinds = new Set(sample().map((el) => `${el.landing}:${el.kind}`));
    expect(kinds.size, "the library grew a landing this file does not exercise").toBeGreaterThan(2);
  });

  it("adds what it says it added, grouped or loose", async () => {
    for (const el of sample()) {
      for (const group of [true, false]) {
        const deck = await destination();
        const before = await Pkg.open(deck).then((p) => inventory(p, "ppt/slides/slide2.xml"));
        const report = await put(el, deck, 1, { target: "onto", group, colours: "deck" });
        const pkg = await Pkg.open(report.base64);
        const after = await inventory(pkg, report.slidePath);

        const survived = after.filter((s) => before.includes(s));
        const arrived = after.length - survived.length;
        expect(arrived, `${el.id} grouped=${group}: report.shapes`).toBe(report.shapes);
        // `grouped` is a claim about the SHAPE of what arrived, and one top-level
        // shape is the only thing it can mean.
        if (report.grouped) expect(report.shapes, `${el.id}: grouped means one`).toBe(1);
      }
    }
  }, 120_000);

  it("takes away exactly as many of the user's own boxes as it says", async () => {
    /**
     * On a slide whose placeholder is EMPTY, which is the only case the count is
     * for — and the first version of this asserted it against the ordinary
     * fixture, where nothing is a ghost and the answer is 0 either way. Setting
     * `placeholders: 0` in the source left it green. A case that compares zero
     * to zero is the exact fault this whole file is hunting, so it is worth
     * saying that it was written here first.
     */
    let sawOne = false;
    for (const el of sample()) {
      const deck = await withEmptyPlaceholder();
      const before = await Pkg.open(deck).then((p) => inventory(p, "ppt/slides/slide2.xml"));
      const report = await put(el, deck, 1, { target: "onto", group: true, colours: "deck" });
      const after = await Pkg.open(report.base64).then((p) => inventory(p, report.slidePath));
      const lost = before.filter((s) => !after.includes(s));
      expect(lost.length, `${el.id}: report.placeholders against what actually went`).toBe(report.placeholders);
      if (report.placeholders > 0) sawOne = true;
    }
    expect(sawOne, "no element removed a ghost, so this compared 0 to 0 again").toBe(true);
  }, 120_000);

  it("loses nothing of the user's that it did not report losing", async () => {
    /**
     * The additive property, and the direction that matters.
     *
     * An insert that quietly drops one of the user's own shapes is the failure
     * this whole family of code is written to avoid, and it would not produce a
     * damaged package — the integrity checker would pass it. The only way to
     * see it is to count the shapes before and after.
     */
    const el = sample()[0] as CatalogueElement;
    const deck = await destination();
    const before = await Pkg.open(deck).then((p) => inventory(p, "ppt/slides/slide2.xml"));
    const report = await put(el, deck, 1, { target: "onto", group: true, colours: "deck" });
    const after = await Pkg.open(report.base64).then((p) => inventory(p, report.slidePath));
    expect(report.placeholders, "the fixture's placeholder has text, so nothing is a ghost").toBe(0);
    for (const shape of before) expect(after, `${shape} was on the slide and is not now`).toContain(shape);
  });
});

describe("what the report says about the slide it landed on", () => {
  it("counts what the slide the request NAMED was holding", async () => {
    // Not the rebuilt copy, and not the blanked clone "as a new slide" starts
    // from — the distinction `SpliceReport.held` makes in its own docstring,
    // because the offer it drives is about the slide the user was looking at.
    const el = sample()[0] as CatalogueElement;
    const deck = await destination();
    const source = await Pkg.open(deck);
    const size = { width: 12192000, height: 6858000 };
    const expected = contentCount(await source.doc(await slideAt(source, 1)), size.width, size.height);

    for (const target of ["onto", "new"] as const) {
      const report = await put(el, deck, 1, { target, group: true, colours: "deck" });
      expect(report.held, `target=${target}: held is about the slide the user was on`).toBe(expected);
    }
  });

  it("puts the element where it says it put it", async () => {
    /**
     * `landed` becomes the preview card's grey rectangle, through `fractionOf`.
     * The wrong-slide-size version of this exact sum has already shipped once —
     * the card divided by the LIBRARY's slide size instead of the user's, so on
     * any deck that borrowed the nearest library it drew the right rectangle in
     * the wrong place. That was found by hunting; this is the arithmetic held
     * against the package.
     */
    const deck = await destination();
    const size = { width: 12192000, height: 6858000 };
    for (const el of sample()) {
      const report = await put(el, deck, 1, { target: "onto", group: true, colours: "deck" });
      const claimed = fractionOf(report.landed, size);
      expect(claimed, `${el.id}: fractionOf could not read the landing`).toBeDefined();

      const pkg = await Pkg.open(report.base64);
      const boxes = occupiedBoxes(await pkg.doc(report.slidePath), size.width, size.height);
      // The element is somewhere in the union of what is on the slide, and its
      // own rectangle has to be one the reader can find within a pixel of the
      // claim. A tenth of one per cent of a slide is under half a pixel at any
      // size the pane draws the card at.
      const near = boxes.some(
        (b) =>
          Math.abs(b.x - (claimed?.x ?? -9)) < 0.001 &&
          Math.abs(b.y - (claimed?.y ?? -9)) < 0.001 &&
          Math.abs(b.w - (claimed?.w ?? -9)) < 0.001 &&
          Math.abs(b.h - (claimed?.h ?? -9)) < 0.001,
      );
      expect(near, `${el.id}: the card would draw a box that is not where the element is`).toBe(true);
    }
  }, 120_000);
});

describe("the round trip that settled the undo defect", () => {
  /**
   * Insert an element, take the slide back, and the slide is what it was.
   *
   * `CLAUDE.md` records how this was caught the first time: an undo that
   * reported success and changed nothing, found by putting a TRIANGLE on a
   * slide holding a title and a white box and comparing the shape inventory
   * before and after, id by id and name by name. A person did that once,
   * against a real PowerPoint, on 2026-09-11.
   *
   * `onlySlide` is the engine's half of that undo — the package holding the
   * user's original slide, which the pane hands back and then removes the
   * rebuilt one. So the property is: what `onlySlide` produces has the same
   * shapes as the slide the insert was aimed at, whatever the insert did.
   */
  it("gives back a slide with the same shapes it had, for every landing", async () => {
    for (const el of sample()) {
      const deck = await destination();
      const source = await Pkg.open(deck);
      const path = await slideAt(source, 1);
      const before = await inventory(source, path);
      expect(before.length, "a fixture slide with nothing on it would prove nothing").toBeGreaterThan(0);

      // The insert happens and is then taken back, which is the sequence the
      // pane performs: the original bytes are what the undo is built from.
      await put(el, deck, 1, { target: "onto", group: true, colours: "deck" });
      const restored = await onlySlide(deck, 1);
      const back = await Pkg.open(restored.base64);
      expect(await inventory(back, restored.path), `${el.id}: the slide came back different`).toEqual(before);
    }
  }, 120_000);

  it("gives back exactly one slide, and it is the one asked for", async () => {
    const deck = await destination();
    const restored = await onlySlide(deck, 1);
    const pkg = await Pkg.open(restored.base64);
    expect(await pkg.slidePaths()).toEqual([restored.path]);
    // The middle slide, by the text only it carries.
    expect(await pkg.text(restored.path)).toContain("Second");
  });

  it("keeps the slide list and the part it names in step", async () => {
    // A package whose `<p:sldIdLst>` names a part that is not there is one
    // PowerPoint opens as repaired, having dropped whatever it decided to drop.
    const deck = await destination();
    const restored = await onlySlide(deck, 2);
    const pkg = await Pkg.open(restored.base64);
    const pres = await pkg.doc("ppt/presentation.xml");
    const list = elements(pres, P_NS, "sldIdLst")[0];
    expect(elements(list as Element, P_NS, "sldId")).toHaveLength(1);
    expect(pkg.has(restored.path)).toBe(true);
  });
});
