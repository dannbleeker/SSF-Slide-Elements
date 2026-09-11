import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { XMLSerializer } from "@xmldom/xmldom";
import { beforeAll, describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { packageProblems } from "../scripts/package-integrity.mjs";
import { Pkg, harvest } from "../src/core/index.js";
import type { Catalogue, Element as CatalogueElement, Names } from "../src/core/index.js";
import { readShapeTags, TAG_CATALOGUE, TAG_ELEMENT } from "../src/core/pptx/tags.js";
import { A_NS, P_NS, R_NS, child, elements, parseXml, serializeXml } from "../src/core/pptx/xml.js";
import { onlySlide, splice, type SpliceElement } from "../src/core/splice/splice.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The splice, against the library the product actually ships.
 *
 * `docs/BACKLOG.md` asked for exactly this — "a sweep over every element in
 * the library into fixture decks, every package validated by the integrity
 * checker, on every commit" — and the entry was removed when the sweep shipped,
 * which is what that file does with work that lands. The synthetic cases below
 * prove each RULE in isolation; the
 * sweep proves the rules add up over 117 real elements the owner drew, which
 * between them carry tables, pictures, embedded objects, a chart with a
 * workbook behind it, groups, and tag parts PowerPoint wrote.
 *
 * The oracle is `scripts/package-integrity.mjs`, the same one CI runs, because
 * the question PowerPoint asks is binary and expensive: a package whose parts,
 * relationships and markup do not agree opens as "repaired", having silently
 * dropped whatever it decided to drop, in somebody's presentation.
 */

const NAMES = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;

/** A destination deck: three ordinary slides, the middle one the user is on. */
function destination(): Promise<Uint8Array> {
  return makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]] }, { paragraphs: [["Third"]] }]);
}

type Parts = Map<string, string | Uint8Array>;

async function partsOf(bytes: Uint8Array): Promise<Parts> {
  const zip = await JSZip.loadAsync(bytes);
  const parts: Parts = new Map();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const xml = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, xml ? await file.async("string") : await file.async("uint8array"));
  }
  return parts;
}

const problems = (parts: Parts): string[] => packageProblems(parts) as string[];

let library: { catalogue: Catalogue; parts: Map<string, Uint8Array | string> };

beforeAll(async () => {
  const pkg = await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx")));
  library = await harvest(pkg, { size: "16:9", names: NAMES });
}, 120_000);

/** An element from the real library, by the id the catalogue gives it. */
function element(id: string): CatalogueElement {
  const found = library.catalogue.elements.find((e) => e.id === id);
  if (!found) throw new Error(`the library has no element "${id}"`);
  return found;
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

/** The store the pane will be: a carried part's bytes, by the path the library knew. */
const store = (path: string): Promise<Uint8Array | string | undefined> => Promise.resolve(library.parts.get(path));

const CATALOGUE = { version: "test-version", carried: {} as Record<string, string> };

function catalogueFor(): { version: string; carried: Record<string, string>; theme: Record<string, string> } {
  return { version: CATALOGUE.version, carried: library.catalogue.carried, theme: library.catalogue.theme };
}

async function spliceOne(
  el: CatalogueElement,
  options: { target?: "onto" | "new"; group?: boolean; slide?: number; colours?: "deck" | "library" } = {},
): ReturnType<typeof splice> {
  return splice({
    deck: await destination(),
    slide: options.slide ?? 1,
    element: asSplice(el),
    options: { target: options.target ?? "onto", group: options.group ?? true, colours: options.colours ?? "deck" },
    catalogue: catalogueFor(),
    store,
  });
}

/** The rebuilt slide's shape tree, out of the package the splice produced. */
async function rebuiltTree(base64: string, slidePath: string): Promise<Element> {
  const out = await Pkg.open(base64);
  const doc = await out.doc(slidePath);
  const cSld = child(doc.documentElement, P_NS, "cSld");
  const spTree = cSld ? child(cSld, P_NS, "spTree") : undefined;
  if (!spTree) throw new Error("the rebuilt slide has no shape tree");
  return spTree;
}

function topLevelOf(spTree: Element): Element[] {
  const out: Element[] = [];
  for (const node of Array.from(spTree.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.namespaceURI === P_NS && (el.localName === "nvGrpSpPr" || el.localName === "grpSpPr")) continue;
    out.push(el);
  }
  return out;
}

function everyShapeId(spTree: Element): string[] {
  const out: string[] = [];
  const walk = (node: Element): void => {
    if (node.namespaceURI === P_NS && node.localName === "cNvPr") out.push(node.getAttribute("id") ?? "");
    for (const kid of Array.from(node.childNodes)) if (kid.nodeType === 1) walk(kid as Element);
  };
  walk(spTree);
  return out;
}

describe("the library the sweep runs over", () => {
  it("is the deck the design says it is, so the sweep is not vacuous", () => {
    expect(library.catalogue.elements.length).toBe(117);
    expect(library.catalogue.elements.filter((e) => e.kind === "part").length).toBe(21);
    // The four part families the elements really carry. A sweep over a library
    // with no pictures and no embedded objects would prove far less than this
    // one does.
    const types = new Set(library.catalogue.elements.flatMap((e) => e.markup.rels.map((r) => r.type)));
    expect([...types].map((t) => t.split("/").pop()).sort()).toEqual(["chart", "image", "oleObject", "tags"]);
  });
});

describe("one element into a deck", () => {
  it("hands back a package that lists exactly the rebuilt slide", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    const report = await spliceOne(el);
    expect(report.deckSlides).toBe(3);

    const out = await Pkg.open(report.base64);
    const listed = await out.slidePaths();
    // ONE, and it is the rebuild. Everything else is still in the zip and still
    // related from the presentation; it is merely unlisted, which is what probe
    // question 1 measured the host accepting.
    expect(listed).toEqual([report.slidePath]);
    expect(out.has("ppt/slides/slide1.xml")).toBe(true);
    expect(out.has("ppt/slides/slide2.xml")).toBe(true);
  });

  it("is a package the integrity checker finds nothing wrong with", async () => {
    const report = await spliceOne(element("hvid-kasse-2x1-vertikale"));
    const bytes = await (await Pkg.open(report.base64)).toBytes();
    expect(problems(await partsOf(bytes))).toEqual([]);
  });

  it("puts the element's shapes on the rebuilt slide and nowhere else", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    const report = await spliceOne(el, { group: false });
    const spTree = await rebuiltTree(report.base64, report.slidePath);
    // The destination's own shape, plus the element's.
    expect(topLevelOf(spTree).length).toBe(1 + el.shapes);
    expect(report.shapes).toBe(el.shapes);
    expect(report.grouped).toBe(false);
  });

  it("gives every shape on the rebuilt slide an id of its own", async () => {
    // Two shapes sharing a `<p:cNvPr id>` is a slide PowerPoint opens with one
    // of them missing, and it is the first thing that goes wrong when markup is
    // moved between slides.
    const report = await spliceOne(element("hvid-kasse-2x1-vertikale"), { group: false });
    const ids = everyShapeId(await rebuiltTree(report.base64, report.slidePath));
    expect(ids.length).toBeGreaterThan(3);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("wraps the shapes in one group when asked, and tags the group", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    const report = await spliceOne(el, { group: true });
    expect(report.grouped).toBe(true);
    expect(report.shapes).toBe(1);

    const out = await Pkg.open(report.base64);
    const tagged = await readShapeTags(out, report.slidePath);
    expect(tagged.map((t) => t.element)).toEqual([el.id]);
    expect(tagged[0]?.catalogue).toBe("test-version");

    const spTree = await rebuiltTree(report.base64, report.slidePath);
    const group = topLevelOf(spTree).find((s) => s.localName === "grpSp");
    expect(group, "the element did not land as a group").toBeDefined();
  });

  it("tags every shape when they land loose, so the deck can still be read back", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    const report = await spliceOne(el, { group: false });
    const out = await Pkg.open(report.base64);
    const tagged = await readShapeTags(out, report.slidePath);
    expect(tagged.length).toBe(el.shapes);
    expect(new Set(tagged.map((t) => t.element))).toEqual(new Set([el.id]));
  });

  it("refuses a slide number the deck does not have, by name", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    await expect(spliceOne(el, { slide: 7 })).rejects.toThrow(/slide 8 is not in this deck, which has 3/);
    await expect(spliceOne(el, { slide: -1 })).rejects.toThrow(/is not in this deck/);
  });

  it("refuses an element whose carried part the store does not hold, rather than shipping a dangling reference", async () => {
    const el = element("en-kasse");
    expect(el.markup.rels.length).toBeGreaterThan(0);
    await expect(
      splice({
        deck: await destination(),
        slide: 1,
        element: asSplice(el),
        options: { target: "onto", group: true, colours: "deck" },
        catalogue: catalogueFor(),
        store: () => Promise.resolve(undefined),
      }),
    ).rejects.toThrow(/the catalogue has no part/);
  });
});

describe("what the element brings with it", () => {
  it("copies a picture under a name the destination is not using, and declares it", async () => {
    // `en-kasse` carries an embedded object and the picture PowerPoint draws in
    // its place: two relationships, two parts, one of them binary.
    const el = element("en-kasse");
    const report = await spliceOne(el);
    expect(report.parts).toBeGreaterThanOrEqual(2);

    const out = await Pkg.open(report.base64);
    const rels = await out.doc(Pkg.relsPathFor(report.slidePath));
    const targets = elements(rels, "http://schemas.openxmlformats.org/package/2006/relationships", "Relationship").map(
      (r) => r.getAttribute("Target") ?? "",
    );
    const carried = targets.filter((t) => t.includes("/media/") || t.includes("/embeddings/"));
    expect(carried.length).toBeGreaterThanOrEqual(2);
    for (const target of carried) {
      const path = out.resolved(report.slidePath, target);
      expect(out.has(path), `${path} is not in the package`).toBe(true);
      expect(await out.contentTypeOf(path), `${path} has no content type`).toBeDefined();
    }
  });

  it("carries a chart's own workbook and repoints the chart at the copy", async () => {
    const el = library.catalogue.elements.find((e) => e.markup.rels.some((r) => r.type.endsWith("/chart")));
    expect(el, "the library no longer has a chart to test with").toBeDefined();
    const report = await spliceOne(el as CatalogueElement);

    const out = await Pkg.open(report.base64);
    const charts = out.partNames().filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n));
    expect(charts.length).toBe(1);
    const chart = charts[0] as string;
    // The chart's OWN relationship, rewritten to the workbook's new name in
    // this package rather than left pointing at the library's.
    const related = await out.relatedParts(chart);
    expect(related.length).toBe(1);
    expect(out.has(related[0] as string)).toBe(true);
    expect(problems(await partsOf(await out.toBytes()))).toEqual([]);
  });

  it("copies a shared part once when one element names it twice", async () => {
    const el = element("en-kasse");
    const report = await spliceOne(el);
    const out = await Pkg.open(report.base64);
    const media = out.partNames().filter((n) => n.startsWith("ppt/media/"));
    // The fixture deck carries none of its own, so every one here is the
    // element's, and a part copied twice would show up as two.
    expect(media.length).toBe(new Set(media).size);
  });
});

describe("as a new slide", () => {
  it("empties the slide it was cloned from, keeping its placeholders", async () => {
    const el = element("hvid-kasse-2x1-vertikale");
    const report = await spliceOne(el, { target: "new", group: true });
    const spTree = await rebuiltTree(report.base64, report.slidePath);
    const texts: string[] = [];
    const walk = (node: Element): void => {
      if (node.namespaceURI === A_NS && node.localName === "t") texts.push(node.textContent ?? "");
      for (const kid of Array.from(node.childNodes)) if (kid.nodeType === 1) walk(kid as Element);
    };
    walk(spTree);
    // "Second" was the destination slide's own text. A new slide does not carry it.
    expect(texts.join(" ")).not.toContain("Second");
  });

  it("does not carry the previous slide's speaker notes", async () => {
    const deck = await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]], notes: "private note" }]);
    const report = await splice({
      deck,
      slide: 1,
      element: asSplice(element("hvid-kasse-2x1-vertikale")),
      options: { target: "new", group: true, colours: "deck" },
      catalogue: catalogueFor(),
      store,
    });
    const out = await Pkg.open(report.base64);
    const related = await out.relatedParts(report.slidePath);
    expect(related.filter((p) => p.startsWith("ppt/notesSlides/"))).toEqual([]);
  });
});

describe("the sweep: every element in the 16:9 library", () => {
  it("splices into a real deck without a single package problem, and never repeats a shape id", async () => {
    // The id half is HERE rather than on one element, and that is the whole
    // reason it means anything. The fixture destination's shapes are numbered
    // 1, 2 and 9; a single library element chosen to test against may use none
    // of those, so a check on one element passes whether or not the splice
    // renumbers anything. Across all 117 it does not: the collision is real,
    // and switching the renumber off turns this red.
    const deck = await destination();
    const failures: string[] = [];
    let collisionsPossible = 0;
    for (const el of library.catalogue.elements) {
      try {
        const report = await splice({
          deck,
          slide: 1,
          element: asSplice(el),
          options: { target: "onto", group: true, colours: "deck" },
          catalogue: catalogueFor(),
          store,
        });
        const found = problems(await partsOf(await (await Pkg.open(report.base64)).toBytes()));
        if (found.length) failures.push(`${el.id}: ${found.slice(0, 3).join("; ")}`);

        const authored = [...el.markup.xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => Number(m[1]));
        if (authored.some((n) => n === 1 || n === 2 || n === 9)) collisionsPossible += 1;
        const ids = everyShapeId(await rebuiltTree(report.base64, report.slidePath));
        if (new Set(ids).size !== ids.length) failures.push(`${el.id}: two shapes share an id`);
      } catch (e) {
        failures.push(`${el.id}: threw ${e instanceof Error ? e.message : String(e)}`);
      }
    }
    expect(failures).toEqual([]);
    // The vacuity guard on the guard: if no element in the library shares an id
    // with the destination, the check above proves nothing and should be
    // rewritten rather than trusted.
    expect(collisionsPossible, "no library element collides with the fixture's ids").toBeGreaterThan(0);
  }, 600_000);

  it("leaves every element's own tag on the slide, so the deck reads back", async () => {
    const deck = await destination();
    const missing: string[] = [];
    for (const el of library.catalogue.elements.slice(0, 20)) {
      const report = await splice({
        deck,
        slide: 1,
        element: asSplice(el),
        options: { target: "onto", group: true, colours: "deck" },
        catalogue: catalogueFor(),
        store,
      });
      const out = await Pkg.open(report.base64);
      const tagged = await readShapeTags(out, report.slidePath);
      if (!tagged.some((t) => t.element === el.id)) missing.push(el.id);
    }
    expect(missing).toEqual([]);
  }, 300_000);

  it("never leaves a placeholder claim on a spliced shape", async () => {
    // A library shape claiming to BE the destination's body would inherit the
    // layout's geometry and jump away from where it was landed.
    const deck = await destination();
    const claims: string[] = [];
    for (const el of library.catalogue.elements.slice(0, 30)) {
      const report = await splice({
        deck,
        slide: 1,
        element: asSplice(el),
        options: { target: "onto", group: false, colours: "deck" },
        catalogue: catalogueFor(),
        store,
      });
      const spTree = await rebuiltTree(report.base64, report.slidePath);
      const added = topLevelOf(spTree).slice(1);
      for (const shape of added) {
        const found = parseXml(`<w>${new XMLSerializer().serializeToString(shape as never)}</w>`);
        if (elements(found, P_NS, "ph").length > 0) claims.push(el.id);
      }
    }
    expect(claims).toEqual([]);
  }, 300_000);
});

describe("the tag keys the deck is read back by", () => {
  it("are the ones the design names, uppercase", () => {
    expect(TAG_ELEMENT).toBe("SSF_SLIDE_ELEMENT");
    expect(TAG_CATALOGUE).toBe("SSF_SLIDE_ELEMENTS_CATALOGUE");
  });
});

describe("the package Undo hands back", () => {
  it("carries the user's own slide, byte for byte, with only that one listed", async () => {
    const deck = await destination();
    const { base64, path } = await onlySlide(deck, 1);
    const out = await Pkg.open(base64);
    expect(await out.slidePaths()).toEqual([path]);
    expect(path).toBe("ppt/slides/slide2.xml");

    // Nothing is rebuilt and nothing is cloned: the slide that comes back is
    // the one the user had, which is the whole point of an undo. Compared as
    // TEXT against the same slide in the deck it came from.
    const before = await Pkg.open(deck);
    expect(await out.text(path)).toBe(await before.text(path));
    expect(problems(await partsOf(await out.toBytes()))).toEqual([]);
  });

  it("refuses a slide number the deck does not have, by name", async () => {
    await expect(onlySlide(await destination(), 9)).rejects.toThrow(/slide 10 is not in this deck, which has 3/);
  });
});

describe("as a new slide, in detail", () => {
  it("keeps a placeholder but empties every paragraph in it", async () => {
    // A `<p:txBody>` with no `<a:p>` at all is schema-invalid and PowerPoint
    // reports the file as damaged without naming the part, so the first
    // paragraph is emptied rather than removed.
    const deck = await makeDeck([
      { paragraphs: [["First"]] },
      { paragraphs: [["Line one"], ["Line two"], ["Line three"]] },
    ]);
    const report = await splice({
      deck,
      slide: 1,
      element: asSplice(element("hvid-kasse-2x1-vertikale")),
      options: { target: "new", group: true, colours: "deck" },
      catalogue: catalogueFor(),
      store,
    });
    const spTree = await rebuiltTree(report.base64, report.slidePath);
    const xml = new XMLSerializer().serializeToString(spTree as never);
    expect(xml).not.toContain("Line two");
    expect(xml).not.toContain("Line three");
    // Still a legal text body: at least one paragraph survives in the
    // placeholder that is left.
    expect(elements(parseXml(`<w>${xml}</w>`), A_NS, "p").length).toBeGreaterThan(0);
    expect(problems(await partsOf(await (await Pkg.open(report.base64)).toBytes()))).toEqual([]);
  });

  it("leaves the deck's own slides untouched, and only re-serialises the one it read", async () => {
    // The splice only ever ADDS a slide to the package it hands over. The
    // user's own slides are unlisted, never edited: if that stopped being true,
    // the undo above would be putting back a slide the splice had changed.
    //
    // Two of the three come back byte-identical. The THIRD does not, and the
    // reason is worth writing down rather than relaxing the test over: the
    // splice parses the slide the user is on to find its title placeholder, and
    // every parsed part is written back when the package is serialised. What
    // comes back differs only in how an empty element is spelled — the fixture
    // writes `<p:nvPr></p:nvPr>` and the serialiser writes `<p:nvPr/>` — which
    // is the same XML. It reaches nobody either way, because an unlisted slide
    // is not what PowerPoint imports.
    const deck = await destination();
    const before = await Pkg.open(deck);
    const asXml = (xml: string): string => serializeXml(parseXml(xml));
    for (const target of ["onto", "new"] as const) {
      const report = await splice({
        deck,
        slide: 1,
        element: asSplice(element("hvid-kasse-2x1-vertikale")),
        options: { target, group: true, colours: "deck" },
        catalogue: catalogueFor(),
        store,
      });
      const out = await Pkg.open(report.base64);
      for (const path of ["ppt/slides/slide1.xml", "ppt/slides/slide3.xml"]) {
        expect(await out.text(path), `${target} changed ${path}`).toBe(await before.text(path));
      }
      const read = "ppt/slides/slide2.xml";
      expect(asXml(await out.text(read)), `${target} changed ${read}`).toBe(asXml(await before.text(read)));
    }
  });
});

describe("what the splice refuses", () => {
  it("refuses a deck with no slides, rather than building a package with nothing in it", async () => {
    // `makeDeck([])` produces a presentation with an empty `<p:sldIdLst>`.
    await expect(
      splice({
        deck: await makeDeck([]),
        slide: 0,
        element: asSplice(element("hvid-kasse-2x1-vertikale")),
        options: { target: "onto", group: true, colours: "deck" },
        catalogue: catalogueFor(),
        store,
      }),
    ).rejects.toThrow(/no slides to insert into/);
  });

  it("refuses a slide number that is not a whole one", async () => {
    await expect(spliceOne(element("hvid-kasse-2x1-vertikale"), { slide: 1.5 })).rejects.toThrow(/is not in this deck/);
  });

  it("refuses an element with no shapes in its markup", async () => {
    const empty = { ...asSplice(element("hvid-kasse-2x1-vertikale")), markup: { xml: "", rels: [] } };
    await expect(
      splice({
        deck: await destination(),
        slide: 1,
        element: empty,
        options: { target: "onto", group: true, colours: "deck" },
        catalogue: catalogueFor(),
        store,
      }),
    ).rejects.toThrow(/has no shapes/);
  });
});

describe("where an element lands, through the whole splice", () => {
  it("puts a stamp top-right rather than where the library drew it", async () => {
    const stamp = library.catalogue.elements.find((e) => e.landing === "top-right");
    expect(stamp, "the library no longer has a stamp to test with").toBeDefined();
    const el = stamp as CatalogueElement;
    const report = await spliceOne(el, { group: false });
    // Right of the middle and in the top half: the rule, not the authored spot.
    expect(report.landed.x + report.landed.cx / 2).toBeGreaterThan(12192000 / 2);
    expect(report.landed.y).toBeLessThan(6858000 / 2);
  });

  it("centres a cursor element on the selected shape when the host named one", async () => {
    const cursor = library.catalogue.elements.find((e) => e.landing === "cursor");
    expect(cursor, "the library no longer has a cursor-landing element").toBeDefined();
    const selection = { x: 6000000, y: 3000000, cx: 400000, cy: 300000 };
    const report = await splice({
      deck: await destination(),
      slide: 1,
      element: asSplice(cursor as CatalogueElement),
      options: { target: "onto", group: false, colours: "deck" },
      catalogue: catalogueFor(),
      store,
      selection,
    });
    const centre = { x: report.landed.x + report.landed.cx / 2, y: report.landed.y + report.landed.cy / 2 };
    expect(Math.abs(centre.x - (selection.x + selection.cx / 2))).toBeLessThan(20000);
    expect(Math.abs(centre.y - (selection.y + selection.cy / 2))).toBeLessThan(20000);
  });

  it("moves the shapes to where it says it landed them", async () => {
    // The report is what the pane shows and the suite asserts; if the markup
    // and the report could disagree, every landing test above would be about
    // arithmetic nobody applied.
    const stamp = library.catalogue.elements.find((e) => e.landing === "top-right") as CatalogueElement;
    const report = await spliceOne(stamp, { group: false });
    const spTree = await rebuiltTree(report.base64, report.slidePath);
    const added = topLevelOf(spTree).slice(1);
    const xs = added
      .map((shape) => {
        const found = parseXml(`<w>${new XMLSerializer().serializeToString(shape as never)}</w>`);
        return Number(elements(found, A_NS, "off")[0]?.getAttribute("x") ?? NaN);
      })
      .filter((n) => Number.isFinite(n));
    expect(xs.length).toBeGreaterThan(0);
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(report.landed.x - 1);
  });
});

describe("a slide that carries a comment", () => {
  /**
   * A modern comment as PowerPoint for the web writes one.
   *
   * The shape matters, not just the presence: the comment part hangs off the
   * SLIDE under a Microsoft-namespaced relationship, and the slide ANCHORS it
   * from its own extension list. That anchor is why `cloneSlide` keeps the
   * relationship — its drop pass only removes what nothing names — and keeping
   * it is right for a slide being rebuilt and wrong for one being made new.
   *
   * Measured on the web on 2026-09-10 before it was written down here: a slide
   * with one comment, replaced twice and then used as the basis for a new
   * slide, came back with the comment on both.
   */
  const COMMENT_REL = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
  const COMMENT_TYPE = "application/vnd.ms-powerpoint.comments+xml";

  async function deckWithComment(): Promise<Uint8Array> {
    const pkg = await Pkg.open(await destination());
    const slide = (await pkg.slidePaths())[1] as string;
    const part = "ppt/comments/modernComment_101_TEST.xml";
    pkg.setText(part, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p188:cmLst/>');
    await pkg.addContentTypeOverride(`/${part}`, COMMENT_TYPE);
    const rId = await pkg.addRel(slide, COMMENT_REL, "../comments/modernComment_101_TEST.xml");

    // The anchor, in the slide's own extension list, naming the relationship.
    const doc = await pkg.doc(slide);
    const cSld = child(doc.documentElement, P_NS, "cSld");
    if (!cSld) throw new Error("the fixture slide has no cSld");
    const extLst = doc.createElementNS(P_NS, "p:extLst");
    const ext = doc.createElementNS(P_NS, "p:ext");
    ext.setAttribute("uri", "{575DDBBC-B0F2-4D9E-9E05-8B25F5ED30E1}");
    const anchor = doc.createElementNS("http://schemas.microsoft.com/office/powerpoint/2018/8/main", "p188:commentRel");
    anchor.setAttributeNS(R_NS, "r:id", rId);
    ext.appendChild(anchor);
    extLst.appendChild(ext);
    cSld.appendChild(extLst);
    return pkg.toBytes();
  }

  async function commentsOn(base64: string, slidePath: string): Promise<string[]> {
    const out = await Pkg.open(base64);
    return (await out.relatedParts(slidePath)).filter((p) => p.startsWith("ppt/comments/"));
  }

  it("keeps the comment when the slide is REBUILT, because it is still that slide", async () => {
    // Losing a reviewer's thread because somebody dropped a box on the slide is
    // the worst thing this add-in could quietly do.
    const report = await splice({
      deck: await deckWithComment(),
      slide: 1,
      element: asSplice(element("hvid-kasse-2x1-vertikale")),
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: catalogueFor(),
      store,
    });
    expect(await commentsOn(report.base64, report.slidePath)).toHaveLength(1);
  });

  it("drops it for a NEW slide, so a review thread is not duplicated", async () => {
    const report = await splice({
      deck: await deckWithComment(),
      slide: 1,
      element: asSplice(element("hvid-kasse-2x1-vertikale")),
      options: { target: "new", group: true, colours: "deck" },
      catalogue: catalogueFor(),
      store,
    });
    expect(await commentsOn(report.base64, report.slidePath)).toEqual([]);
    // And the package is still whole: the comment part is left behind as an
    // orphan rather than deleted out from under a relationship.
    expect(problems(await partsOf(await (await Pkg.open(report.base64)).toBytes()))).toEqual([]);
  });

  it("drops the previous slide's notes for a new slide too, for the same reason", async () => {
    const deck = await makeDeck([{ paragraphs: [["First"]] }, { paragraphs: [["Second"]], notes: "private note" }]);
    const report = await splice({
      deck,
      slide: 1,
      element: asSplice(element("hvid-kasse-2x1-vertikale")),
      options: { target: "new", group: true, colours: "deck" },
      catalogue: catalogueFor(),
      store,
    });
    const out = await Pkg.open(report.base64);
    expect((await out.relatedParts(report.slidePath)).filter((p) => p.startsWith("ppt/notesSlides/"))).toEqual([]);
  });
});

describe("the colour switch", () => {
  /** The rebuilt slide, serialised: what PowerPoint would draw. */
  async function slideXml(base64: string, slidePath: string): Promise<string> {
    const out = await Pkg.open(base64);
    return serializeXml(await out.doc(slidePath));
  }

  const WITH_ACCENT = "sort-streg-3-vertikale-kausalitet";

  it("leaves every scheme colour alone under This deck's theme, which is what makes the element take the deck", async () => {
    const report = await spliceOne(element(WITH_ACCENT));
    const xml = await slideXml(report.base64, report.slidePath);
    expect(xml).toContain(`<a:schemeClr val="accent1"`);
    expect(report.pinned).toBe(0);
  });

  it("pins them to the library's own values under As in the library", async () => {
    const report = await spliceOne(element(WITH_ACCENT), { colours: "library" });
    const xml = await slideXml(report.base64, report.slidePath);
    // The 16:9 library is on the stock Office palette, measured 2026-09-11:
    // accent1 is 5B9BD5 and tx1 is black, through a sysClr the theme states as
    // `windowText`.
    expect(xml).toContain(`<a:srgbClr val="5B9BD5"`);
    expect(xml).not.toContain("schemeClr");
    expect(report.pinned).toBeGreaterThan(0);
  });

  it("pins the colours inside a CARRIED part too, so a chart does not disagree with the shapes around it", async () => {
    // The library's one chart states 17 scheme colours of its own. A switch
    // that rewrote the markup and not the part it carries would leave the
    // series following the destination's theme and the frame around it not.
    const deck = element("bridging-numbers");
    const chartOf = async (base64: string): Promise<string> => {
      const out = await Pkg.open(base64);
      const path = out.partNames().find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p));
      if (!path) throw new Error("the package carries no chart");
      return out.text(path);
    };
    const following = await spliceOne(deck);
    expect(await chartOf(following.base64)).toContain("schemeClr");

    const pinnedReport = await spliceOne(deck, { colours: "library" });
    const chart = await chartOf(pinnedReport.base64);
    expect(chart).not.toContain("schemeClr");
    expect(chart).toContain(`<a:srgbClr val="5B9BD5"`);
    // And the package is still one PowerPoint would open.
    expect(problems(await partsOf(await (await Pkg.open(pinnedReport.base64)).toBytes()))).toEqual([]);
  });
});
