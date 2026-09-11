import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  HarvestError,
  Pkg,
  boxOf,
  countKeys,
  countedNoun,
  harvest,
  occupiedBoxes,
  parseXml,
  partName,
  sizeRuns,
  slug,
  tagsFor,
  union,
} from "../src/core/index.js";
import { A_NS, P_NS, element } from "../src/core/pptx/xml.js";
import type { Names } from "../src/core/index.js";
import { makeDeck, type SlideSpec } from "./fixtures/deck.js";

/**
 * The harvest: a library deck in, a catalogue out.
 *
 * Two halves. The synthetic decks prove each RULE in isolation (what a heading
 * is, what the marker does, what is chrome, where a rotated shape is), and the
 * committed decks prove the rules add up to the library the owner authored: the
 * counts here were measured against the owner's print on 2026-09-08 and are
 * what `docs/DESIGN.md` says the library holds.
 */

const W = 12192000;
const H = 6858000;

/** A plain rectangle at (x, y) of size (w, h) EMU, optionally rotated (degrees) and named. */
function rect(
  x: number,
  y: number,
  w: number,
  h: number,
  options: { rot?: number; name?: string; text?: string; id?: number } = {},
): string {
  const rot = options.rot === undefined ? "" : ` rot="${Math.round(options.rot * 60000)}"`;
  const text =
    options.text === undefined
      ? ""
      : `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="da-DK"/><a:t>${options.text}</a:t></a:r></a:p></p:txBody>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${options.id ?? 20}" name="${options.name ?? "Rektangel 5"}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${text}</p:sp>`
  );
}

/** A picture whose blip names the relationship the fixture's `icons` option creates. */
function picture(rId: string, x = 1000000, y = 1000000): string {
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="31" name="Billede 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="800000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}

/** A table whose frame is narrower than its columns add up to, the way PowerPoint writes one. */
function table(frameW: number, cols: number[], rows: number[]): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="40" name="Tabel 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="${frameW}" cy="${rows[0] ?? 0}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>` +
    cols.map((w) => `<a:gridCol w="${w}"/>`).join("") +
    `</a:tblGrid>` +
    rows
      .map(
        (h) => `<a:tr h="${h}"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody></a:tc></a:tr>`,
      )
      .join("") +
    `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

const heading = (title: string, notes?: string): SlideSpec => ({ paragraphs: [], title, noBody: true, notes });

const NAMES: Names = {
  categories: { Kasser: "Boxes", Stempler: "Stamps and labels", Figurer: "Shapes", Samling: "Collection" },
  names: {
    "Kasse, 2 vertikale": "Boxes, 2 vertical",
    "Kasse, 3 vertikale": "Boxes, 3 vertical",
    Tabel: "Table",
    Udkast: "Draft stamp",
    Bred: "Wide bar",
    Dokument: "Document",
    "Figurer 1": "Shape 1",
    "Stempler 1": "Scales",
    Billede: "Picture",
  },
};

async function harvested(slides: SlideSpec[], names: Names = NAMES) {
  const pkg = await Pkg.open(await makeDeck(slides));
  return harvest(pkg, { size: "16:9", names });
}

describe("headings and elements", () => {
  it("makes a heading of a slide with a title and nothing else, and an element of every slide under it", async () => {
    const { catalogue } = await harvested([
      { paragraphs: [["Cover"]] },
      heading("Kasser"),
      {
        paragraphs: [["hello"]],
        title: "Kasse, 2 vertikale",
        box: `<a:xfrm><a:off x="1000000" y="1500000"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>`,
      },
      {
        paragraphs: [["hello"]],
        title: "Kasse, 3 vertikale",
        box: `<a:xfrm><a:off x="1000000" y="1500000"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>`,
      },
    ]);
    expect(catalogue.categories).toEqual([{ key: "Kasser", name: "Boxes" }]);
    expect(catalogue.elements.map((e) => [e.key, e.name, e.kind, e.slide, e.landing])).toEqual([
      ["Kasse, 2 vertikale", "Boxes, 2 vertical", "slide", 3, "layout"],
      ["Kasse, 3 vertikale", "Boxes, 3 vertical", "slide", 4, "layout"],
    ]);
    // The cover before the first heading is not an element.
    expect(catalogue.elements.some((e) => e.slide === 1)).toBe(false);
    expect(catalogue.width).toBe(W);
    expect(catalogue.height).toBe(H);
  });

  it("gives the two-and-three run one key, a noun and a count each", async () => {
    const { catalogue } = await harvested([
      heading("Kasser"),
      { paragraphs: [["a"]], title: "Kasse, 2 vertikale" },
      { paragraphs: [["b"]], title: "Kasse, 3 vertikale" },
    ]);
    expect(catalogue.elements.map((e) => e.run)).toEqual([
      { key: "Boxes, N vertical", noun: "boxes", count: 2 },
      { key: "Boxes, N vertical", noun: "boxes", count: 3 },
    ]);
  });

  it("ids are slugs of the key, stable and ascii", async () => {
    const { catalogue } = await harvested([heading("Kasser"), { paragraphs: [["a"]], title: "Kasse, 2 vertikale" }]);
    expect(catalogue.elements[0]?.id).toBe("kasse-2-vertikale");
    expect(slug("Grå kasse: Æbler & øl")).toBe("graa-kasse-aebler-oel");
  });

  it("boxes the element around its content and leaves the title, an empty placeholder and an off-slide shape out", async () => {
    const empty = `<p:sp><p:nvSpPr><p:cNvPr id="50" name="Content 4"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;
    const { catalogue } = await harvested([
      heading("Kasser"),
      {
        paragraphs: [],
        noBody: true,
        title: "Kasse, 2 vertikale",
        shapes: [
          rect(1219200, 685800, 2438400, 1371600),
          rect(6096000, 685800, 2438400, 2743200, { id: 21 }),
          rect(W + 100, 0, 1000000, 1000000, { id: 22 }),
          empty,
        ],
      },
    ]);
    const [el] = catalogue.elements;
    expect(el?.shapes).toBe(2);
    expect(el?.box).toEqual({ x: 0.1, y: 0.1, w: 0.6, h: 0.4 });
  });

  it("fails naming every key and category without an English name, and a key used twice", async () => {
    const run = harvested(
      [
        heading("Ukendt"),
        { paragraphs: [["a"]], title: "Kasse, 2 vertikale" },
        { paragraphs: [["b"]], title: "Kasse, 2 vertikale" },
        { paragraphs: [["c"]], title: "Navnløs" },
      ],
      NAMES,
    );
    await expect(run).rejects.toBeInstanceOf(HarvestError);
    const problems = await run.catch((e: HarvestError) => e.problems);
    expect(problems).toEqual([
      'category "Ukendt" has no English name in the names file',
      '"Kasse, 2 vertikale" is the key of both slide 2 and slide 3',
      '"Navnløs" (slide 4) has no English name in the names file',
    ]);
  });

  it("fails a slide with content but no title, and a deck without a slide size", async () => {
    await expect(harvested([heading("Kasser"), { paragraphs: [["orphan"]] }])).rejects.toMatchObject({
      problems: ["slide 2 has content but no title, so it has no key"],
    });
    const pkg = await Pkg.open(await makeDeck([heading("Kasser")]));
    const pres = await pkg.doc("ppt/presentation.xml");
    const sz = element(pres, P_NS, "sldSz");
    sz?.parentNode?.removeChild(sz);
    await expect(harvest(pkg, { size: "16:9", names: NAMES })).rejects.toThrow(/no slide size/);
  });
});

describe("collection slides", () => {
  const marker = "SSF: ét element pr. figur";
  const stampSlide = (title: string, notes?: string): SlideSpec => ({
    paragraphs: [],
    noBody: true,
    title,
    notes,
    shapes: [
      rect(2344971, 2068917, 2017713, 792163, { rot: -29.056, text: "Udkast", id: 60 }),
      rect(0, 0, 9000000, 260000, { name: "Bred", id: 61 }),
      rect(4000000, 4000000, 500000, 500000, { name: "Gruppe 7", id: 62 }),
    ],
  });

  it("splits a slide under a marked heading one element per shape, titled by the slide, named by the shape", async () => {
    const { catalogue } = await harvested([heading("Samling", marker), stampSlide("Stempler")]);
    expect(catalogue.categories).toEqual([{ key: "Stempler", name: "Stamps and labels" }]);
    expect(catalogue.elements.map((e) => [e.key, e.kind, e.landing, e.shapes, e.tags])).toEqual([
      ["Udkast", "part", "top-right", 1, ["stamps", "small"]],
      ["Bred", "part", "as-authored", 1, ["stamps", "small"]],
      ["Stempler 1", "part", "top-right", 1, ["stamps", "small"]],
    ]);
  });

  it("boxes a rotated part by its rotated extent, not the frame the XML states", async () => {
    const { catalogue } = await harvested([heading("Samling", marker), stampSlide("Stempler")]);
    const stamp = catalogue.elements.find((e) => e.key === "Udkast");
    // The frame is 2017713 x 792163 at 29°: the extent is wider and much taller.
    expect(stamp?.box.w).toBeCloseTo((2017713 * Math.cos(0.5071) + 792163 * Math.sin(0.5071)) / W, 3);
    expect(stamp?.box.h).toBeCloseTo((2017713 * Math.sin(0.5071) + 792163 * Math.cos(0.5071)) / H, 3);
    expect(stamp?.box.h).toBeGreaterThan(792163 / H);
  });

  it("lands parts at the cursor on any other collection slide", async () => {
    const { catalogue } = await harvested([
      heading("Samling", marker),
      {
        paragraphs: [],
        noBody: true,
        title: "Figurer",
        shapes: [
          rect(1000000, 1000000, 500000, 500000, { text: "[ Dokument ]" }),
          rect(2000000, 1000000, 500000, 500000, { id: 21 }),
        ],
      },
    ]);
    expect(catalogue.elements.map((e) => [e.key, e.landing])).toEqual([
      ["Dokument", "cursor"],
      ["Figurer 1", "cursor"],
    ]);
  });

  it("lets a slide's own notes override the heading, both ways", async () => {
    const { catalogue } = await harvested([
      heading("Samling", marker),
      { ...stampSlide("Tabel", "SSF: ét element pr. dias") },
      heading("Kasser"),
      {
        paragraphs: [],
        noBody: true,
        title: "Figurer",
        notes: "SSF: one element per shape",
        shapes: [rect(1000000, 1000000, 500000, 500000, { text: "Dokument" })],
      },
    ]);
    expect(catalogue.elements.map((e) => [e.key, e.kind])).toEqual([
      ["Tabel", "slide"],
      ["Dokument", "part"],
    ]);
  });
});

describe("markup and carried parts", () => {
  it("carries the picture a shape names and lists the relationship it needs, and nothing the slide merely has", async () => {
    // `icons: true` gives the slide two image relationships; the shape here uses one of them.
    const { catalogue, parts } = await harvested([
      heading("Kasser"),
      { paragraphs: [], noBody: true, title: "Billede", icons: true, shapes: [picture("rId20")] },
    ]);
    // The icon shapes the fixture adds are content too, so the element carries their pictures as well.
    const [el] = catalogue.elements;
    expect(el?.markup.xml).toContain('r:embed="rId20"');
    expect(el?.markup.rels.map((r) => r.id)).toEqual(["rId20", "rId21", "rId22"]);
    expect(el?.markup.parts.length).toBeGreaterThan(0);
    for (const part of el?.markup.parts ?? []) expect(part).toMatch(/^ppt\/media\//);
    // The notes page and the layout the slide also relates to are not carried.
    expect([...parts.keys()].sort()).toEqual([...(el?.markup.parts ?? [])].sort());
  });

  it("reaches a chart's workbook through the chart", async () => {
    const { catalogue, parts } = await harvested([
      heading("Kasser"),
      {
        paragraphs: [],
        noBody: true,
        title: "Tabel",
        chart: { title: "Sales", workbook: ["a"] },
        shapes: [
          `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="70" name="Diagram 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="4000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3"/></a:graphicData></a:graphic></p:graphicFrame>`,
        ],
      },
    ]);
    const [el] = catalogue.elements;
    expect(el?.markup.parts.some((p) => p.startsWith("ppt/charts/"))).toBe(true);
    expect(el?.markup.parts.some((p) => p.startsWith("ppt/embeddings/"))).toBe(true);
    expect(parts.has("ppt/charts/_rels/chart2.xml.rels")).toBe(true);
  });
});

describe("boxes", () => {
  const shape = (xml: string): Element => {
    const doc = parseXml(
      `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${xml}</p:spTree>`,
    );
    return doc.documentElement?.firstChild as Element;
  };

  it("takes a table's size from its columns and rows when the frame is smaller", () => {
    const box = boxOf(shape(table(1000000, [2000000, 2000000, 2000000], [400000, 400000])), W, H);
    expect(box).toEqual({ x: 1000000 / W, y: 2000000 / H, w: 6000000 / W, h: 800000 / H });
  });

  it("keeps the frame when it is the larger", () => {
    const box = boxOf(shape(table(9000000, [2000000], [400000])), W, H);
    expect(box?.w).toBe(9000000 / W);
  });

  it("answers undefined for a shape with no frame, and unions boxes", () => {
    expect(
      boxOf(shape(`<p:sp><p:nvSpPr><p:cNvPr id="1" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`), W, H),
    ).toBeUndefined();
    expect(union([])).toEqual({ x: 0, y: 0, w: 1, h: 1 });
    expect(
      union([
        { x: 0.1, y: 0.1, w: 0.2, h: 0.2 },
        { x: 0.5, y: 0.3, w: 0.2, h: 0.4 },
      ]),
    ).toEqual({ x: 0.1, y: 0.1, w: 0.6, h: 0.6 });
  });

  it("boxes a group by its own frame", () => {
    const grp = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Gruppe 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="2438400" cy="1371600"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></p:grpSpPr>${rect(0, 0, 5, 5)}</p:grpSp>`;
    expect(boxOf(shape(grp), W, H)).toEqual({ x: 0.1, y: 0.1, w: 0.2, h: 0.2 });
  });
});

describe("names, runs and tags", () => {
  const shape = (xml: string): Element => {
    const doc = parseXml(`<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${xml}</p:spTree>`);
    return doc.documentElement?.firstChild as Element;
  };

  it("names a part by its text with brackets stripped and cut at a colon, else its shape name, else the slide numbered", () => {
    const used = new Set<string>();
    const nameless = { count: 0 };
    expect(
      partName(shape(rect(0, 0, 1, 1, { text: "[ Dokumentnavn ] > [ Afsnit ]: mere" })), "S", used, nameless),
    ).toBe("Dokumentnavn > Afsnit");
    expect(partName(shape(rect(0, 0, 1, 1, { text: "Tekst", name: "Kicker box" })), "S", used, nameless)).toBe(
      "Kicker box",
    );
    expect(partName(shape(rect(0, 0, 1, 1, { name: "Gruppe 12" })), "S", used, nameless)).toBe("S 1");
    expect(partName(shape(rect(0, 0, 1, 1, { name: "Group 3" })), "S", used, nameless)).toBe("S 2");
    expect(partName(shape(rect(0, 0, 1, 1, { text: "Kicker box" })), "S", used, nameless)).toBe("Kicker box (2)");
    const long = "Kommentering af indholdselement: Indholdselement til venstre".replace(":", ";");
    expect(partName(shape(rect(0, 0, 1, 1, { text: long })), "S", used, nameless)).toBe(
      "Kommentering af indholdselement…",
    );
  });

  it("reads one counting number out of a name and never one inside 2×2 or 1-2-3", () => {
    expect(countKeys("Process flow, horizontal, 3 boxes with table")).toEqual([
      { key: "Process flow, horizontal, N boxes with table", n: 3 },
    ]);
    expect(countKeys("Process flow, horizontal, 1 box with table")).toEqual([
      { key: "Process flow, horizontal, N boxes with table", n: 1 },
    ]);
    expect(countKeys("Matrix, three rows, boxes with arrows")).toEqual([
      { key: "Matrix, N rows, boxes with arrows", n: 3 },
    ]);
    expect(countKeys("White boxes, 2×2 vertical")).toEqual([]);
    expect(countKeys("Hierarchy, 3 levels, 1-3-2 boxes")).toEqual([{ key: "Hierarchy, N levels, 1-3-2 boxes", n: 3 }]);
    expect(countKeys("Process flow, 2 boxes, 3 areas each with key figures")).toEqual([
      { key: "Process flow, N boxes with key figures", n: 2 },
    ]);
    expect(countKeys("Process flow, 1 box and 6 with key figures")).toEqual([
      { key: "Process flow, N boxes with key figures", n: 1 },
    ]);
  });

  it("groups a run only with distinct counts, and names what it counts", () => {
    const runs = sizeRuns([
      "Black lines, 2 vertical",
      "Black lines, 3 vertical",
      "Black lines, 2 vertical, causality",
      "Table, 8 rows and 9 columns",
      "Table, 8 rows and 9 columns, cells cannot be coloured",
    ]);
    expect(runs.get("Black lines, 2 vertical")).toEqual({ key: "Black lines, N vertical", noun: "lines", count: 2 });
    expect(runs.has("Black lines, 2 vertical, causality")).toBe(false);
    expect(runs.has("Table, 8 rows and 9 columns")).toBe(false);
    expect(countedNoun("Matrix, N rows, boxes with arrows")).toBe("rows");
    expect(countedNoun("Two vertical boxes".replace("Two", "N"))).toBe("boxes");
    expect(countedNoun("N things")).toBe("items");
  });

  it("derives tags from the Danish words, and marks a part small", () => {
    expect(tagsFor("Hvid kasse, 2 vertikale med kausalitet", "Hvide kasser")).toEqual([
      "boxes",
      "causality",
      "columns",
      "white",
    ]);
    expect(tagsFor("Draft", "Stempler og lignende", true)).toEqual(["stamps", "small"]);
  });

  it("reaches an English part key as well as its Danish spelling", () => {
    // A part is keyed by its own text on the slide, and that text went English
    // on 2026-09-11. A rule written only in Danish stops reaching these, which
    // is how four elements silently lost their "meeting" tag.
    for (const key of ["Document", "Documents", "Working document", "Document name > Current section from…"]) {
      expect(tagsFor(key, "Stempler og lignende", true)).toContain("meeting");
    }
    // the Danish spellings still reach it, for any deck that has not moved
    expect(tagsFor("Arbejdsdokument", "Stempler og lignende", true)).toContain("meeting");
    // and the two that were already covered in both languages
    expect(tagsFor("Decision", "Flowchart ikoner", true)).toContain("meeting");
    expect(tagsFor("Process / Action", "Flowchart ikoner", true)).toContain("process");
  });
});

describe("the committed library", () => {
  const names = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;
  const load = async (file: string, size: "16:9" | "4:3") =>
    harvest(await Pkg.open(new Uint8Array(readFileSync(`template/${file}`))), { size, names });

  it("harvests the 16:9 deck into the library the design describes", async () => {
    const { catalogue } = await load("library-16x9.pptx", "16:9");
    const parts = catalogue.elements.filter((e) => e.kind === "part");
    expect(catalogue.elements.length).toBe(117);
    expect(parts.length).toBe(21);
    expect(new Set(parts.map((e) => e.category.name))).toEqual(
      new Set(["Markers", "Stamps and labels", "Flowchart shapes"]),
    );
    expect(new Set(catalogue.elements.filter((e) => e.run).map((e) => e.run?.key)).size).toBe(12);
    // A part is keyed by its own text on the slide, and that text is English
    // since 2026-09-11. The last one keeps a Danish key because it has no text
    // of its own and falls back to the slide title numbered, and titles stay
    // Danish (docs/DESIGN.md section 2).
    expect(catalogue.elements.filter((e) => e.landing === "top-right").map((e) => e.key)).toEqual([
      "Confidential",
      "Draft",
      "Confidential (2)",
      "Discussion paper",
      "Working document",
      "Stempler og lignende 1",
    ]);
    // Every key has an English name, and no two keys share a slug.
    expect(new Set(catalogue.elements.map((e) => e.id)).size).toBe(117);
  }, 30000);

  it("harvests the 4:3 deck into the same keys", async () => {
    const wide = await load("library-16x9.pptx", "16:9");
    const std = await load("library-4x3.pptx", "4:3");
    const keys = (c: typeof wide) => new Set(c.catalogue.elements.map((e) => e.key));
    expect(keys(std)).toEqual(keys(wide));
    expect(std.catalogue.width).toBe(9144000);
  }, 30000);
});

describe("what a destination slide already holds", () => {
  const slide = (body: string): Document =>
    parseXml(
      `<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `${body}</p:spTree></p:cSld></p:sld>`,
    );

  /** A placeholder of a given type, with or without text in it. */
  const placeholder = (type: string, text?: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="9" name="${type}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="2438400" cy="1371600"/></a:xfrm></p:spPr>` +
    (text === undefined
      ? `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>`
      : `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>`) +
    `</p:sp>`;

  it("answers a box per shape, in z-order, in fractions of the slide", () => {
    const boxes = occupiedBoxes(slide(rect(0, 0, W / 2, H / 2) + rect(W / 2, H / 2, W / 4, H / 4)), W, H);
    expect(boxes).toEqual([
      { x: 0, y: 0, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0.5, w: 0.25, h: 0.25 },
    ]);
  });

  it("leaves out the ghosts the insert is about to remove, and keeps a placeholder with something in it", () => {
    // `docs/DESIGN.md` section 6: a whole-slide element removes the empty
    // "Click to add text" placeholders it lands over. Drawing one as occupied
    // would show the user an obstacle that is about to be taken away.
    expect(occupiedBoxes(slide(placeholder("body")), W, H)).toEqual([]);
    expect(occupiedBoxes(slide(placeholder("body", "real text")), W, H)).toHaveLength(1);
    // A title is a placeholder too, and a title with a title in it is content
    // the element has to land below.
    expect(occupiedBoxes(slide(placeholder("title", "Q4 results")), W, H)).toHaveLength(1);
  });

  it("leaves out a shape with no frame of its own, rather than drawing it at the origin", () => {
    // A placeholder that inherits its geometry from the layout says nothing
    // about where it is, and a box at 0,0 would be a lie exactly where the user
    // is looking for the truth.
    const bare =
      `<p:sp><p:nvSpPr><p:cNvPr id="4" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp>`;
    expect(occupiedBoxes(slide(bare), W, H)).toEqual([]);
  });

  it("leaves out a shape parked off the slide", () => {
    expect(occupiedBoxes(slide(rect(W + 100000, 0, 500000, 500000)), W, H)).toEqual([]);
  });

  it("answers nothing for an empty slide", () => {
    expect(occupiedBoxes(slide(""), W, H)).toEqual([]);
  });
});
