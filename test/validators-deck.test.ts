import { readFileSync, statSync } from "node:fs";
import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { packageProblems } from "../scripts/package-integrity.mjs";
import { Pkg, harvest } from "../src/core/index.js";
import type { Catalogue, Names } from "../src/core/index.js";
import { libraryFor } from "../src/pane/catalogue.js";
import { splice } from "../src/core/splice/splice.js";

/**
 * The deck that goes with the AppSource submission.
 *
 * `docs/LISTING.md` promises the validators a presentation to try the add-in
 * on, and the reason it is not built by this repo is recorded there: **a deck
 * built by this project's own code is the wrong thing to hand the people
 * checking whether this project's code produces sound files.** So PowerPoint
 * authored it — driven over COM with no window, on 2026-09-11 — and this file
 * holds the committed bytes to what the listing says about them.
 *
 * A binary in a repository is the one file a reviewer cannot read, which is why
 * every claim about it is checked here rather than described.
 */
const DECK = "template/validators.pptx";

let pkg: Pkg;
let paths: string[];

beforeAll(async () => {
  pkg = await Pkg.open(new Uint8Array(readFileSync(DECK)));
  paths = await pkg.slidePaths();
});

describe("the validators' deck", () => {
  it("is a presentation PowerPoint wrote, not one this repo assembled", async () => {
    const app = await pkg.text("docProps/app.xml");
    expect(app).toContain("Microsoft Office PowerPoint");
    // A deck from `makeDeck` has no app properties at all, which is the
    // difference this case exists to hold.
    expect(app).toMatch(/<AppVersion>/);
  });

  it("holds three slides, and they say what the listing says they say", async () => {
    expect(paths.length).toBe(3);
    const texts = [];
    for (const path of paths) {
      const xml = await pkg.text(path);
      texts.push([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(" "));
    }
    expect(texts[0]).toContain("SSF Slide Elements");
    expect(texts[0]).toContain("Nothing here is confidential");
    // The second is where a validator is told to insert; the third already has
    // something on it, so the insert has something to land beside and Undo has
    // something to put back.
    expect(texts[1]).toContain("empty slide to insert onto");
    expect(texts[2]).toContain("already has a shape");
    expect(texts[2]).toContain("Already here");
  });

  it("is widescreen at the size PowerPoint gives a new deck, so no library is borrowed", async () => {
    const pres = await pkg.text("ppt/presentation.xml");
    expect(pres).toContain('cx="12192000"');
    expect(pres).toContain('cy="6858000"');
    const index = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
      version: string;
      sizes: Record<string, Catalogue>;
    };
    const library = libraryFor(index, 12192000, 6858000);
    expect(library.size).toBe("16:9");
    // A borrowed library puts a line under the pane's header. On the deck the
    // reviewers are handed, there should be nothing to explain.
    expect(library.borrowed).toBeUndefined();
  });

  it("is a package nothing objects to, and small enough to attach to a form", async () => {
    const zip = await JSZip.loadAsync(readFileSync(DECK));
    const parts = new Map<string, string | Uint8Array>();
    for (const [name, file] of Object.entries(zip.files)) {
      if (file.dir) continue;
      const text = name.endsWith(".xml") || name.endsWith(".rels");
      parts.set(name, text ? await file.async("string") : await file.async("uint8array"));
    }
    expect(packageProblems(parts)).toEqual([]);
    expect(statSync(DECK).size).toBeLessThan(500_000);
  });

  it("takes an element, which is the whole thing a validator will do with it", async () => {
    const names = JSON.parse(readFileSync("template/names.en.json", "utf8")) as Names;
    const library = await harvest(await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
      size: "16:9",
      names,
    });
    const el = library.catalogue.elements.find((e) => e.id === "hvid-kasse-2x1-vertikale");
    if (!el) throw new Error("the library has no hvid-kasse-2x1-vertikale");
    const report = await splice({
      deck: new Uint8Array(readFileSync(DECK)),
      // The slide the deck tells them to try.
      slide: 1,
      element: { id: el.id, name: el.name, kind: el.kind, box: el.box, landing: el.landing, markup: el.markup },
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: { version: "v1", carried: library.catalogue.carried, theme: library.catalogue.theme },
      store: (path) => Promise.resolve(library.parts.get(path)),
    });
    expect(report.deckSlides).toBe(3);
    const out = await Pkg.open(report.base64);
    const zip = await JSZip.loadAsync(await out.toBytes());
    const parts = new Map<string, string | Uint8Array>();
    for (const [name, file] of Object.entries(zip.files)) {
      if (file.dir) continue;
      const text = name.endsWith(".xml") || name.endsWith(".rels");
      parts.set(name, text ? await file.async("string") : await file.async("uint8array"));
    }
    expect(packageProblems(parts)).toEqual([]);
  }, 120_000);
});
