/**
 * The gate that exercises the whole product against the real library.
 *
 * Harvest output in, a package PowerPoint would open out, ninety-eight times.
 * It is slow — around fifteen seconds — and it is the only test here that would
 * have caught either of the two bugs the engine actually shipped with:
 * `<mc:AlternateContent>` branches being split apart, and 41 think-cell shape
 * tags left dangling. Both were found by running this, at 0/98.
 *
 * It reads the COMMITTED catalogue rather than re-harvesting, deliberately: the
 * committed files are what the add-in ships, and a test that regenerates them
 * first would pass against a catalogue nobody is going to install.
 */
import { readFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import "./setup-base64.js";
import { Pkg } from "../src/core/pptx/pkg.js";
import { splice } from "../src/core/element/splice.js";
import { verify } from "../src/core/element/verify.js";
import { choosePlacement } from "../src/host/placement.js";
import type { CatalogueIndex, ElementPayload } from "../src/core/catalogue/types.js";

const CATALOGUE = join(process.cwd(), "public", "catalogue");
const TEMPLATE = join(process.cwd(), "template", "Diasskabeloner-widescreen.pptx");

const index = JSON.parse(readFileSync(join(CATALOGUE, "index.json"), "utf8")) as CatalogueIndex;

describe("the committed catalogue", () => {
  it("has the sections and elements the library deck holds", () => {
    // Both numbers are the measurement this repo's whole design rests on: the
    // layout tells you what a slide is for, and it did so for every slide in
    // the deck with no exceptions. A change to either number means the library
    // changed or the rule stopped holding, and both are worth stopping for.
    expect(index.sections.length).toBe(10);
    expect(index.elements.length).toBe(98);
  });

  it("gives every element a section that exists", () => {
    const ids = new Set(index.sections.map((s) => s.id));
    for (const e of index.elements) expect(ids.has(e.section)).toBe(true);
  });

  it("ships a payload for every element in the index", () => {
    for (const e of index.elements) {
      expect(existsSync(join(CATALOGUE, "elements", `${e.id}.json`))).toBe(true);
    }
  });

  it("gives every element a preview and a non-empty box", () => {
    for (const e of index.elements) {
      expect(e.preview.startsWith("<svg")).toBe(true);
      expect(e.bounds.cx).toBeGreaterThan(0);
      expect(e.bounds.cy).toBeGreaterThan(0);
    }
  });
});

describe("every element splices into a real deck", () => {
  it("produces a slide with no dangling reference, missing part or id collision", async () => {
    const deck = new Uint8Array(readFileSync(TEMPLATE));
    const failures: string[] = [];
    for (const entry of index.elements) {
      const payload = JSON.parse(
        readFileSync(join(CATALOGUE, "elements", `${entry.id}.json`), "utf8"),
      ) as ElementPayload;
      // A fresh package per element: a splice mutates, and reusing one would
      // test ninety-eight elements piled onto one slide rather than each alone.
      const pkg = await Pkg.open(deck);
      const slides = await pkg.slidePaths();
      const target = slides[0];
      if (target === undefined) throw new Error("the library deck has no slides");
      const canvas = await pkg.slideSize();
      const at = choosePlacement(payload.bounds, canvas, []);
      const report = await splice(pkg, target, payload, at);
      const findings = await verify(pkg, target);
      if (findings.length > 0) failures.push(`${entry.id}: ${findings[0]!.kind} — ${findings[0]!.detail}`);
      if (report.shapes !== payload.shapes.length) failures.push(`${entry.id}: only ${report.shapes} shapes spliced`);
      // Nothing in the shipped library should lose a reference. If this ever
      // fires, the harvest and the splice disagree about what travels.
      if (report.droppedRefs.length > 0)
        failures.push(`${entry.id}: dropped ${report.droppedRefs.length} reference(s)`);
    }
    expect(failures).toEqual([]);
  });
});
