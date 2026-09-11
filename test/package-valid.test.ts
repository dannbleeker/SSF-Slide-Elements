import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { resolveTargetSpellings } from "../src/core/pptx/pkg.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * What the engine HANDS OVER, checked as a package rather than as slides.
 *
 * Ported from SSF-Merge's `test/package-valid.test.ts` on 2026-09-08: the
 * oracle and the two tests that prove it is not vacuous. The splice's own
 * sweep over all 117 elements went to `scripts/package-integrity.mjs` instead,
 * because it runs outside the suite as well; `problemsIn` stays here as the
 * file-local oracle those two tests are about.
 *
 * The question PowerPoint asks is whether the file it is given is a legal OOXML
 * package, and the answer is binary and expensive: a deck that opens as
 * "repaired" has lost whatever PowerPoint decided to drop, silently, in
 * somebody's presentation. The rules below are the package format's, not this
 * project's, so they hold however the engine changes. SSF-Merge shipped the
 * shape of bug they catch once — the user's whole deck sent back because the
 * remove-the-rest path was wrong — and a structural check over the output is
 * what would have caught it without a PowerPoint.
 */
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const parse = (text: string): Document => new DOMParser().parseFromString(text, "text/xml") as unknown as Document;
const els = (d: Document, ns: string, n: string): Element[] => Array.from(d.getElementsByTagNameNS(ns, n));

/** Every way the package could be malformed, as a list of sentences. */
async function problemsIn(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n]?.dir);
  const has = (p: string): boolean => names.includes(p);
  /** Whichever spelling of a resolved target this package holds — see `Pkg.resolved`. */
  const held = (asWritten: string, decoded: string): string =>
    asWritten !== decoded && has(asWritten) ? asWritten : decoded;
  const read = async (p: string): Promise<Document> => parse(await zip.file(p)!.async("string"));
  const problems: string[] = [];

  // A relationship pointing at nothing is how a deck opens as "repaired".
  for (const name of names.filter((n) => n.includes("/_rels/") && n.endsWith(".rels"))) {
    const owner = name.replace("/_rels/", "/").replace(/\.rels$/, "");
    const doc = await read(name);
    const rels = els(doc, PKG_REL, "Relationship");
    for (const rel of rels) {
      // An external target is a URL and is not a package path.
      if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
      const target = rel.getAttribute("Target") ?? "";
      // BOTH spellings, like the engine. OPC maps a part name to a ZIP item by
      // stripping the leading `/` and nothing else, so a percent-encoded name is
      // stored verbatim — an oracle that only decoded would report every
      // legitimately encoded package as broken.
      const path = held(...resolveTargetSpellings(owner, target));
      if (!has(path))
        problems.push(`${name}: ${rel.getAttribute("Id")} points at ${target}, which is not in the package`);
    }
    const ids = rels.map((r) => r.getAttribute("Id"));
    const dupes = [...new Set(ids.filter((v, i) => ids.indexOf(v) !== i))];
    if (dupes.length) problems.push(`${name}: duplicate rIds ${dupes.join(", ")}`);
  }

  // A part with no content type opens as damaged, and PowerPoint does not say
  // which part it could not classify.
  //
  // The generic rule alone is TOOTHLESS for the parts that matter, and this was
  // measured rather than assumed: a real .pptx declares
  // `<Default Extension="xml" ContentType="application/xml"/>`, so every XML
  // part passes it. Deleting the clone's `addContentTypeOverride` — a defect
  // that would ship every merged slide untyped — left this check green. A slide
  // needs its OWN override naming the slide content type, so those are checked
  // by name.
  const ct = await read("[Content_Types].xml");
  const declared = new Map(
    els(ct, CT_NS, "Override").map((o) => [o.getAttribute("PartName") ?? "", o.getAttribute("ContentType") ?? ""]),
  );
  const defaults = new Set(els(ct, CT_NS, "Default").map((d) => (d.getAttribute("Extension") ?? "").toLowerCase()));
  for (const name of names) {
    if (name === "[Content_Types].xml") continue;
    const ext = (name.split(".").pop() ?? "").toLowerCase();
    if (!declared.has(`/${name}`) && !defaults.has(ext)) problems.push(`no content type declared for ${name}`);
  }
  for (const [part] of declared)
    if (part && !has(part.replace(/^\//, ""))) problems.push(`content type declared for ${part}, which is gone`);

  const NEEDS: [RegExp, string][] = [
    [/^ppt\/slides\/slide\d+\.xml$/, "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"],
    [
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/,
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    ],
  ];
  for (const name of names) {
    for (const [pattern, type] of NEEDS) {
      if (!pattern.test(name)) continue;
      const got = declared.get(`/${name}`);
      if (got === undefined) problems.push(`${name} has no content-type override of its own`);
      else if (got !== type) problems.push(`${name} is declared as ${got}, not ${type}`);
    }
  }

  // The deck's own order, which is what decides the slides a reader sees.
  const pres = await read("ppt/presentation.xml");
  const presRels = await read("ppt/_rels/presentation.xml.rels");
  const target = new Map(
    els(presRels, PKG_REL, "Relationship").map((r) => [r.getAttribute("Id"), r.getAttribute("Target")]),
  );
  const seen = new Set<string>();
  const listed = new Set<string>();
  for (const sldId of els(pres, P_NS, "sldId")) {
    const id = sldId.getAttribute("id") ?? "";
    if (seen.has(id)) problems.push(`two slides share the id ${id}`);
    seen.add(id);
    // The format's range. Outside it PowerPoint rejects the file outright.
    const n = Number(id);
    if (!(n >= 256 && n <= 2147483647)) problems.push(`slide id ${id} is outside the format's range`);
    const rId = sldId.getAttributeNS(R_NS, "id") ?? sldId.getAttribute("r:id");
    const to = rId === null ? undefined : target.get(rId);
    if (to === undefined) problems.push(`the deck lists a slide whose relationship ${rId ?? "(none)"} does not exist`);
    else {
      const path = held(...resolveTargetSpellings("ppt/presentation.xml", to ?? ""));
      listed.add(path);
      if (!has(path)) problems.push(`the deck lists ${path}, which is not in the package`);
    }
  }
  // A slide part nothing lists is dead weight the user still carries.
  for (const name of names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    if (!listed.has(name)) problems.push(`${name} is in the package but not in the deck`);
  }
  return problems;
}

describe("the package oracle", () => {
  it("has something to say about a package that is broken", async () => {
    /**
     * The check that the check is not vacuous. Every assertion above is
     * `toEqual([])`, which is what an empty list of RULES also produces — and
     * this suite has twice caught a gate that measured nothing and reported
     * success. So a deliberately damaged package must come back with each kind
     * of complaint.
     */
    const deck = await makeDeck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    const zip = await JSZip.loadAsync(deck);
    // Remove a slide part while leaving the deck listing it, and strip its
    // content type. One edit, three rules.
    zip.remove("ppt/slides/slide2.xml");
    const found = await problemsIn(await zip.generateAsync({ type: "uint8array" }));
    expect(found.some((p) => p.includes("not in the package"))).toBe(true);
    expect(found.some((p) => p.includes("content type declared for"))).toBe(true);

    // And the rule the generic one cannot reach: a slide typed only by the
    // package's `Default Extension="xml"`.
    const untyped = await JSZip.loadAsync(await makeDeck([{ paragraphs: [["a"]] }]));
    const types = await untyped.file("[Content_Types].xml")!.async("string");
    untyped.file("[Content_Types].xml", types.replace(/<Override PartName="\/ppt\/slides\/slide1\.xml"[^>]*\/>/, ""));
    const second = await problemsIn(await untyped.generateAsync({ type: "uint8array" }));
    expect(second).toContain("ppt/slides/slide1.xml has no content-type override of its own");
  });

  it("accepts a percent-encoded part name, and still catches one that is missing", async () => {
    /**
     * OPC maps a part name to a ZIP item by stripping the leading `/` and
     * nothing else, so `notes%20slide1.xml` is stored under exactly that name.
     * This oracle decoded every target, so it would have called a legitimately
     * encoded package broken — and, being the check the engine's own fix was
     * measured against, it could not have caught the engine getting it wrong.
     *
     * Both halves are asserted, because an oracle that accepts everything is
     * the vacuous measurement this file exists to refuse.
     */
    const deck = await makeDeck([{ paragraphs: [["a"]], notes: "note" }, { paragraphs: [["b"]] }]);
    const zip = await JSZip.loadAsync(deck);
    const notes = await zip.file("ppt/notesSlides/notesSlide1.xml")!.async("uint8array");
    const rels = await zip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels")!.async("uint8array");
    zip.remove("ppt/notesSlides/notesSlide1.xml");
    zip.remove("ppt/notesSlides/_rels/notesSlide1.xml.rels");
    zip.file("ppt/notesSlides/notes%20slide1.xml", notes);
    zip.file("ppt/notesSlides/_rels/notes%20slide1.xml.rels", rels);
    for (const path of ["ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml"]) {
      const text = await zip.file(path)!.async("string");
      zip.file(path, text.replaceAll("notesSlide1.xml", "notes%20slide1.xml"));
    }
    expect(await problemsIn(await zip.generateAsync({ type: "uint8array" }))).toEqual([]);

    // And the same package with that part taken out is still reported.
    zip.remove("ppt/notesSlides/notes%20slide1.xml");
    const broken = await problemsIn(await zip.generateAsync({ type: "uint8array" }));
    expect(broken.some((p) => p.includes("not in the package"))).toBe(true);
  });
});
