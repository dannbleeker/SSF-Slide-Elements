/**
 * Whether each committed print still belongs to the deck beside it.
 *
 * `docs/DESIGN.md` section 3 cuts every element's preview out of these prints.
 * The design's own check — the print's page count against the deck's slide
 * count — cannot tell a current print from a stale one, because the count does
 * not change when the content does. The rules in
 * `scripts/print-provenance.mjs` name the bytes instead.
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types. The rules live THERE so
// the stamping script and this file cannot read different ones.
import { pdfPageCount, printProblems, sha256, slideCountOf } from "../scripts/print-provenance.mjs";

const DECKS = ["library-16x9", "library-4x3"] as const;

type Stamp = {
  deck: string;
  deckSha256: string;
  slides: number;
  print: string;
  printSha256: string;
  pages: number;
  takenAt: string;
  powerPoint: string;
  how: string;
};

const read = (name: string) => ({
  deckBytes: readFileSync(`template/${name}.pptx`),
  printBytes: readFileSync(`template/${name}.pdf`),
  stamp: JSON.parse(readFileSync(`template/${name}.print.json`, "utf8")) as Stamp,
});

async function slidesIn(deckBytes: Buffer): Promise<number> {
  const zip = await JSZip.loadAsync(deckBytes);
  return slideCountOf(await zip.file("ppt/presentation.xml")!.async("string")) as number;
}

describe("the committed prints", () => {
  for (const name of DECKS) {
    it(`${name}: the print is of the deck committed beside it`, async () => {
      const { deckBytes, printBytes, stamp } = read(name);
      const slides = await slidesIn(deckBytes);
      expect(printProblems({ name, stamp, deckBytes, printBytes, slides })).toEqual([]);
    });
  }

  it("says which of the two routes took it, in the words the stamping script writes", async () => {
    // `how` is the sidecar's provenance, and a sentence nobody checks is a
    // sentence that drifts: a sidecar carrying the Export dialog's settings for
    // a print that never saw the dialog would be the one field here that lies.
    // Held to the constants rather than to a keyword, so the two spellings
    // cannot diverge.
    //
    // BOTH decks now say COM. They did not always: the 16:9 print came off the
    // dialog and the 4:3 one through COM, and that difference is what made this
    // field worth checking in the first place. Both were re-printed through COM
    // on 2026-09-16, when the Flowchart shapes slide was deleted from each — so
    // the assertion that they DIFFER has gone, and what is left is that each
    // names a route the stamping script actually writes.
    // @ts-expect-error — plain .mjs with no types.
    const { HOW } = (await import("../scripts/print-provenance.mjs")) as { HOW: Record<string, string> };
    const routes = Object.values(HOW);
    expect(routes.length).toBeGreaterThan(1);
    for (const name of DECKS) {
      expect(routes, `${name}: how`).toContain(read(name).stamp.how);
    }
    // Each names a route the script writes, whichever route that is.
    expect(routes).toContain(read("library-16x9").stamp.how);
    expect(routes).toContain(read("library-4x3").stamp.how);
  });

  it("reads a page count two independent ways, or refuses to answer", () => {
    const { printBytes } = read("library-16x9");
    expect(pdfPageCount(printBytes)).toBe(109);
    expect(pdfPageCount(read("library-4x3").printBytes)).toBe(107);
    // Something that is not a PDF has no page count, and does not get a guess.
    expect(pdfPageCount(Buffer.from("%PDF-1.7\nnot really\n"))).toBeNull();
  });

  describe("the rules go red when they should", () => {
    // A gate that cannot fail is not a gate (CLAUDE.md). Each case below moves
    // ONE thing and names the sentence it expects back.
    const base = () => {
      const { deckBytes, printBytes, stamp } = read("library-16x9");
      return { name: "library-16x9", deckBytes, printBytes, stamp: { ...stamp }, slides: 110 };
    };

    it("catches a deck edited after the print was taken", () => {
      const args = base();
      args.deckBytes = Buffer.concat([args.deckBytes, Buffer.from([0])]);
      expect(printProblems(args).join(" ")).toContain("the deck has changed since the print was taken");
    });

    it("catches a print replaced without re-stamping", () => {
      const args = base();
      args.stamp.printSha256 = sha256(Buffer.from("a different print"));
      expect(printProblems(args).join(" ")).toContain("the print has changed since it was stamped");
    });

    it("catches a deck that gained a slide", () => {
      const args = base();
      args.slides = 110;
      const said = printProblems(args).join(" ");
      expect(said).toContain("sidecar says 109 slides, the deck has 110");
      expect(said).toContain("109 pages for 110 slides");
    });

    it("catches a missing sidecar", () => {
      const args = base();
      expect(printProblems({ ...args, stamp: undefined }).join(" ")).toContain("no print sidecar");
    });

    it("catches a sidecar that does not say how the print was taken", () => {
      const args = base();
      delete (args.stamp as Partial<Stamp>).how;
      expect(printProblems(args).join(" ")).toContain("does not say how");
    });
  });
});
