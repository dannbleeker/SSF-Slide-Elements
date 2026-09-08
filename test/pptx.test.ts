import { describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import { Pkg } from "../src/core/pptx/pkg.js";
import { P_NS, elements, xmlSafe } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The package layer, on bytes the fixture built.
 *
 * Ported from SSF-Merge's `test/pptx.test.ts` on 2026-09-08: the blocks that
 * need only `Pkg`, the XML helpers and the deck fixture. The blocks that drove
 * SSF-Merge's clone, its tags and its merge run stay there until the splice
 * brings the code they test. The measurements and incidents the comments
 * narrate ("a merge was quadratic", "every merged copy") are SSF-Merge's.
 */
async function deck(...args: Parameters<typeof makeDeck>): Promise<Pkg> {
  return Pkg.open(await makeDeck(...args));
}

const ONE = [{ paragraphs: [["Hello {{Name}}"]], creationId: 111 }];
describe("Pkg", () => {
  it("lists slides in presentation order, not zip order", async () => {
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }]);
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"]);
  });

  it("gives a new relationship the highest id plus one", async () => {
    const pkg = await deck(ONE);
    const first = await pkg.addRel("ppt/presentation.xml", "http://example/t", "x.xml");
    const second = await pkg.addRel("ppt/presentation.xml", "http://example/t", "y.xml");
    expect(Number(first.slice(3))).toBeLessThan(Number(second.slice(3)));
  });

  it("never names a part the package already holds, however big the numbers are", async () => {
    /**
     * "Highest existing number plus one" stops being that above 2^53, where
     * `max + 1 === max` — so a package holding `slide99999999999999999999.xml`
     * answered a number already in use, `copyPart` overwrote it silently, and
     * three merged slides pointed at one part while the deck stayed
     * structurally valid and every check passed.
     *
     * It is the defect `nextNumber`'s own comment says it exists to prevent,
     * reached by a route the comment does not cover. A digit run too large to
     * count exactly is ignored rather than counted, which leaves the maximum
     * exact — and that is also what makes the free-number search terminate.
     */
    const pkg = await deck(ONE);
    // 2^53 exactly, which is the sharp case: `Number` reads it back precisely,
    // `max + 1` rounds straight back to it, and the name that produces is the
    // name already in the package. A longer run of nines is the same defect
    // with a different symptom — `max + 1` answers 1e+20, and the part named
    // after it is nonsense rather than a collision.
    for (const path of [
      "ppt/slides/slide9007199254740992.xml",
      "ppt/charts/chart9007199254740992.xml",
      "ppt/media/image9007199254740992.png",
    ]) {
      pkg.setBytes(path, new Uint8Array([1]));
    }
    expect(pkg.has(`ppt/slides/slide${pkg.nextSlideNumber()}.xml`), "the slide number is already taken").toBe(false);
    expect(pkg.has(`ppt/charts/chart${pkg.nextNumber("ppt/charts/chart")}.xml`)).toBe(false);
    expect(pkg.has(`ppt/media/image${pkg.nextMediaNumber()}.png`)).toBe(false);
    // And the ordinary contract is untouched: the highest real number plus one,
    // never filling a gap.
    expect(pkg.nextSlideNumber()).toBe(2);
  });

  it("refuses to name a part when the package's numbers leave nowhere safe to go", async () => {
    /**
     * The hole the rule above leaves. Ignoring a digit run too large to count
     * keeps the maximum exact — but a package holding the largest COUNTABLE
     * number and the one after it counts the first and ignores the second, so
     * "highest plus one" answers a name that is already there. That is the
     * collision the whole guard exists to prevent, one step further out.
     *
     * There is no larger safe number to hand back, so the answer is to refuse.
     * A deck reaching this has sixteen-digit part numbers and is not a deck.
     */
    const pkg = await deck(ONE);
    pkg.setBytes(`ppt/charts/chart${Number.MAX_SAFE_INTEGER}.xml`, new Uint8Array([1]));
    pkg.setBytes("ppt/charts/chart9007199254740992.xml", new Uint8Array([1]));
    expect(() => pkg.nextNumber("ppt/charts/chart")).toThrow(/too large to extend/);
  });

  it("reads a part's relationships once, not once per relationship added", async () => {
    /**
     * `addRel` re-read every `<Relationship>` in the part to find the highest
     * id, and the presentation's rels is the part that grows by one per merged
     * slide — so a merge was quadratic in the rows. Measured before the fix:
     * 250 rows took 353 ms and 2000 took 7364, eight times the work for
     * twenty-one times the time; after it, 4000 rows take 2705 ms.
     *
     * The assertion is on WORK rather than wall clock, which measures the
     * machine it happens to run on. One walk of the list is all this needs, and
     * the ids must still be distinct and ascending.
     */
    const pkg = await deck(ONE);
    const path = Pkg.relsPathFor("ppt/presentation.xml");
    const doc = await pkg.doc(path);
    let walks = 0;
    const real = doc.getElementsByTagNameNS.bind(doc);
    doc.getElementsByTagNameNS = ((ns: string, local: string) => {
      if (local === "Relationship") walks++;
      return real(ns, local);
    }) as typeof doc.getElementsByTagNameNS;

    const ids: string[] = [];
    for (let i = 0; i < 40; i++) ids.push(await pkg.addRel("ppt/presentation.xml", "http://example/t", `x${i}.xml`));
    expect(walks, "one walk per relationship added is what made a merge quadratic").toBeLessThanOrEqual(1);
    expect(new Set(ids).size, "two relationships were given one id").toBe(ids.length);
    const numbers = ids.map((id) => Number(id.slice(3)));
    expect([...numbers].sort((a, b) => a - b)).toEqual(numbers);
  });

  it("declares a content type once, without re-reading the whole list each time", async () => {
    // `addContentTypeOverride` asked `.some()` over a list it is itself
    // appending to, once per cloned chart or notes page.
    const pkg = await deck(ONE);
    const types = await pkg.doc("[Content_Types].xml");
    let walks = 0;
    const real = types.getElementsByTagNameNS.bind(types);
    types.getElementsByTagNameNS = ((ns: string, local: string) => {
      if (local === "Override") walks++;
      return real(ns, local);
    }) as typeof types.getElementsByTagNameNS;

    for (let i = 0; i < 40; i++) {
      await pkg.addContentTypeOverride(`/ppt/charts/chart${i}.xml`, "application/example");
    }
    expect(walks).toBeLessThanOrEqual(1);
    // Every one declared, and asking again does not add a second.
    await pkg.addContentTypeOverride("/ppt/charts/chart7.xml", "application/example");
    const declared = [...(await pkg.text("[Content_Types].xml")).matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]);
    expect(declared.filter((n) => n === "/ppt/charts/chart7.xml")).toHaveLength(1);
    expect(declared).toContain("/ppt/charts/chart39.xml");
  });

  it("survives a round trip through base64", async () => {
    const pkg = await deck(ONE);
    const again = await Pkg.open(await pkg.toBase64());
    expect(await again.slidePaths()).toEqual(["ppt/slides/slide1.xml"]);
  });
});

describe("where a part's relationships live", () => {
  it("puts them beside the part", () => {
    expect(Pkg.relsPathFor("ppt/slides/slide1.xml")).toBe("ppt/slides/_rels/slide1.xml.rels");
    expect(Pkg.relsPathFor("ppt/presentation.xml")).toBe("ppt/_rels/presentation.xml.rels");
  });

  it("handles a part at the package root", () => {
    // `lastIndexOf("/")` answers -1, and `slice(0, -1)` then drops the part's
    // last CHARACTER: the old answer was
    // `[Content_Types].xm/_rels/[Content_Types].xml.rels`, a plausible-looking
    // path to nowhere. Nothing calls it that way today — this closes it before
    // something does, because the failure is silent.
    expect(Pkg.relsPathFor("[Content_Types].xml")).toBe("_rels/[Content_Types].xml.rels");
  });
});

describe("reading a part the run has edited but not written back", () => {
  it("answers with the edit, not the bytes on disk", async () => {
    /**
     * `Pkg` hands out parsed documents and only serialises them at the end, so
     * a reader that goes straight to the zip sees the version the file was
     * opened with. The changelog records this having bitten once already —
     * "three tests were passing on the version from disk" — and `maybeText` is
     * the reader my chart scan added a caller to.
     */
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["before"]] }]));
    const doc = await pkg.doc("ppt/slides/slide1.xml");
    const t = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t")[0];
    expect(t, "the fixture changed shape").toBeTruthy();
    t!.textContent = "after";
    expect(await pkg.maybeText("ppt/slides/slide1.xml")).toContain("after");
    expect(await pkg.maybeText("ppt/slides/slide1.xml")).not.toContain("before");
  });

  it("answers nothing for a part that is not there", async () => {
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["x"]] }]));
    expect(await pkg.maybeText("ppt/charts/chart1.xml")).toBeUndefined();
  });
});

describe("relating a part that had no relationships at all", () => {
  it("creates the rels file rather than failing", async () => {
    // `addRel` writes an empty `<Relationships>` when the part has none. Every
    // caller so far happened to work on a part that already had a rels file,
    // so the branch that creates one had never run.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["x"]] }]));
    const owner = "ppt/theme/theme1.xml";
    expect(pkg.has(Pkg.relsPathFor(owner)), "the fixture already gave theme1 a rels part").toBe(false);
    const rId = await pkg.addRel(owner, "http://example.invalid/rel", "../media/image1.png");
    expect(rId).toBe("rId1");
    expect(await pkg.relTarget(owner, rId)).toBe("ppt/media/image1.png");
  });
});

describe("a template slide someone has commented on", () => {
  /**
   * A comment hangs off the SLIDE, so the wholesale rels copy a clone starts
   * from hands every copy a relationship to the TEMPLATE's comment part. Three
   * slides, one `modernComment_101_AEAB9DA1.xml` — measured before the fix.
   *
   * That puts a reviewer's "check this with Legal" on all 240 merged slides, as
   * one shared thread. Copying the part per clone would be worse rather than
   * better: the same note 240 times, deliberately. A comment is an annotation
   * about the template, not content the template produces.
   *
   * And dropping them is what makes the two template routes AGREE. On a 1.10
   * host `exportAsBase64Presentation` drops comments and `ppt/authors.xml`
   * outright (office-js#6867 — measured on this host 2026-08-28, four comment
   * parts in and none out), so the subset route already produced comment-free
   * clones while the file route produced shared ones. Two routes, two different
   * decks, from one template.
   */
  const MODERN = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
  it("finds a part whose name is percent-encoded, the way OPC stores it", async () => {
    /**
     * OPC maps a part name to a ZIP item name by stripping the leading `/` and
     * nothing else, so `ppt/charts/my%20chart.xml` is stored under exactly that
     * name. The resolver percent-DECODED every segment, answered
     * `ppt/charts/my chart.xml`, and `has` said no — and `cloneSlideGraphics`
     * reads "not in the package" as "skip", so every merged copy silently went
     * on pointing at the template's own chart.
     *
     * Both spellings are produced now and the package decides. A writer whose
     * zip entries carry the decoded name still resolves: that is the second
     * assertion, and it is why decoding could not simply be removed.
     */
    const bytes = await makeDeck([{ paragraphs: [["Hello {{First}}"]] }]);
    const zip = await JSZip.loadAsync(bytes);
    zip.file("ppt/charts/my%20chart.xml", '<?xml version="1.0"?><c/>');
    zip.file("ppt/charts/other chart.xml", '<?xml version="1.0"?><c/>');
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/my%20chart.xml"/>` +
          `<Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/other%20chart.xml"/></Relationships>`,
      ),
    );
    const pkg = await Pkg.open(await zip.generateAsync({ type: "uint8array" }));

    const encoded = await pkg.relTarget("ppt/slides/slide1.xml", "rIdA");
    expect(encoded, "the spelling the package actually holds").toBe("ppt/charts/my%20chart.xml");
    expect(pkg.has(encoded ?? "")).toBe(true);

    // The other direction: a package whose entry carries the decoded name.
    const decoded = await pkg.relTarget("ppt/slides/slide1.xml", "rIdB");
    expect(decoded).toBe("ppt/charts/other chart.xml");
    expect(pkg.has(decoded ?? "")).toBe(true);
  });

  async function deckWithComment(relType: string, part: string): Promise<Pkg> {
    const bytes = await makeDeck([{ paragraphs: [["Hello {{First}}"]] }, { paragraphs: [["after"]] }]);
    const zip = await JSZip.loadAsync(bytes);
    zip.file(part, '<?xml version="1.0"?><cm/>');
    const rels = await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="rIdC" Type="${relType}" Target="../${part.replace("ppt/", "")}"/></Relationships>`,
      ),
    );
    return Pkg.open(await zip.generateAsync({ type: "uint8array" }));
  }

  it("takes the comment part away with the slide it belonged to", async () => {
    /**
     * The other half. A comment belongs to ONE slide and is unreachable once
     * that slide is gone — and now that a clone no longer references the
     * template's part, removing the template on the way out would strand it: a
     * part with a content-type override and nothing pointing at it.
     */
    const part = "ppt/comments/modernComment_101_AEAB9DA1.xml";
    const pkg = await deckWithComment(MODERN, part);
    expect(pkg.has(part)).toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has(part), "the comment part outlived its slide").toBe(false);
  });
});

describe("a part that arrives with a byte order mark", () => {
  /**
   * OPC permits a BOM on an XML part and .NET's default `UTF8Encoding` emits
   * one, so any deck from a third-party generator built on it carries one on
   * every part that generator wrote. PowerPoint opens such a deck without a
   * murmur.
   *
   * `@xmldom/xmldom` does not: a leading `U+FEFF` puts the XML declaration at
   * position 1 and it throws `processing instruction at position 1 is an xml
   * declaration which is only at the start of the document`. JSZip's
   * `async("string")` hands the character through — it decodes UTF-8 and has
   * no opinion about what the first code point means — so the mark reaches the
   * parser as content and the merge dies on the first slide it reads, naming
   * neither the part nor the reason a user could act on.
   */
  async function deckWithBom(part: string): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await makeDeck([{ paragraphs: [["Hello {{Name}}"]] }]));
    const file = zip.file(part);
    if (!file) throw new Error(`the fixture has no ${part}, so this test proves nothing`);
    zip.file(part, `\uFEFF${await file.async("string")}`);
    return zip.generateAsync({ type: "uint8array" });
  }

  it("reads a slide whose markup starts with one", async () => {
    const pkg = await Pkg.open(await deckWithBom("ppt/slides/slide1.xml"));
    // The mark really is still in the bytes — if JSZip ever started stripping
    // it, this test would pass while proving nothing.
    expect((await pkg.text("ppt/slides/slide1.xml")).charCodeAt(0)).toBe(0xfeff);
    expect(elements(await pkg.doc("ppt/slides/slide1.xml"), P_NS, "cSld")).toHaveLength(1);
  });
});

describe("sweeping the template block off a long deck", () => {
  /**
   * `removeSlide` used to resolve every `<p:sldId>` in the deck through
   * `relTarget`, which re-walks the presentation's whole relationship list for
   * one id — so a single removal cost `deck x deck`, and `src/office/merge.ts`
   * removes the template block one slide at a time. Measured on a 100-slide
   * user deck plus 400 clones: 4.5 SECONDS of blocking work in a task-pane
   * WebView, after the merge had already finished and with nothing on screen to
   * say why.
   *
   * Asserted as work rather than as wall clock, because a timing threshold on a
   * shared runner is a test that fails for reasons nobody can act on. The
   * property is that resolving the id list is not a per-slide question: the
   * relationships are read once for the whole sweep. Pre-fix this counts one
   * `relTarget` call per `<p:sldId>` per removal — 4 950 of them for the deck
   * below — and the number is what the seconds were made of.
   */
  it("does not re-resolve the whole slide id list once per slide", async () => {
    const pkg = await Pkg.open(await makeDeck(Array.from({ length: 100 }, () => ({ paragraphs: [["a"]] }))));
    const paths = await pkg.slidePaths();
    const seen = vi.spyOn(pkg, "relTarget");
    for (const path of paths.slice(0, 50)) await pkg.removeSlide(path);
    expect(await pkg.slidePaths()).toHaveLength(50);
    expect(seen.mock.calls.length, "the id list is being re-resolved per slide").toBeLessThanOrEqual(paths.length);
    seen.mockRestore();
  });
});

describe("text carrying a character XML cannot hold", () => {
  /**
   * XML 1.0 forbids most C0 controls, the lone surrogates and `FFFE`/`FFFF`
   * outright, with no escape: `&#11;` is exactly as ill-formed as the byte.
   * `@xmldom/xmldom` writes such a character straight through and reads it
   * back, so a part holding one passes every gate here and PowerPoint calls
   * the whole file damaged. `xmlSafe` is the one rule, in `pptx/xml.ts`, that
   * every writer of text into the package shares; SSF-Merge learned it once
   * for slide text and once for tag values before putting it in one place.
   */
  const FORBIDDEN = (text: string) =>
    [...text].filter((c) => {
      const n = c.charCodeAt(0);
      return (n < 0x20 && n !== 9 && n !== 10 && n !== 13) || n === 0xfffe || n === 0xffff;
    });

  it("replaces what a conforming parser would refuse with a space", () => {
    for (const [what, value] of [
      ["a NUL", "a\u0000b"],
      ["a vertical tab", "a\u000bb"],
      ["a form feed", "a\u000cb"],
      ["U+FFFE", "p￾q"],
    ] as const) {
      const safe = xmlSafe(value);
      expect(FORBIDDEN(safe), what).toEqual([]);
      // A space rather than nothing: the likeliest such character is a line
      // break, and dropping it would join two words.
      expect(safe, what).toHaveLength(value.length);
      expect(safe, what).toContain(" ");
    }
  });

  it("keeps a lone surrogate out, and an astral character in", () => {
    // An unpaired half is already broken text; a well-formed pair is one
    // ordinary code point above FFFF and must survive, which is the
    // distinction `\p{Surrogate}` under the `u` flag draws.
    expect(xmlSafe("x\ud800y")).not.toContain("\ud800");
    expect(xmlSafe(`ok ${String.fromCodePoint(0x1f600)}`)).toContain(String.fromCodePoint(0x1f600));
  });

  it("leaves the whitespace XML allows alone", () => {
    expect(xmlSafe("a\tb\nc\rd")).toBe("a\tb\nc\rd");
  });
});
