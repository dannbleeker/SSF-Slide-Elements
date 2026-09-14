import { describe, expect, it, vi } from "vitest";
import JSZip from "jszip";
import { Pkg, extensionOf } from "../src/core/pptx/pkg.js";
import { REL_TYPE } from "../src/core/pptx/parts.js";
import { P_NS, element, elements, xmlSafe } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

/**
 * The package layer, on bytes the fixture built.
 *
 * Ported from SSF-Merge's `test/pptx.test.ts` on 2026-09-08: the blocks that
 * need only `Pkg`, the XML helpers and the deck fixture. The splice brought
 * the clone and the tag writer on 2026-09-10, and their blocks came with them:
 * they are in `pptx-clone.test.ts` and `pptx-tags.test.ts`, not here. The
 * merge run's blocks stay in SSF-Merge, where the merge is. The measurements and incidents the comments
 * narrate ("a merge was quadratic", "every merged copy") are SSF-Merge's.
 */
async function deck(...args: Parameters<typeof makeDeck>): Promise<Pkg> {
  return Pkg.open(await makeDeck(...args));
}

const ONE = [{ paragraphs: [["Hello {{Name}}"]], creationId: 111 }];
const TWO = [{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }];
const PRESENTATION = "ppt/presentation.xml";
const SLIDE1 = "ppt/slides/slide1.xml";
const SLIDE1_RELS = "ppt/slides/_rels/slide1.xml.rels";
const CONTENT_TYPES = "[Content_Types].xml";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/** A `.rels` part holding exactly these entries. */
function rels(...entries: string[]): string {
  return `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">${entries.join("")}</Relationships>`;
}

/** A `[Content_Types].xml` holding exactly these entries. */
function types(...entries: string[]): string {
  return `<?xml version="1.0"?><Types xmlns="${CT_NS}">${entries.join("")}</Types>`;
}

/**
 * A real fixture deck with named parts replaced, or removed with `null`.
 *
 * Rebuilt as a zip and reopened rather than edited through `setText`, so a part
 * genuinely ABSENT is reachable — which `Pkg` deliberately offers no way to do
 * from the outside. The same helper, for the same reason, is in
 * `pptx-malformed.test.ts`; these cases are about the package layer's own
 * counters and sweeps rather than about reading somebody else's XML, so they
 * live beside the rest of `Pkg`.
 */
async function craft(edits: Record<string, string | null>): Promise<Pkg> {
  const source = await JSZip.loadAsync(await makeDeck(TWO));
  const out = new JSZip();
  for (const name of Object.keys(source.files).filter((n) => !source.files[n]?.dir)) {
    if (name in edits) continue;
    out.file(name, await (source.file(name) as JSZip.JSZipObject).async("uint8array"));
  }
  for (const [name, body] of Object.entries(edits)) if (body !== null) out.file(name, body);
  return Pkg.open(await out.generateAsync({ type: "uint8array" }));
}
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

  it("keeps an absolute part name absolute", () => {
    // The other spelling a part name comes in. A content-type Override names
    // its part the way ECMA-376 does, with the leading slash
    // (`addContentTypeOverride` refuses anything else), and the answer for one
    // has to stay in that spelling rather than turn into a third thing. The
    // root case below is the ZIP spelling of the same question, and the two
    // branches of this function are exactly those two spellings: the `< 0`
    // arm is for a ZIP name with no directory at all, and everything else —
    // absolute names included — goes through the slice.
    expect(Pkg.relsPathFor("/ppt/slides/slide1.xml")).toBe("/ppt/slides/_rels/slide1.xml.rels");
    expect(Pkg.relsPathFor("/[Content_Types].xml")).toBe("/_rels/[Content_Types].xml.rels");
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

describe("a part whose file name begins with a dot", () => {
  it("reads the extension of a dotfile, which is the whole name after the dot", () => {
    /**
     * `_rels/.rels` is in every package there is, and its content type comes
     * from the `<Default Extension="rels">` every package declares — so "what
     * is the extension of `.rels`" is a question `contentTypeOf` asks about a
     * real part, and the answer is `rels`.
     *
     * Two ways to get it wrong, and both answer the empty string, which reads
     * as "nothing covers this part": taking the segment from one character
     * past the slash drops the dot, and treating a dot at position 0 as no dot
     * at all refuses the name outright.
     */
    expect(extensionOf("_rels/.rels")).toBe("rels");
    expect(extensionOf(".rels")).toBe("rels");
    expect(extensionOf("ppt/slides/.hidden")).toBe("hidden");
    // And the neighbouring case stays where it was: a dot that is the whole
    // name has nothing after it.
    expect(extensionOf(".")).toBe("");
  });
});

describe("counters that must not drift", () => {
  it("does not lower a part number when the package loses a part", async () => {
    /**
     * The counters are a HIGH-WATER MARK for this `Pkg`, not a fact about the
     * package: `usedNumbers` scans the zip once and `noteWritten` keeps the set
     * current, but `removePart` deliberately never takes a number back out. A
     * re-scan instead of the memo would answer 3 here — the number of the slide
     * that has just been deleted — and the next write would land on a part
     * another slide may still relate to.
     *
     * The memo has to be PRIMED before the removal for this to say anything:
     * asked for the first time afterwards, both readings agree.
     */
    const pkg = await deck([{ paragraphs: [["a"]] }, { paragraphs: [["b"]] }, { paragraphs: [["c"]] }]);
    expect(pkg.nextSlideNumber(), "the counter is read here, before the deck changes").toBe(4);
    await pkg.removeSlide("ppt/slides/slide3.xml");
    expect(pkg.nextSlideNumber(), "a removal lowered the counter, so the next write reuses a name").toBe(4);
  });

  it("does not let a part of one family advance another family's counter", async () => {
    // Every counter is filled by the SAME scan of every path in the package,
    // and a path that does not match the family's pattern has to contribute
    // nothing to it. Contributing a number instead — any number — costs the
    // family its first name for the life of the `Pkg`, and the part it skips
    // is one nothing in the package has ever held.
    const pkg = await deck(ONE);
    expect(pkg.nextNumber("ppt/charts/chart"), "the deck holds no chart, so the first is chart1").toBe(1);
    pkg.setText("ppt/notesSlides/notesSlide7.xml", "<notes/>");
    expect(pkg.nextNumber("ppt/charts/chart"), "a notes page advanced the chart counter").toBe(1);
  });
});

describe("the slide id list at its edges", () => {
  it("gives the first slide in an empty list PowerPoint's own starting id", async () => {
    // `<p:sldIdLst>` with nothing in it is what a deck with no slides has, and
    // it is the one case where the floor decides the answer rather than the
    // ids already there. The format reserves everything below 256, so the
    // first id handed out is 256 exactly — one lower is a value PowerPoint
    // will not take, one higher is a number wasted for no reason.
    const pkg = await deck(ONE);
    const list = element(await pkg.doc(PRESENTATION), P_NS, "sldIdLst");
    if (!list) throw new Error("the fixture changed shape");
    for (const sldId of elements(list, P_NS, "sldId")) list.removeChild(sldId);
    expect(await pkg.appendSldId("rId2")).toBe(256);
    expect(await pkg.appendSldId("rId2")).toBe(257);
  });

  it("uses the highest id the format allows rather than refusing it", async () => {
    // The range is closed at the top: 2147483647 is the largest value
    // `<p:sldId id="…">` takes, so a deck whose highest is one below still has
    // exactly one id left. Refusing at the boundary instead would lose a slide
    // the format has room for, and the refusal reads as a deck problem.
    const pkg = await deck(ONE);
    const sldId = element(await pkg.doc(PRESENTATION), P_NS, "sldId");
    if (!sldId) throw new Error("the fixture changed shape");
    sldId.setAttribute("id", "2147483646");
    expect(await pkg.appendSldId("rId2")).toBe(2147483647);
    // And the one after it is the one there is no room for.
    await expect(pkg.appendSldId("rId3")).rejects.toThrow(/run out of slide ids/);
  });
});

describe("numbering relationships in a part whose ids are not rIdN", () => {
  it("starts at rId1 and goes up one at a time", async () => {
    /**
     * `addRel` answers "the highest `rIdN` in the part, plus one", and a part
     * whose entries are not in that family has no highest — so the first id it
     * hands out is `rId1`. An unreadable id that contributes a number instead
     * skips `rId1` for good, and a high-water mark left one past the id just
     * issued puts a gap between every pair.
     *
     * Neither is a collision, which is why they are invisible until counted:
     * the ids stay unique and ascending either way. `pptx-malformed.test.ts`
     * pins the same function against a part that DOES carry an `rIdN`.
     */
    const pkg = await craft({
      [SLIDE1_RELS]: rels(
        `<Relationship Type="${REL_TYPE.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>`,
        `<Relationship Id="docRel" Type="http://example/t" Target="a.xml"/>`,
      ),
    });
    expect(await pkg.addRel(SLIDE1, "http://example/t", "b.xml"), "an unreadable id consumed a number").toBe("rId1");
    expect(await pkg.addRel(SLIDE1, "http://example/t", "c.xml"), "the ids have a gap in them").toBe("rId2");
  });
});

describe("a relationship that names no target at all", () => {
  it("resolves the rest of the slide list past it rather than raising", async () => {
    /**
     * `relTargets` walks a `.rels` and resolves every entry in it. An entry
     * with an `Id` and no `Target` is the shape that matters: it looks usable
     * up to the point where the target is read, and the resolver is handed the
     * DOM's null. Reading the deck is the very first thing an insert does, so
     * a raise here is a deck the pane cannot open at all.
     */
    const pkg = await craft({
      "ppt/_rels/presentation.xml.rels": rels(
        `<Relationship Id="rId1" Type="${REL_TYPE.slideMaster}" Target="slideMasters/slideMaster1.xml"/>`,
        `<Relationship Id="rId2" Type="${REL_TYPE.slide}"/>`,
        `<Relationship Id="rId3" Type="${REL_TYPE.slide}" Target="slides/slide2.xml"/>`,
      ),
    });
    // `rId2` is slide 1 in the fixture and now names nothing, so the list is
    // the one slide that still resolves.
    expect(await pkg.slidePaths()).toEqual(["ppt/slides/slide2.xml"]);
    await pkg.removeSlide("ppt/slides/slide2.xml");
    expect(await pkg.slidePaths()).toEqual([]);
  });
});

describe("what the orphan sweep is allowed to destroy", () => {
  it("leaves a part the package never held, and everything named after it, alone", async () => {
    /**
     * A chart owns its embedded workbook, so the sweep follows one hop out of
     * the chart and takes what it finds. A DANGLING target — a workbook the
     * package does not hold — must not reach that list, and "it is not there,
     * so removing it does nothing" is not the reason: `removePart` also
     * deletes the target's own `.rels` and its content-type declaration, and
     * neither of those is checked against the part existing.
     *
     * So a package that kept a workbook's relationships after losing the
     * workbook loses those too, silently, on a removal that had nothing to do
     * with them. Measured on 2026-09-14 — `pptx-malformed.test.ts` calls the
     * same check belt and braces, which is true of the deck it uses and not of
     * this one.
     */
    const pkg = await craft({
      [SLIDE1_RELS]: rels(`<Relationship Id="rId1" Type="${REL_TYPE.chart}" Target="../charts/chart1.xml"/>`),
      "ppt/charts/chart1.xml": '<?xml version="1.0"?><c/>',
      "ppt/charts/_rels/chart1.xml.rels": rels(
        `<Relationship Id="rId1" Type="${REL_TYPE.package}" Target="../embeddings/wb.xlsx"/>`,
      ),
      "ppt/embeddings/_rels/wb.xlsx.rels": rels(
        `<Relationship Id="rId1" Type="${REL_TYPE.image}" Target="../media/nothing.png"/>`,
      ),
    });
    await pkg.addContentTypeOverride("/ppt/embeddings/wb.xlsx", "application/x-absent");
    await pkg.removeSlide(SLIDE1);
    // The chart is the slide's own and goes, which is what makes the sweep run.
    expect(pkg.has("ppt/charts/chart1.xml")).toBe(false);
    expect(
      pkg.has("ppt/embeddings/_rels/wb.xlsx.rels"),
      "the sweep took the relationships of a part it never found",
    ).toBe(true);
    expect(await pkg.text(CONTENT_TYPES)).toContain("/ppt/embeddings/wb.xlsx");
  });

  it("does not read every relationships part in the deck for a slide that owns nothing", async () => {
    /**
     * The referrer scan reads every `.rels` in the package to decide whether a
     * chart, a tag or a picture is still spoken for. A slide that owns none of
     * those has nothing to decide, and most slides own none — so the scan is
     * skipped rather than run and thrown away.
     *
     * Asserted as WORK, like the id-list sweep above, because the alternative
     * is a clock on a shared runner. It is the same shape of defect the rest of
     * this file records: the sweep runs once per slide removed, so scanning the
     * deck inside it costs `removed x deck`.
     */
    const pkg = await Pkg.open(await makeDeck(Array.from({ length: 12 }, () => ({ paragraphs: [["a"]] }))));
    const seen = vi.spyOn(pkg, "relatedParts");
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(
      seen.mock.calls.length,
      "every .rels in the deck is being read for a slide that owns nothing",
    ).toBeLessThanOrEqual(2);
    seen.mockRestore();
  });

  it("does not end up holding the whole deck's relationships either", async () => {
    // The other half of the same skip, and its own case because the first
    // assertion of a case is the only one a failure reaches. Every part the
    // referrer scan reads is PARSED, and a parsed part stays parsed — which is
    // the cost `release` and `peek` exist to keep off a task-pane WebView.
    const pkg = await Pkg.open(await makeDeck(Array.from({ length: 12 }, () => ({ paragraphs: [["a"]] }))));
    const held = pkg.cachedParts();
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.cachedParts() - held, "the sweep parsed and kept the whole deck's relationships").toBeLessThanOrEqual(3);
  });
});

describe("content types that nothing really covers", () => {
  it("does not cover a part that has no extension with a Default that names none", async () => {
    /**
     * `ppt/embeddings/workbook` is a real shape — an OLE embedding whose
     * producer gave it no extension — and a `<Default>` with no `Extension`
     * attribute is another, from a writer that left it out. Reading the second
     * as the answer for the first declares a part as something nobody said it
     * was, and the package is then handed to PowerPoint claiming it.
     *
     * Both halves are the empty string once `getAttribute` has answered null,
     * which is why they match: the part's missing extension is not a key to
     * look anything up by.
     */
    const pkg = await craft({
      [CONTENT_TYPES]: types(
        `<Default ContentType="application/x-mystery"/>`,
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
        `<Default Extension="xml" ContentType="application/xml"/>`,
      ),
      "ppt/embeddings/workbook": "no extension at all",
    });
    expect(await pkg.contentTypeOf("ppt/embeddings/workbook")).toBeUndefined();
    // The Defaults that DO name an extension still answer.
    expect(await pkg.contentTypeOf("_rels/.rels")).toBe("application/vnd.openxmlformats-package.relationships+xml");
  });

  it("says nothing for a Default that names no content type, rather than answering null", async () => {
    // The same distinction `pptx-malformed.test.ts` draws for an Override, on
    // the other arm of the same function. `getAttribute` answers null and the
    // signature promises `string | undefined`; a null reaching a caller that
    // has already checked for undefined is the shape of bug that survives
    // typechecking.
    const pkg = await craft({
      [CONTENT_TYPES]: types(
        `<Default Extension="png"/>`,
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
        `<Default Extension="xml" ContentType="application/xml"/>`,
      ),
      "ppt/media/image1.png": "not really a png",
    });
    expect(await pkg.contentTypeOf("ppt/media/image1.png")).toBeUndefined();
  });
});

describe("reading a part the package is already holding", () => {
  it("hands peek the parsed copy rather than a second parse of it", async () => {
    /**
     * `peek` exists so that a reader does not pay to hold a part for the rest
     * of the run: on the file route the deck is the user's WHOLE presentation,
     * and a pass over it parsed three hundred documents before it had done
     * anything. Reading a part that is ALREADY parsed through the same door —
     * serialising the held copy and parsing the text again — gives that reader
     * a second copy of the biggest part in the deck while the first is still
     * held, which is the cost `peek` was written to avoid, and hands back
     * something that is not the document the package will write out.
     */
    const pkg = await deck(TWO);
    const held = await pkg.doc(SLIDE1);
    expect(await pkg.peek(SLIDE1, (doc) => doc === held), "peek parsed the part a second time").toBe(true);
    // And the other half of the contract, which is what the identity above is
    // for: a part peek brought in itself is not kept.
    const before = pkg.cachedParts();
    await pkg.peek("ppt/slides/slide2.xml", (doc) => doc.documentElement.localName);
    expect(pkg.cachedParts(), "peek retained the part").toBe(before);
  });
});

describe("the two places base64 is converted", () => {
  it("hands bytes straight to the zip rather than through the base64 decoder", async () => {
    /**
     * `open` takes bytes or the base64 a host hands over, and only the second
     * needs decoding. Sending bytes through the decoder as well is not wrong —
     * `Buffer.from` copies a Uint8Array and hands back the same bytes — it is
     * a whole extra copy of the deck, which on the file route is the user's
     * entire presentation and is measured in tens of megabytes.
     *
     * Asserted on WHICH object reaches JSZip, because that is the difference:
     * the caller's own bytes, or a copy of them.
     */
    const bytes = await makeDeck(ONE);
    const loading = vi.spyOn(JSZip, "loadAsync");
    const pkg = await Pkg.open(bytes);
    expect(await pkg.slidePaths()).toHaveLength(1);
    expect(loading.mock.calls[0]?.[0], "the deck was copied on the way in").toBe(bytes);
    loading.mockRestore();
  });

  it("encodes the base64 itself rather than asking the zip to generate it again", async () => {
    /**
     * The measured half of the same decision, the other way round: on 45 MB,
     * JSZip's own encoder costs 2.5 seconds of character shuffling and the
     * platform's costs 24 (`base64.ts` has the table). `toBase64` builds the
     * bytes once and encodes them; falling back to `generateAsync` when the
     * platform HAS a route means zipping the whole package twice for one
     * insert.
     *
     * The fallback arm itself stays uncovered on purpose — see `open`'s
     * comment. This pins that it is a fallback.
     */
    const pkg = await deck(ONE);
    const generating = vi.spyOn(JSZip.prototype, "generateAsync");
    const base64 = await pkg.toBase64();
    expect(base64.startsWith("UEsDB"), "not a zip").toBe(true);
    expect(
      generating.mock.calls.map((call) => (call[0] as { type?: string }).type),
      "the package was generated twice for one base64 string",
    ).toEqual(["uint8array"]);
    generating.mockRestore();
  });
});
