import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Pkg, resolveTarget } from "../src/core/pptx/pkg.js";
import { parseXml } from "../src/core/pptx/xml.js";
import { makeDeck } from "./fixtures/deck.js";

const ROOT_RELS_PATH = "_rels/.rels";

/**
 * The claims on the front of SECURITY.md, as tests.
 *
 * That page says the add-in makes no network calls and never writes markup.
 * Both are properties of the SOURCE, so both can be read off it — and a
 * security page whose claims nothing re-checks is the failure the page's own
 * preamble warns about. The package-level cases below are ported from
 * SSF-Merge's sweeps of 2026-08-29 and 2026-08-30. The splice brought the
 * first thing here that writes text, and its value-as-text cases went with it,
 * into `pptx-tags.test.ts` beside the escaping they are about.
 */
const sources = (): string[] => {
  const out: string[] = [];
  const walk = (dir: string): void => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = `${dir}/${entry.name}`;
      if (entry.isDirectory()) walk(path);
      else if (/\.(ts|html)$/.test(entry.name)) out.push(path);
    }
  };
  walk("src");
  return out;
};

const code = (path: string): string =>
  readFileSync(path, "utf8")
    // Comments are stripped, because this repo's files explain themselves and
    // several of them name these APIs in prose.
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/^\s*\/\/.*$/gm, "");

describe("the claims on the front of SECURITY.md are executable", () => {
  it("reads the source at all", () => {
    // The vacuity guard: an empty file list would satisfy every assertion
    // below forever.
    expect(sources().length).toBeGreaterThan(5);
  });

  /**
   * This used to forbid `fetch(` outright, and that was the right guard for a
   * pane with nothing to load. The picker has something to load: the catalogue
   * is static files on the add-in's own site, because the element markup and
   * the parts it carries come to about 16 MB and `docs/DESIGN.md` sections 3
   * and 11 put them behind lazy requests rather than in the bundle. A blanket
   * ban and the design record could not both be right, and the record is what
   * the build is held to.
   *
   * So the guard is narrower and STRONGER: it now asserts the property anybody
   * actually cares about rather than a proxy for it. The add-in can only GET a
   * static file from its own origin, and has no way to send anything anywhere.
   * One named file may fetch; it may not name an absolute URL, and it may not
   * pass request options, which is where a method, a body or a header would
   * have to go.
   */
  const LOADER = "src/pane/catalogue.ts";

  it("sends nothing anywhere: no channel that could carry data out exists at all", () => {
    for (const path of sources()) {
      const src = code(path);
      for (const call of ["XMLHttpRequest", "WebSocket", "sendBeacon"]) {
        expect(src.includes(call), `${path} uses ${call}`).toBe(false);
      }
    }
  });

  it("reaches the network from one named file and nowhere else", () => {
    const others = sources()
      .map((p) => p.replaceAll("\\", "/"))
      .filter((p) => !p.endsWith(LOADER))
      .filter((p) => code(p).includes("fetch("));
    expect(others, `only ${LOADER} may fetch`).toEqual([]);
    // The vacuity guard: if the loader stops fetching, this block would pass by
    // finding nothing and the two rules below would be asserting about nothing.
    expect(code(LOADER).includes("fetch("), "the loader no longer fetches, so these rules check nothing").toBe(true);
  });

  it("can only GET, and only from the origin the pane was served from", () => {
    const loader = code(LOADER);
    // An absolute or protocol-relative URL is a request to somewhere else.
    expect(/["'`](?:https?:)?\/\//.test(loader), `${LOADER} names an absolute URL`).toBe(false);
    // `fetch` with one argument is a GET and can carry nothing. A second
    // argument is the only place a method, a body or a header could go.
    expect(/fetch\([^)]*,/.test(loader), `${LOADER} passes request options to fetch`).toBe(false);
    for (const sink of ["method:", "body:", "headers:"]) {
      expect(loader.includes(sink), `${LOADER} sets ${sink}`).toBe(false);
    }
  });

  it("writes text, never markup", () => {
    for (const path of sources()) {
      const src = code(path);
      for (const sink of ["outerHTML", "insertAdjacentHTML", "document.write", "eval(", "new Function"]) {
        expect(src.includes(sink), `${path} uses ${sink}`).toBe(false);
      }
      // `innerHTML` is READ once, in the page's own no-Office fallback, to ask
      // whether anything has been drawn yet. Assigning to it is what this
      // forbids.
      expect(/innerHTML\s*=/.test(src), `${path} assigns innerHTML`).toBe(false);
    }
  });
});

describe("a relationship target comes out of the deck", () => {
  /**
   * `resolveTarget` honours a leading `/` and any number of `..`, which is what
   * the format says to do — so it can and does answer a path outside the part's
   * own directory. That is fine for READING, where a path that names nothing
   * simply is not found. It matters where the answer is used to DELETE.
   */
  it("can name a part outside the slide's directory", () => {
    const owner = "ppt/slides/slide1.xml";
    expect(resolveTarget(owner, "/[Content_Types].xml")).toBe("[Content_Types].xml");
    expect(resolveTarget(owner, "../../[Content_Types].xml")).toBe("[Content_Types].xml");
    // `..` past the root is clamped by `pop()` on an empty array rather than
    // escaping into something above it: the answer stays a package-relative
    // name, so nothing here can reach a real filesystem.
    expect(resolveTarget(owner, "../../../../../../etc/passwd")).toBe("etc/passwd");
  });

  it("does not let a crafted notes target delete a part the deck needs", async () => {
    // A deck arrives from anywhere. Removing a slide collects that slide's
    // notes page and comments, and before this was guarded a target of
    // `/[Content_Types].xml` had it delete the one part a presentation cannot
    // open without — turning a merge of somebody else's deck into a file that
    // will not open.
    const pkg = await Pkg.open(await makeDeck([{ paragraphs: [["Cover"]] }, { paragraphs: [["Second"]] }]));
    const relsPath = Pkg.relsPathFor("ppt/slides/slide1.xml");
    const rels =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="/[Content_Types].xml"/>` +
      `</Relationships>`;
    pkg.setText(relsPath, rels);

    expect(pkg.has("[Content_Types].xml"), "the fixture should start with one").toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("[Content_Types].xml"), "a crafted target deleted it").toBe(true);
  });
});

describe("a chart's own relationships come out of the deck too", () => {
  /**
   * The same class of hole as the one above, one level down — found only by
   * looking for siblings after fixing the first, which is the step that gets
   * skipped.
   *
   * Removing a slide sweeps the parts its charts and diagrams own. The parent
   * of that list was held to an allowlist; the CHILD was not. So a chart whose
   * relationships named `/ppt/presentation.xml` had that part counted as
   * something the chart owned, nothing else in the package referred to it — its
   * only referrer is the root `_rels/.rels`, which the referrer scan does not
   * read — and the sweep took it.
   *
   * `/ppt/presentation.xml` was the worse one: deleted SILENTLY, the merge
   * finished, and the output cannot open. `/[Content_Types].xml` was deleted and
   * then threw.
   */
  async function deckWhoseChartClaims(victim: string) {
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    pkg.setText(
      "ppt/charts/_rels/chart1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="${victim}"/>` +
        `</Relationships>`,
    );
    return pkg;
  }

  for (const victim of ["/ppt/presentation.xml", "/[Content_Types].xml", "/ppt/slides/slide2.xml"]) {
    it(`keeps ${victim} when a chart claims to own it`, async () => {
      const pkg = await deckWhoseChartClaims(victim);
      const key = victim.slice(1);
      expect(pkg.has(key), "the fixture should start with it").toBe(true);
      await pkg.removeSlide("ppt/slides/slide1.xml");
      expect(pkg.has(key), `a crafted chart relationship swept ${key}`).toBe(true);
    });
  }

  it("counts the package's own relationships as referrers", async () => {
    /**
     * The referrer scan tested every `.rels` whose path contained `/_rels/`,
     * which is every one EXCEPT the package's own `_rels/.rels`. So a part
     * named only from the root was invisible to it and looked unreferenced.
     *
     * `ppt/presentation.xml` is the real instance, and the allowlist above now
     * keeps it out of reach — which means this needs a part the allowlist DOES
     * admit to be observable at all. A picture referenced from the root, the
     * way a thumbnail is, and claimed by the chart on the way out.
     */
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    pkg.setBytes("ppt/media/image9.png", new Uint8Array([137, 80, 78, 71]));
    const root = await pkg.text(ROOT_RELS_PATH);
    pkg.setText(
      ROOT_RELS_PATH,
      root.replace(
        "</Relationships>",
        `<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/thumbnail" Target="ppt/media/image9.png"/></Relationships>`,
      ),
    );
    pkg.setText(
      "ppt/charts/_rels/chart1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.png"/>` +
        `</Relationships>`,
    );

    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("ppt/media/image9.png"), "a part referenced only from the root was swept").toBe(true);
  });

  it("still sweeps what the slide's chart really does own", async () => {
    // The guard must not become "never sweep anything": a template slide going
    // out on the whole-deck route has to take its chart with it, or the output
    // ships parts nothing points at.
    const pkg = await Pkg.open(
      await makeDeck([
        { paragraphs: [["Cover"]], chart: { title: "T", workbook: ["a"] } },
        { paragraphs: [["Second"]] },
      ]),
    );
    expect(pkg.has("ppt/charts/chart1.xml")).toBe(true);
    await pkg.removeSlide("ppt/slides/slide1.xml");
    expect(pkg.has("ppt/charts/chart1.xml"), "the chart should have gone with its slide").toBe(false);
  });
});

describe("the XML parser does not fetch or expand what a deck tells it to", () => {
  /**
   * From the sweep of 2026-08-30. A .pptx is a zip of XML and a user can be
   * sent one, so the two classic parser attacks are the first question anybody
   * asks of this codebase — and until now the answer was a property of a
   * dependency that nothing here had checked.
   *
   * `@xmldom/xmldom` refuses both. These run so that a version bump which
   * changes its mind is a red test rather than a discovery.
   */
  it("does not resolve an external entity", () => {
    const doc = parseXml('<?xml version="1.0"?>\n<!DOCTYPE r [ <!ENTITY x SYSTEM "file:///etc/passwd"> ]>\n<r>&x;</r>');
    // The reference comes through as TEXT. Not "the file was empty" — the
    // parser never went looking, which is why the literal is still there.
    expect(doc.documentElement?.textContent).toBe("&x;");
  });

  it("does not expand a nested entity bomb", () => {
    const bomb =
      '<?xml version="1.0"?>\n<!DOCTYPE r [\n' +
      '<!ENTITY a "aaaaaaaaaa">\n' +
      '<!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">\n' +
      '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">\n' +
      '<!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">\n' +
      '<!ENTITY e "&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;">\n' +
      "]>\n<r>&e;</r>";
    // Four levels is 100,000 characters if expanded, and the published attack
    // has nine. Custom entities are not expanded at all, so what comes out is
    // the reference itself.
    const text = parseXml(bomb).documentElement?.textContent ?? "";
    expect(text.length).toBeLessThan(100);
    expect(text).toBe("&e;");
  });
});
