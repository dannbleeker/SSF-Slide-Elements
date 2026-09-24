import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

/**
 * The store listing, held to what the product actually is.
 *
 * `docs/LISTING.md` is copy: nothing imports it, no build reads it, and it is
 * the text a stranger decides on. That is exactly the kind of file that goes
 * quietly wrong — a name changed in the manifest, an element added to the
 * library, a support page moved — and nothing notices until a reviewer at
 * Microsoft does, weeks later.
 *
 * So every claim in it that has a source of truth in this repo is asserted
 * against that source here. Claims that have none (what the pane feels like to
 * use) are the owner's to write and are left alone.
 */

const LISTING = readFileSync("docs/LISTING.md", "utf8");

/** A value out of the fields table, by its row label. */
function field(name: string): string {
  const row = new RegExp(`^\\| ${name} \\| (.+?) \\|$`, "m").exec(LISTING);
  if (!row) throw new Error(`docs/LISTING.md has no "${name}" row`);
  return (row[1] as string).trim();
}

describe("the listing says what the manifests say", () => {
  it("carries the same name, provider and short description", async () => {
    // Three strings in two places. The manifest's are what Office shows once
    // the add-in is installed and the listing's are what the store shows before
    // it is, and a user who sees two different names has met two products.
    // @ts-expect-error — plain .mjs with no types.
    const { DEFINITION } = await import("../scripts/manifest-source.mjs");
    expect(field("Name")).toBe(DEFINITION.displayName);
    expect(field("Provider")).toBe(DEFINITION.provider);
    expect(field("Short description")).toBe(DEFINITION.shortDescription);
  });

  it("points at the same support and privacy pages, which is what certification checks", () => {
    // Two manifests, and they carry these differently: the XML one has a
    // `SupportUrl` element and NO privacy element at all — an XML-manifest
    // submission states the privacy URL on the form instead — while the
    // unified JSON one carries both. So each is asserted where it actually
    // lives, rather than against whichever file was convenient.
    // The doc writes a URL as <https://…>, which is how Markdown makes one a
    // link without repeating it.
    const xml = readFileSync("manifest-prod.xml", "utf8");
    const support = /<SupportUrl DefaultValue="([^"]*)"/.exec(xml)?.[1];
    expect(support, "SupportUrl").toBeTruthy();
    expect(field("Support URL")).toBe(`<${support as string}>`);

    const json = JSON.parse(readFileSync("manifest-prod.json", "utf8")) as {
      developer: { privacyUrl: string; websiteUrl: string };
    };
    expect(field("Privacy URL")).toBe(`<${json.developer.privacyUrl}>`);
  });

  it("points at terms we publish ourselves, on a page that is in this repo", () => {
    // Until 2026-09-16 this was Microsoft's standard EULA, offered to
    // publishers with no lawyer of their own. The standard text is written for
    // an add-in that might do anything, so it could say nothing about the two
    // properties that matter here — that the presentation never leaves the
    // pane, and that an insert REWRITES one slide of the file it is given.
    //
    // `termsOfUseUrl` lives in the JSON manifests only; the XML ones have no
    // terms element, which is why moving it needed no re-sideload.
    const json = JSON.parse(readFileSync("manifest-prod.json", "utf8")) as {
      developer: { termsOfUseUrl: string };
    };
    const terms = json.developer.termsOfUseUrl;
    expect(field("Terms")).toBe(`<${terms}>`);
    // On OUR origin, and a file that ships. Microsoft's submission rules forbid
    // a listing URL pointing at a GitHub repository, so a page that is only in
    // the repository is not an answer.
    expect(terms).toContain("ssf-slide-elements.struktureretsundfornuft.dk");
    expect(existsSync("public/terms.html"), "the terms URL has no page behind it").toBe(true);
  });

  it("publishes the licence on its own origin rather than pointing at the repository", () => {
    const licence = field("Licence");
    expect(licence).toContain("ssf-slide-elements.struktureretsundfornuft.dk");
    expect(existsSync("public/license.html")).toBe(true);
    // The page QUOTES the licence rather than summarising it, and the copy in
    // it has to be the repository's own. A summary of a licence is not a
    // licence, and a quoted licence that has drifted from the real one is
    // worse than either.
    const page = readFileSync("public/license.html", "utf8");
    const quoted = /<pre>([\s\S]*?)<\/pre>/.exec(page)?.[1];
    expect(quoted, "public/license.html quotes no licence text").toBeTruthy();
    const normalise = (s: string) =>
      s
        .replace(/&quot;/g, '"')
        .replace(/&amp;/g, "&")
        .replace(/\r\n/g, "\n")
        .trim();
    expect(normalise(quoted as string)).toBe(normalise(readFileSync("LICENSE", "utf8")));
  });

  /**
   * The publisher is named in four places and they must agree.
   *
   * It was "StruktureretSundFornuft" in the manifests and on the pages here,
   * while the sibling SSF Merge's terms page said "StruktureretSundFornuft
   * ApS". Two siblings disagreeing about the name of the company that publishes
   * them is the kind of thing nobody notices until Partner Center asks, and the
   * owner settled it on 2026-09-16: the ApS is correct, because that is the
   * registered entity and a publisher display name has to match it.
   *
   * So it is asserted rather than remembered. The manifest is the source, and
   * the two policy pages that name a publisher are held to it — the ones a
   * reviewer opens from the listing.
   */
  it("names the same publisher in the manifests and on the pages that state one", async () => {
    // @ts-expect-error — plain .mjs with no types.
    const { DEFINITION } = await import("../scripts/manifest-source.mjs");
    const provider = DEFINITION.provider as string;

    const xml = readFileSync("manifest-prod.xml", "utf8");
    expect(/<ProviderName>([^<]*)<\/ProviderName>/.exec(xml)?.[1]).toBe(provider);
    const json = JSON.parse(readFileSync("manifest-prod.json", "utf8")) as { developer: { name: string } };
    expect(json.developer.name).toBe(provider);

    for (const page of ["public/privacy.html", "public/terms.html"]) {
      expect(
        readFileSync(page, "utf8"),
        `${page} states a publisher name, and it is not the one the manifests carry`,
      ).toContain(`<strong>${provider}</strong>`);
    }
  });

  it("names the ribbon button by the label the manifest gives it", () => {
    // The testing notes tell a validator what to click. A label that has moved
    // on leaves them looking for a button that is not there, and the first
    // thing they conclude is that the add-in did not install.
    const xml = readFileSync("manifest-prod.xml", "utf8");
    const label = /<bt:String id="OpenPane.Label" DefaultValue="([^"]*)"/.exec(xml)?.[1];
    expect(label).toBeTruthy();
    expect(LISTING).toContain(`click **${label as string}**`);
  });
});

describe("the listing says what the library holds", () => {
  it("quotes the element count the committed catalogue actually has", () => {
    // "106 elements in each of the two slide sizes" is a number a reader takes
    // as a promise, and the library is the owner's to add to. A count copied
    // once and never checked is the classic figure that rots.
    const index = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
      sizes: Record<string, { elements: unknown[] }>;
    };
    const counts = Object.values(index.sizes).map((s) => s.elements.length);
    expect(new Set(counts).size, "the two sizes hold different numbers of elements").toBe(1);
    expect(LISTING).toContain(`${counts[0] as number} elements in each of the two slide sizes`);
  });

  /**
   * The copy must not offer a category the library has not got.
   *
   * The count above was checked and correct while the long description promised
   * **flowchart shapes** — twice — for the ten days after the owner had that
   * whole category taken out (#105). A number that is watched and a sentence
   * that is not is exactly how a listing ends up describing a product that no
   * longer exists, and the reader it misleads is a Microsoft reviewer or a
   * customer deciding whether to install.
   *
   * Named categories only. The copy also says "hierarchies", "matrices" and
   * "triangles", which are kinds of element rather than categories, and holding
   * prose to a vocabulary would be a gate that fails on good writing. What this
   * catches is the case that actually happened: a phrase that IS one of the
   * library's own category names, still in the copy after the category went.
   */
  it("offers no category the library has stopped having", () => {
    const index = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
      sizes: Record<string, { categories: { name: string }[] }>;
    };
    const live = new Set(Object.values(index.sizes).flatMap((s) => s.categories.map((c) => c.name.toLowerCase())));
    // Every category name this library has ever shipped, so a removal is what
    // the sweep notices. Adding one here is how a category that goes away later
    // stays watched.
    const everShipped = [
      "White boxes",
      "White boxes with black headings",
      "Lines only",
      "Process flows",
      "Tables",
      "Triangles",
      "Document structure",
      "Grey boxes",
      "Markers",
      "Stamps and labels",
      "One-page templates",
      "Flowchart shapes",
      "Icons",
    ];
    const gone = everShipped.filter((name) => !live.has(name.toLowerCase()));
    const description = LISTING.slice(LISTING.indexOf("## Long description"));
    expect(
      gone.filter((name) => new RegExp(name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i").test(description)),
      "the long description offers a category the library no longer has. It is the text a reviewer and a customer read",
    ).toEqual([]);
  });
});

describe("the listing's own files", () => {
  it("names a store logo that exists", () => {
    const logo = /`(public\/assets\/store-\d+\.png)`/.exec(LISTING)?.[1];
    expect(logo).toBeTruthy();
    expect(existsSync(logo as string), logo as string).toBe(true);
  });

  it("still says the screenshot is the owner's, and that the name was decided rather than cleared", () => {
    // The point of this one is the DIRECTION it fails in. The screenshot cannot
    // be produced without a real PowerPoint, and the failure mode worth guarding
    // is somebody — including a future me — quietly deciding a composited
    // screenshot will do. If one is genuinely done, this test is what makes
    // removing the caveat a deliberate act: it went red the day the test deck
    // was written, and was edited rather than deleted. It went red again on
    // 2026-09-16 when the owner settled the NAME, and was edited again.
    const owners = LISTING.slice(LISTING.indexOf("## Still the owner's"));
    expect(owners).toContain("1366×768");
    // The name is decided, and what this now holds is the HONESTY of how it is
    // recorded. Nobody read it against certification policy 1100.7, and a
    // record that said "checked" would licence a later reader to tell a
    // reviewer it was. If someone softens this into a clearance, this goes red.
    expect(owners, "the listing name's record must not claim a policy clearance nobody performed").toContain(
      "not as a policy clearance",
    );
    expect(owners, "the test deck is written; it no longer belongs in this list").not.toContain("test deck");
  });

  it("names a test deck that exists, is a presentation, and is not one this repo assembled", () => {
    // The claim in the table is that PowerPoint wrote it. A .pptx this project
    // built would satisfy "a file is there" and defeat the whole reason the
    // deck is attached to the submission.
    const deck = /`(template\/[\w-]+\.pptx)`, three slides, written by PowerPoint/.exec(LISTING)?.[1];
    expect(deck, "the listing does not name a test deck").toBeTruthy();
    expect(existsSync(deck as string), deck as string).toBe(true);
    const bytes = readFileSync(deck as string);
    expect(bytes.subarray(0, 2).toString("latin1"), "not a zip").toBe("PK");
    // That it is a presentation PowerPoint itself wrote is checked where the
    // zip is already open: `test/validators-deck.test.ts` reads its
    // `docProps/app.xml`. The bytes here are DEFLATED, so looking for the
    // application's name in them finds nothing whatever the file is — which is
    // how this assertion read before it was tried.
    expect(bytes.subarray(30, 60).toString("latin1")).toContain("[Content_Types].xml");
  });
});

describe("the testing notes admit which platforms nobody has measured", () => {
  /**
   * The disclosure the design record has been claiming for days.
   *
   * `docs/DESIGN.md` section 9 said "the testing notes say the publisher has
   * not measured it" about iPad. They did not — the sentence was never
   * written, and nothing was checking, which is the exact shape of the four
   * stale claims found on 2026-09-11 one level up. Read against the answer
   * sheets rather than a hand-kept list: a platform with no sheet must be
   * disclosed, and a platform WITH one must not be, so the day a Mac sheet is
   * filed this goes red until the notes stop calling Mac unmeasured.
   *
   * Certification policy 1120 has validators exercise every platform the
   * manifest claims, and a PowerPoint task pane cannot exclude any of them. So
   * the notes telling a validator that a Mac or iPad finding is a FIRST
   * measurement is the difference between a useful report and a surprise.
   */

  /** `Office.context.diagnostics.platform`, in the words the listing uses. */
  const AS_LISTED: Record<string, string> = {
    OfficeOnline: "web",
    PC: "Windows",
    Mac: "Mac",
    iOS: "iPad",
  };

  /** Every platform an answer sheet under `docs/host-answers/` was taken on. */
  function measured(): Set<string> {
    const out = new Set<string>();
    for (const name of readdirSync("docs/host-answers")) {
      if (!name.endsWith(".json")) continue;
      const sheet = JSON.parse(readFileSync(`docs/host-answers/${name}`, "utf8")) as { platform?: string };
      const listed = AS_LISTED[sheet.platform ?? ""];
      if (listed) out.add(listed);
    }
    return out;
  }

  it("discloses every platform with no answer sheet, and claims none that has one", () => {
    const taken = measured();
    expect(taken.size, "no answer sheets read — the sweep is broken, not the listing").toBeGreaterThan(0);
    // The Products row is what the submission claims to run on.
    const products = field("Products");
    const claimed = Object.values(AS_LISTED).filter((name) => products.includes(name));
    expect(claimed.length, "the Products row names no platform this sweep knows").toBeGreaterThan(2);

    const notes = LISTING.slice(LISTING.indexOf("## Testing notes"));
    const said = notes.toLowerCase().replace(/\s+/g, " ");
    for (const platform of claimed) {
      const disclosed =
        said.includes(`not been measured on ${platform.toLowerCase()}`) || unmeasuredPair(said, platform);
      if (taken.has(platform)) {
        expect(disclosed, `${platform} has an answer sheet and the notes still call it unmeasured`).toBe(false);
      } else {
        expect(disclosed, `no answer sheet was taken on ${platform} and the notes do not say so`).toBe(true);
      }
    }
  });

  /** "not been measured on Mac or on iPad" discloses both, not just the first. */
  function unmeasuredPair(said: string, platform: string): boolean {
    const at = said.indexOf("not been measured on");
    if (at < 0) return false;
    return said.slice(at, at + 60).includes(platform.toLowerCase());
  }
});

describe("the screenshots the submission carries", () => {
  /**
   * The listing's pictures, held to the one thing a reviewer rejects on: size.
   *
   * AppSource asks for 1366×768. A capture at the wrong size **looks right** —
   * that is the whole danger, and it is not hypothetical: the first attempt at
   * `scripts/listing-shot.ps1` produced 1366×2180 because a maximised window
   * ignores `MoveWindow`'s height, and an earlier one would have been wrong
   * again because Windows lies about sizes to a process that has not declared
   * itself DPI-aware. Neither is visible by looking at the image.
   *
   * SWEPT rather than listed. This pinned one file while the submission carries
   * five, so four pictures that a reviewer rejects on exactly this were covered
   * by nothing — and a sixth added later would have been uncovered too. The
   * glob is the gate: anything named `docs/listing-*.png` is in the submission
   * and is checked.
   *
   * Read out of the PNG header rather than with a library: the IHDR chunk's
   * width and height are the two big-endian integers at bytes 16 and 20, and
   * the file is checked to be a PNG first so a JPEG renamed `.png` cannot pass
   * by having plausible bytes there.
   */
  const SHOTS = readdirSync("docs")
    .filter((name) => /^listing-.*\.png$/.test(name))
    .sort()
    .map((name) => `docs/${name}`);

  it("finds every picture the submission carries", () => {
    // A sweep that finds nothing passes every test below it, which is the way
    // a glob-driven gate fails silently.
    expect(SHOTS.length, `found: ${SHOTS.join(", ")}`).toBeGreaterThanOrEqual(5);
    expect(SHOTS).toContain("docs/listing-screenshot.png");
  });

  for (const shot of SHOTS) {
    const bytes = readFileSync(shot);

    it(`${shot} is a PNG, not something renamed to look like one`, () => {
      expect(bytes.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
      // The first chunk of a PNG is IHDR, which is where the size below is read.
      expect(bytes.subarray(12, 16).toString("ascii")).toBe("IHDR");
    });

    it(`${shot} is exactly the size the store asks for`, () => {
      expect(bytes.readUInt32BE(16), "width").toBe(1366);
      expect(bytes.readUInt32BE(20), "height").toBe(768);
    });

    it(`${shot} is small enough to upload without anybody thinking about it`, () => {
      expect(bytes.length).toBeLessThan(2_000_000);
    });
  }

  it("is the file the listing notes point at", () => {
    expect(LISTING).toContain("listing-screenshot.png");
  });
});
