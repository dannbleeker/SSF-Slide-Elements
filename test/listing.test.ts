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
    // "117 elements in each of the two slide sizes" is a number a reader takes
    // as a promise, and the library is the owner's to add to. A count copied
    // once and never checked is the classic figure that rots.
    const index = JSON.parse(readFileSync("public/catalogue/catalogue.json", "utf8")) as {
      sizes: Record<string, { elements: unknown[] }>;
    };
    const counts = Object.values(index.sizes).map((s) => s.elements.length);
    expect(new Set(counts).size, "the two sizes hold different numbers of elements").toBe(1);
    expect(LISTING).toContain(`${counts[0] as number} elements in each of the two slide sizes`);
  });
});

describe("the listing's own files", () => {
  it("names a store logo that exists", () => {
    const logo = /`(public\/assets\/store-\d+\.png)`/.exec(LISTING)?.[1];
    expect(logo).toBeTruthy();
    expect(existsSync(logo as string), logo as string).toBe(true);
  });

  it("still says the screenshot and the name are the owner's", () => {
    // The point of this one is the DIRECTION it fails in. Neither can be
    // produced without a real PowerPoint, and the failure mode worth guarding
    // is somebody — including a future me — quietly deciding a composited
    // screenshot will do. If one is genuinely done, this test is what makes
    // removing the caveat a deliberate act: it went red the day the test deck
    // was written, and was edited rather than deleted.
    const owners = LISTING.slice(LISTING.indexOf("## Still the owner's"));
    expect(owners).toContain("1366×768");
    expect(owners).toContain("naming policy");
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
