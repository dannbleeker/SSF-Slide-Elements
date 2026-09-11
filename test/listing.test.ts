import { existsSync, readFileSync } from "node:fs";
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

  it("still says the screenshot, the test deck and the name are the owner's", () => {
    // The point of this one is the DIRECTION it fails in. These three cannot be
    // produced without a real PowerPoint, and the failure mode worth guarding
    // is somebody — including a future me — quietly deciding a generated deck
    // or a composited screenshot will do. If they are genuinely done, this test
    // is what makes removing the caveat a deliberate act.
    const owners = LISTING.slice(LISTING.indexOf("## Still the owner's"));
    expect(owners).toContain("1366×768");
    expect(owners).toContain("test deck");
    expect(owners).toContain("naming policy");
  });
});
