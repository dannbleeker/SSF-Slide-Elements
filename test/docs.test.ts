import { existsSync, readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { EMPTY, STEP_TITLE, STEPS, primary } from "../src/pane/steps.js";

/**
 * The lockstep guard.
 *
 * Documentation that is updated when someone remembers is documentation that is
 * wrong, and stale docs are worse than none: a reader trusts them. These tests
 * read the surfaces out of the SOURCE and fail when the docs have not kept up,
 * so a feature and its documentation land in the same change or neither does.
 *
 * Small today, because the pane is small. Every hook here grows with the
 * thing it reads: a new step title, a new button, a new picture.
 */

const manual = readFileSync("docs/MANUAL.md", "utf8");
const readme = readFileSync("README.md", "utf8");
const changelog = readFileSync("CHANGELOG.md", "utf8");
const backlog = readFileSync("docs/BACKLOG.md", "utf8");

/** The manual, whitespace-collapsed: the formatter wraps a quoted label across lines. */
const prose = manual.replace(/\s+/g, " ");

describe("the manual keeps up with the pane", () => {
  it("names every step the pane can draw", () => {
    expect(STEPS.length).toBeGreaterThan(0);
    for (const id of STEPS) {
      const title = STEP_TITLE[id];
      expect(title.length).toBeGreaterThan(0);
      expect(prose, `step "${title}" is not in the manual`).toContain(title);
    }
  });

  it("quotes button labels that still exist in the pane", () => {
    // Read from the DECISION, not from a list kept here: a rename lands in
    // `steps.ts` and this test asks the manual to follow it.
    const labels = STEPS.map((id) => primary(EMPTY, id).label);
    expect(labels.length).toBeGreaterThan(0);
    const source = readFileSync("src/pane/steps.ts", "utf8") + readFileSync("src/pane/render.ts", "utf8");
    for (const label of labels) {
      expect(source, `the pane no longer has a "${label}" button`).toContain(label);
      expect(prose, `the manual does not mention "${label}"`).toContain(label);
    }
  });

  it("describes the pane that is built rather than calling it planned", () => {
    // Claiming less than you have is not the harmless direction: it tells a
    // reader not to look for the thing that is there. A sibling's manual said
    // its pane was not written for days after it shipped.
    const at = manual.indexOf("## The pane");
    expect(at, "the manual has no pane section").toBeGreaterThan(-1);
    const section = manual.slice(at + 3);
    const body = section.slice(0, section.indexOf("\n## "));
    expect(body.toLowerCase()).not.toContain("planned");
  });

  /**
   * A feature table row that calls a built thing planned — anywhere in the
   * file, not only under one heading.
   *
   * The guard above this one reads the manual's pane section and nothing else,
   * and on 2026-09-11 every word of it passed while the manual's status block
   * told the reader "there is nothing to insert yet" and three rows of its
   * feature table said `planned` for the picker, the insert and the undo. The
   * README said the same. All three had shipped weeks earlier. A guard that
   * only looks under one heading is a guard the stale text simply sits outside
   * of.
   *
   * So each row below is pinned to a SYMBOL, and the pair is what makes it a
   * test rather than a list: if the symbol is gone the row is allowed to say
   * planned again, and the first half of the assertion fails loudly if the
   * symbol was renamed rather than removed. Prose is never matched against
   * prose.
   */
  const BUILT: Array<{ row: string; file: string; proof: string }> = [
    { row: "The picker", file: "src/pane/steps.ts", proof: `"browse"` },
    { row: "The insert", file: "src/pane/main.ts", proof: "insertPackage(" },
    { row: "Taking it back", file: "src/host/insert.ts", proof: "export function undoPlan" },
    { row: "Jumping to a slide", file: "src/host/jump.ts", proof: "export function jumpOutcome" },
  ];

  it.each(BUILT)("does not call $row planned while $proof is in $file", ({ row, file, proof }) => {
    expect(readFileSync(file, "utf8"), `${file} no longer contains ${proof}`).toContain(proof);
    for (const [name, text] of [
      ["the manual", manual],
      ["the README", readme],
    ] as const) {
      // The row's own line, wherever the table sits in the file.
      const line = text.split("\n").find((l) => l.startsWith(`| ${row} |`));
      if (line === undefined) continue;
      expect(line.toLowerCase(), `${name} still calls "${row}" planned`).not.toContain("planned");
    }
  });

  it("never tells the reader the add-in cannot insert", () => {
    // The status sentence, rather than the table. Same defect, different
    // shape: a table row can be right while the paragraph above it is not.
    expect(readFileSync("src/pane/main.ts", "utf8")).toContain("insertPackage(");
    for (const denial of ["nothing to insert", "inserts nothing", "reads nothing and writes nothing"]) {
      for (const [name, text] of [
        ["the manual", manual],
        ["the README", readme],
        ["the security page", readFileSync("SECURITY.md", "utf8")],
        // The pages the public actually reads. The site's front page said "it
        // inserts nothing yet" for three days after the insert shipped, and
        // nothing in this guard was looking at it.
        ["the site's front page", readFileSync("public/index.html", "utf8")],
        ["the privacy page", readFileSync("public/privacy.html", "utf8")],
      ] as const) {
        expect(text.toLowerCase(), `${name} still says "${denial}"`).not.toContain(denial);
      }
    }
  });

  it("says which parts are not built yet, rather than describing them as shipped", () => {
    // The manual documents a design that is ahead of the code. That is fine as
    // long as it never claims to be behind it, and as long as the promise is
    // spelled the same way everywhere: `planned`.
    expect(manual).toContain("planned");
  });
});

describe("the documentation set is whole", () => {
  it("keeps an Unreleased section in the changelog", () => {
    expect(changelog).toContain("## [Unreleased]");
  });

  it("keeps a rejected list in the backlog, so the same idea is not re-proposed", () => {
    expect(backlog).toContain("Rejected");
  });

  it("points at every document from the README", () => {
    for (const doc of ["docs/MANUAL.md", "docs/BACKLOG.md", "docs/DESIGN.md", "CHANGELOG.md"]) {
      expect(readme, `${doc} is not linked from the README`).toContain(doc);
    }
  });

  it("does not promise work the backlog is not carrying", () => {
    /**
     * A table cell or a bullet in the manual may say `planned` on its own only
     * while the backlog's open section actually holds an entry. It says nothing
     * about WHICH entry, because a gate that tried to pair them would be
     * matching prose against prose.
     */
    const heading = "## Open";
    expect(backlog, "the backlog's open section is not where this test looks").toContain(heading);
    const open = backlog.slice(backlog.indexOf(heading) + heading.length);
    const openBody = open.slice(0, open.indexOf("\n## "));
    const openEntries = [...openBody.matchAll(/^### /gm)].length;

    // Lookahead, not a consumed closing pipe: `| a | b |` shares its pipes, so
    // a consuming match swallows the separator and reads every OTHER cell.
    const cells = [...manual.matchAll(/\|([^|\n]+)(?=\|)/g)].map((m) => (m[1] ?? "").trim());
    expect(cells.length, "no table cells found in the manual").toBeGreaterThan(4);
    const bullets = [...manual.matchAll(/^\s*-\s+.*?\*[Pp]lanned\.?\*/gm)].map(() => "planned");
    const promises = [...cells.filter((c) => c.replace(/\*/g, "").toLowerCase() === "planned"), ...bullets];

    if (openEntries === 0) {
      expect(promises, `the manual marks ${promises.length} row(s) "planned" while the backlog is empty`).toEqual([]);
    } else {
      // The other direction: an open backlog with a manual that promises
      // nothing is a manual that has stopped describing the design.
      expect(promises.length, "the backlog is open and the manual marks nothing planned").toBeGreaterThan(0);
    }
  });

  it("shows every picture it links to, and links to every picture it ships", () => {
    // Both directions, once there are pictures at all: a picture referenced and
    // never committed draws a broken image on the page somebody reads first,
    // and a picture committed and never referenced is a file nobody will know
    // to delete or re-render.
    const linked = [...manual.matchAll(/\]\(images\/([^)]+)\)/g)].map((m) => m[1] ?? "");
    const shipped = existsSync("docs/images") ? readdirSync("docs/images") : [];
    for (const name of linked) {
      expect(shipped, `the manual links to docs/images/${name}, which is not committed`).toContain(name);
    }
    for (const name of shipped) {
      expect(linked, `docs/images/${name} is committed and nothing links to it`).toContain(name);
    }
  });

  it("has a heading behind every link inside the manual", () => {
    // A renamed section leaves a link that scrolls nowhere — silently, because
    // Markdown renders a dead anchor as an ordinary link.
    const anchors = new Set(
      [...manual.matchAll(/^#{1,4}\s+(.*)$/gm)].map((m) =>
        (m[1] ?? "")
          .trim()
          .toLowerCase()
          .replace(/[^\w\s-]/g, "")
          .replace(/\s+/g, "-"),
      ),
    );
    const links = [...manual.matchAll(/\]\(#([^)]+)\)/g)].map((m) => m[1] ?? "");
    expect(links.length, "the manual has no internal links").toBeGreaterThan(5);
    for (const link of links) {
      expect(anchors, `the manual links to #${link}, which is not a heading in it`).toContain(link);
    }
  });
});

describe("the page a visitor actually lands on", () => {
  /**
   * `public/index.html` is served from the production origin — it is what the
   * custom domain resolves to, and the first thing anybody sees. A sibling's
   * said "in development" three days after its first release. Staleness itself
   * cannot be caught by a test; what CAN be held is that the page offers a way
   * in at all.
   */
  const page = readFileSync("public/index.html", "utf8");

  it("says where to get the add-in", () => {
    expect(page, "the landing page names no way to install it").toContain("releases/latest");
    expect(page).toContain("manifest-prod.xml");
  });

  it("does not call it in development", () => {
    // The specific phrase that was wrong on the sibling, so re-adding it is
    // deliberate. "Not yet released" is the honest spelling before a release.
    expect(page.toLowerCase()).not.toContain("in development");
  });

  it("gives a link somewhere to unfurl into, with absolute URLs", () => {
    for (const tag of ["og:title", "og:description", "og:url", "og:image"]) {
      expect(page, `no ${tag}`).toContain(tag);
    }
    for (const m of page.matchAll(/property="og:(?:url|image)" content="([^"]*)"/g)) {
      expect(m[1], "an og URL that is not absolute").toMatch(/^https:\/\//);
    }
  });

  it("has an icon, so the tab is not blank", () => {
    expect(page).toMatch(/rel="icon"/);
  });

  it("declares its language, like the task pane does", () => {
    // WCAG 3.1.1.
    expect(page).toMatch(/<html lang="[a-z]{2}"/);
  });
});
