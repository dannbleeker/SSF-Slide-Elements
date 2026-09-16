import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";
import { makeDeck } from "./fixtures/deck.js";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { packageProblems } from "../scripts/package-integrity.mjs";
// @ts-expect-error — plain .mjs with no types, shared with the scripts.
import { PROBE_AUTHOR, anonymise, identityProblems } from "../scripts/deck-identity.mjs";
import { Pkg } from "../src/core/index.js";

/**
 * The deck the probe's question 4 needs.
 *
 * Question 4 asks what each read of the deck DROPS, and the only way to see a
 * drop is to hand the host something droppable. The round of 2026-09-14 was run
 * on the validators' deck, which carries no comment and therefore no
 * `ppt/authors.xml`, so the reader could only say NOT ASKED — a whole round
 * that could not reach the one question it was run for.
 *
 * So: a deck PowerPoint authored over COM on 2026-09-14, windowless, with a
 * real modern comment on slide 3, scrubbed of the account PowerPoint stamped
 * into it (`scripts/deck-identity.mjs`) and checked back through PowerPoint
 * afterwards — opened over COM, 3 slides, the comment still there and
 * attributed to the invented author.
 *
 * Every claim about it is held here, because a binary is the one file a
 * reviewer cannot read.
 */
const DECK = "template/probe-comments.pptx";

let pkg: Pkg;
let parts: Map<string, string | Uint8Array>;
let text: Record<string, string>;

beforeAll(async () => {
  const bytes = new Uint8Array(readFileSync(DECK));
  pkg = await Pkg.open(bytes);
  const zip = await JSZip.loadAsync(readFileSync(DECK));
  parts = new Map();
  text = {};
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    const isText = name.endsWith(".xml") || name.endsWith(".rels");
    parts.set(name, isText ? await file.async("string") : await file.async("uint8array"));
    if (isText) text[name] = await file.async("string");
  }
});

describe("the probe's comment deck", () => {
  it("carries the comment part and the author list, which is the whole point of it", async () => {
    const names = [...parts.keys()];
    const comments = names.filter((name) => name.startsWith("ppt/comments/"));
    expect(comments.length, "a deck with nothing to drop cannot answer question 4").toBe(1);
    expect(names, "the part SSF-Merge's sixth sheet found the export dropping").toContain("ppt/authors.xml");
    const only = comments[0] ?? "";
    // A modern comment, which is the shape the export was measured dropping —
    // not the classic `ppt/comments/comment1.xml` of PowerPoint 2016.
    expect(only).toMatch(/modernComment/);
    const body = await pkg.text(only);
    expect(body).toContain("p188:cmLst");
  });

  it("relates the comment from a slide, so the host has a reason to carry it", async () => {
    const slides = await pkg.slidePaths();
    expect(slides.length).toBe(3);
    const rels = await pkg.text("ppt/slides/_rels/slide3.xml.rels");
    expect(rels).toContain("ppt/comments/".replace("ppt/", "../"));
    expect(rels).toContain("/relationships/comments");
  });

  it("names the invented author and nobody real, because this repository is public", () => {
    // The gate that matters: PowerPoint stamps the signed-in account into both
    // `ppt/authors.xml` and `docProps/core.xml`, and on 2026-09-14 it stamped a
    // display name, initials, a Windows Live `userId` and the provider that
    // issued it. Re-authoring this deck and committing it unscrubbed is a red
    // test here rather than a published account id.
    expect(identityProblems(text)).toEqual([]);
    expect(text["ppt/authors.xml"]).toContain(`name="${PROBE_AUTHOR.name}"`);
  });

  it("is widescreen, so the pane borrows no library when it is open", async () => {
    const pres = await pkg.text("ppt/presentation.xml");
    expect(pres).toContain('cx="12192000"');
    expect(pres).toContain('cy="6858000"');
  });

  it("is a package nothing objects to, and small enough to upload anywhere", () => {
    expect(packageProblems(parts)).toEqual([]);
    expect(statSync(DECK).size).toBeLessThan(200_000);
  });
});

describe("who a committed deck says wrote it", () => {
  /**
   * The rules, each against a package that breaks exactly one of them. A check
   * that only ever sees the committed deck passes for as long as that deck is
   * untouched, which is the same as not being a check at all.
   */
  const clean = {
    "ppt/authors.xml": `<p188:authorLst><p188:author id="{A}" name="${PROBE_AUTHOR.name}" initials="${PROBE_AUTHOR.initials}" userId="${PROBE_AUTHOR.userId}" providerId="${PROBE_AUTHOR.providerId}"/></p188:authorLst>`,
    "docProps/core.xml": `<cp:coreProperties><dc:creator>${PROBE_AUTHOR.name}</dc:creator><cp:lastModifiedBy>${PROBE_AUTHOR.name}</cp:lastModifiedBy></cp:coreProperties>`,
  };

  it("passes a package that names only the invented author", () => {
    expect(identityProblems(clean)).toEqual([]);
  });

  it("catches a display name", () => {
    const stamped = {
      ...clean,
      "ppt/authors.xml": clean["ppt/authors.xml"].replace(PROBE_AUTHOR.name, "A Real Person"),
    };
    expect(identityProblems(stamped)).toEqual([
      `ppt/authors.xml gives name="A Real Person", not "${PROBE_AUTHOR.name}"`,
    ]);
  });

  it("catches the account id, which is the part that is not a name and still identifies", () => {
    const stamped = {
      ...clean,
      "ppt/authors.xml": clean["ppt/authors.xml"].replace(PROBE_AUTHOR.userId, "911d2b9d99de05ee"),
    };
    expect(identityProblems(stamped)).toEqual([
      `ppt/authors.xml gives userId="911d2b9d99de05ee", not "${PROBE_AUTHOR.userId}"`,
    ]);
  });

  it("catches the identity provider", () => {
    const stamped = {
      ...clean,
      "ppt/authors.xml": clean["ppt/authors.xml"].replace(PROBE_AUTHOR.providerId, "Windows Live"),
    };
    expect(identityProblems(stamped)).toEqual([
      `ppt/authors.xml gives providerId="Windows Live", not "${PROBE_AUTHOR.providerId}"`,
    ]);
  });

  it("catches both document properties, not just the first", () => {
    const stamped = {
      ...clean,
      "docProps/core.xml": `<cp:coreProperties><dc:creator>Someone</dc:creator><cp:lastModifiedBy>Someone Else</cp:lastModifiedBy></cp:coreProperties>`,
    };
    expect(identityProblems(stamped)).toEqual([
      `docProps/core.xml gives dc:creator as "Someone", not "${PROBE_AUTHOR.name}"`,
      `docProps/core.xml gives cp:lastModifiedBy as "Someone Else", not "${PROBE_AUTHOR.name}"`,
    ]);
  });

  it("catches an address in a part it has no rule for", () => {
    // The allow list only knows three parts. A deck that carried a person's
    // address in a slide's own text would pass every rule above, which is why
    // there is a second net under them.
    const stamped = { ...clean, "ppt/slides/slide1.xml": "<a:t>write to someone@example.com</a:t>" };
    expect(identityProblems(stamped)).toEqual([
      "ppt/slides/slide1.xml carries what reads as an address, someone@example.com",
    ]);
  });

  it("catches the organisation, which the allow list used to have no rule for at all", () => {
    // `docProps/app.xml` was outside the allow list until 2026-09-16. The
    // 2026-09-14 reading that built the list found the owner in two parts and
    // no others — on ONE machine, whose Office carries no company name. An
    // Office that has one writes it into every file it saves, and the check
    // would have stayed green on exactly the machine where it mattered.
    const stamped = {
      ...clean,
      "docProps/app.xml": `<Properties><Application>Microsoft Office PowerPoint</Application><Company>A Real Employer A/S</Company><Manager>A Real Manager</Manager></Properties>`,
    };
    expect(identityProblems(stamped)).toEqual([
      `docProps/app.xml gives Company as "A Real Employer A/S", and a committed deck names no organisation`,
      `docProps/app.xml gives Manager as "A Real Manager", and a committed deck names no organisation`,
    ]);
    // Emptied, not filled in with a stand-in, and nothing else in the part moves.
    const cleaned = anonymise(stamped);
    expect(identityProblems(cleaned)).toEqual([]);
    expect(cleaned["docProps/app.xml"]).toContain("<Application>Microsoft Office PowerPoint</Application>");
    expect(cleaned["docProps/app.xml"]).toContain("<Company></Company>");
  });

  it("says nothing about an app part that names no organisation, which is most of them", () => {
    // Absent and empty are both fine; only a value is a problem. A rule that
    // confused them would fail every ordinary deck in `template/`.
    expect(
      identityProblems({
        ...clean,
        "docProps/app.xml": `<Properties><Company></Company><Slides>3</Slides></Properties>`,
      }),
    ).toEqual([]);
  });

  it("says nothing about a deck with no comment in it, which has no author list at all", () => {
    // `ppt/authors.xml` exists only once there is a modern comment. Absent is
    // not the same as wrong, and a rule that confused the two would fail every
    // ordinary deck in `template/`.
    expect(identityProblems({ "docProps/core.xml": clean["docProps/core.xml"] })).toEqual([]);
    expect(identityProblems({})).toEqual([]);
  });

  it("rewrites exactly what it complains about, and leaves the rest alone", () => {
    const stamped = {
      "ppt/authors.xml": `<p188:authorLst><p188:author id="{A}" name="A Real Person" initials="ARP" userId="911d2b9d99de05ee" providerId="Windows Live"/></p188:authorLst>`,
      "docProps/core.xml": `<cp:coreProperties><dc:creator>A Real Person</dc:creator><cp:lastModifiedBy>A Real Person</cp:lastModifiedBy></cp:coreProperties>`,
      "ppt/slides/slide1.xml": "<a:t>untouched</a:t>",
    };
    const cleaned = anonymise(stamped);
    expect(identityProblems(cleaned)).toEqual([]);
    expect(cleaned["ppt/slides/slide1.xml"]).toBe(stamped["ppt/slides/slide1.xml"]);
    // The author's own id is a GUID this deck generated, not an account, and it
    // is what the comment's `authorId` points at — rewriting it would orphan
    // the comment.
    expect(cleaned["ppt/authors.xml"]).toContain('id="{A}"');
  });
});

describe("the scrubber's own command line", () => {
  /**
   * A scrub it cannot finish must not leave a file behind.
   *
   * `anonymise` only knows the parts the allow list names, so a package can come
   * out of it still naming somebody — through the address net, or a field
   * nothing here has a rule for. The CLI used to write the output file and THEN
   * say so, which leaves a deck that looks scrubbed on disk, under the name the
   * operator chose for the scrubbed copy, with nothing but a line of console
   * output and an exit code against it. The next person to reach for that file
   * reaches for a file that names a person.
   *
   * Run as a subprocess because that is the surface under test: `main` is what
   * the exit code comes from, and calling the pure functions would not see it.
   */
  it("refuses to write a copy it could not finish scrubbing", async () => {
    const dir = mkdtempSync(join(tmpdir(), "ssf-identity-"));
    const input = join(dir, "dirty.pptx");
    const output = join(dir, "scrubbed.pptx");
    // An address in a slide's own text: past every allow-list rule, caught by
    // the second net, and not something `anonymise` can rewrite.
    writeFileSync(input, await makeDeck([{ paragraphs: [["write to someone@example.com"]] }]));

    const run = spawnSync(process.execPath, ["scripts/deck-identity.mjs", input, output], { encoding: "utf8" });

    expect(run.status, run.stdout).toBe(1);
    expect(run.stdout).toContain("NOT written");
    expect(existsSync(output), "a deck that still names somebody was left on disk").toBe(false);
    rmSync(dir, { recursive: true, force: true });
  });

  it("writes the copy when the scrub does reach everything", async () => {
    // The other half, so the refusal above is a refusal and not a CLI that
    // never writes anything.
    const dir = mkdtempSync(join(tmpdir(), "ssf-identity-"));
    const input = join(dir, "ordinary.pptx");
    const output = join(dir, "scrubbed.pptx");
    writeFileSync(input, await makeDeck([{ paragraphs: [["nothing identifying here"]] }]));

    const run = spawnSync(process.execPath, ["scripts/deck-identity.mjs", input, output], { encoding: "utf8" });

    expect(run.status, run.stdout).toBe(0);
    expect(existsSync(output)).toBe(true);
    rmSync(dir, { recursive: true, force: true });
  });
});
