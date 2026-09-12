# SSF Slide Elements — project memory

A library of ready-made slide elements for PowerPoint, as an Office.js task pane
add-in. Part of the SSF family: SSF-Merge is the model this repo is built to,
and the source of the package route; SSF-Charts is the record of what
shape-by-shape drawing costs on this host, and of most of what the family knows
about it. Both are public; read them with this one checked out beside them.

## Architecture in one paragraph

A .pptx is a zip of XML parts. The insert will happen **in the file**: read the
open presentation, splice an element's markup into the slide the user is on,
reduce the package to that one slide, hand it to PowerPoint through one
`insertSlidesFromBase64`, and remove the slide it replaced — insert first, prove
the deck grew, then remove, so the failure mode is a duplicate the user can
delete rather than a lost slide. Nothing is drawn shape by shape, so none of
SSF-Charts' per-shape failure surfaces exist. `src/core` is pure TypeScript with
**zero Office imports**, enforced by `test/architecture.test.ts` in both
directions.

**The whole path exists and has been run against a real PowerPoint.** The two
library decks under `template/` are read into `public/catalogue/` by
`npm run harvest` (the index committed, the rest built on deploy); the pane
reads the open deck, splices an element into a copy of the slide you are on,
hands it back in one insert and removes the original by position. Every tile
carries PowerPoint's own rendering of its element, cut from the committed PDF
print of the deck it came from (`template/library-*.pdf`, with the geometry in
`*.print.json`); the landing diagram is the fallback for an element the cut
could not reach. What is NOT built is in `docs/BACKLOG.md`, and none of it is
code this repo can write on its own: a round that exercises "Remove from N
slides", rounds on Mac and iPad, and the release — a screenshot, the listing
name and the Partner Center submission.

## Where things live

| directory | what it owns |
| --- | --- |
| `src/core/` | the engine, pure: `pptx/` is the package layer (`Pkg` over a .pptx as parts, relationships, content types and the slide list; `xml.ts`, `parts.ts`); `catalogue/` is the harvest (a deck into elements: headings, the collection marker, boxes with rotation and table columns, size runs, tags, carried parts, and `cut.ts`, where on the deck's PDF print each element's picture is); `splice/` puts an element into a copy of a slide and takes one back out |
| `src/host/` | the DECISIONS about talking to a host, all pure and all tested: `capability.ts` (the version floor), `coalesce.ts` (one selection read at a time), `errors.ts` (a raise as a bounded sentence), `insert.ts` (what a measured delta means, and the undo plan), `jump.ts` (whether the host was seen on the slide), `links.ts` (what may reach a URL), `memory.ts` (which storage bucket a deck remembers itself in), `probe.ts` (what each probe observation means), `timeout.ts` (every budget, and the backoff a lagging count needs) |
| `src/office/` | the Office.js CALLS, and nothing else. Every judgement is imported from `src/host` |
| `src/pane/` | `steps.ts` (which step, what the one button says, why it is blocked), `render.ts` (the DOM), `catalogue.ts` (the index and an element's markup, fetched from the site), `main.ts` (**the only file here allowed to touch Office.js**), plus the HTML and the SSF stylesheet |
| `scripts/` | the manifest generator and its rules, the icon drawer, the test-count floor, the release pre-flight, the pane audit, the sibling sweep and its `TRIAGED` table, the harvest, the probe builder (`build-probe.mjs`, `probe-fixture.mjs`) and the answer reader (`read-answers.mjs`) |
| `probe/` | `probe-snippet.ts`, GENERATED for Script Lab and committed; CI rebuilds and diffs it. Pasted into PowerPoint by the owner, never imported here |
| `docs/PROBE.md`, `docs/host-answers/` | how to run the probe, and every answer sheet it has produced, stamped |
| `public/` | copied verbatim into `dist/`: the CNAME, the landing page, the support and privacy pages the manifests point at, the icons |
| `template/` | the owner's library: the two decks (one per slide size), the PDF print of each and the `*.print.json` saying where every element sits on it, `names.en.json` (the English name of every element keyed by the deck's Danish title), and `validators.pptx`, the deck AppSource's reviewers are given |
| `docs/DESIGN.md` | the design record: every decision, dated, and the seven host questions the probe asks |

**`src/host` decides, `src/office` calls, and the architecture test holds both
directions.** An Office.js import in `src/host` makes a rule untestable; a rule
reimplemented inline in `src/office` looks tidier and rots quietly, so every
file there must import from `src/host` or `src/core`.

`src/office` is **not** in the coverage include list (`scripts/coverage-scope.mjs`),
deliberately. Pooling a well-tested engine with untestable host calls produces
one number that hides both.

**The pane is the one surface the suite cannot judge.** `pane-render.test.ts`
pins its behaviour in jsdom, which has no layout and no colour. `npm run
pane-shots` renders every state at 320 and 512 in both themes, and once more
at 320 in FORCED COLOURS, and MEASURES overflow, contrast, focus rings, hit
areas, decorations that vanish when colours are forced, and axe; a weekly job
runs it and looking at the output is part of done for any change to the pane. The layout
rules: one orange element per view, one column, one primary control drawn last.

## What THIS host answered

**PowerPoint for the web and PowerPoint on Windows.** Six answer sheets under
`docs/host-answers/` — four from the web on 2026-09-10, a Windows pair on
2026-09-11; `docs/DESIGN.md` section 15 reads them and is the one place to
change when a sheet is filed. The splice ran against the web on 2026-09-10, and
the whole product — pane, insert and Undo — against the web and then against
Windows on 2026-09-11. **Mac and iPad have had no round**, so every claim about
those two is still borrowed and must say so.

**Windows answered every question the way the web did**, and every timing
difference went the same direction: faster. Two of them matter, because the
code carries a workaround for each and neither workaround is load-bearing on
Windows — they stay because the web still needs them:

- **The slide count does NOT lag an insert on Windows.** Measured twice on
  2026-09-11, polling `slides.getCount()` every 300 ms through an
  `insertSlidesFromBase64` whose promise was timestamped: the new count was
  already being returned 240 ms and 247 ms BEFORE the call resolved. The web
  sat at the old value for 2.8 seconds. `countReaching` costs nothing here and
  is not evidence about Windows.
- **A selection read straight after an insert does not hang on Windows.** 3 ms,
  against a four-second budget in the pane. The web could take seconds.

Two rules about the INSTRUMENT, both learned on 2026-09-11, when a round that
should have settled the undo defect settled nothing:

- **Office.js lives in the add-in's frame, not the host's.** A `PowerPoint.run`
  evaluated inside PowerPoint for the web's own editor frame (`ppt.aspx`) does
  not run, and a driver that catches the failure reports "could not read the
  deck" — which reads as a fact about the host and is a fact about the driver.
  Read the deck from the pane's frame, where Office.js actually is. Three
  readings came back as `-1` this way and were nearly written up as a host
  limitation.
- **Insert something the slide does not already carry.** The undo was checked by
  inserting the same element a previous broken round had left on the slide, so a
  working undo and a broken one produced the same picture and the round proved
  nothing either way. The round that settled it put a TRIANGLE onto a slide
  holding a title and a white box, and compared the shape inventory before and
  after, id by id and name by name. Choose the payload so the two hypotheses
  cannot look alike.

## Host rules, learned the expensive way

These come from the siblings' rounds against PowerPoint on the web. They are
recordings, not opinions. Do not re-derive them, and do not build on a guess
where one of them applies. `docs/SIBLING.md` is the ledger: which finding each
rule came from, what was done about it here, and the rule that keeps a borrowed
counter dated.

- **A slide the run just added does not resolve by id.** So metadata goes into
  the package before the insert, and undo is **positional** with clamps, never
  by id.
- **Tag writes through a shape proxy are refused.** Writing `ppt/tags/tagN.xml`
  cannot be refused, because nothing is asked. Find the next free `tagN`.
- **Tags written into the package SURVIVE the insert.** Measured here on the
  web, 2026-09-11: the deck read back after a session of inserts carries seven
  `SSF_SLIDE_ELEMENT` parts, referenced from the shapes through
  `<p:custDataLst>` on four slides. The same deck carries think-cell's tags in
  the same folder, which is exactly why the next free `tagN` is found rather
  than assumed.
- **Without `targetSlideId`, `insertSlidesFromBase64` inserts at the FRONT.** A
  real run put 37 generated slides ahead of a title slide.
- **A queued call that raises nothing has not necessarily happened.** Confirm
  every destructive step with a second read; the deck DELTA is the evidence,
  never the absence of an error.
- **The deck's own size can lag a call that has already happened.** Measured
  HERE, on the web, 2026-09-11: polling `slides.getCount()` every 300 ms
  through a real insert, the count stayed at its old value for 2.8 seconds and
  then went up. So the delta is the evidence and ONE read of it is not: a size
  that decides anything is asked again, backed off, until the deck agrees or the
  backoff runs out (`countReaching`). Re-reading the measurement, never
  re-trying the call.
- **A call can raise and still have done the work.** SSF-Merge's insert timed
  out with both slides landed. Read the delta.
- **An empty collection read is not an empty slide.** Never claim data loss
  from a read; an id refusal anywhere makes every id in that operation suspect.
- **Two `slides.add()` calls 0.4s apart killed the tab.** Slides arrive through
  one insert, never a loop of adds. **Do not wait after adding a slide** either.
- **Probe for the method, do not trust the requirement set.** The floor is
  checked at runtime by `checkFloor` and never declared in the manifest.
- **`load("items")` does not load the items' properties.** Name them:
  `load("items/id")`. Collection loads over ~50 items can answer short.
- **Nothing calls `setSelectedShapes`.** It wedges the web host's selection
  subsystem. `getSelectedSlides` is read-only and measured safe.
  `setSelectedSlides` is called by ONE thing, the jump in "Used in this deck"
  (`src/host/jump.ts`), on SSF-Charts' dated web measurement (2,429 rungs,
  2026-08-13 to 2026-09-04, none silent), never on desktop evidence — so the
  pane reads the selection back after every call and claims only what it saw.
  Probe question 7 measures the call; until a sheet answers it, the jump is
  borrowed, not measured, on every platform.
- **Shape tags do not survive cut/paste on the web.** Say so in the docs; do
  not try to detect it.
- **A custom XML part written at the package root is invisible to Office.js.**
  Relate it from `ppt/presentation.xml`.
- **The two ways to read a deck return DIFFERENT things.** `getFileAsync` (the
  floor) hands back the USER'S ENTIRE PRESENTATION;
  `exportAsBase64Presentation` (1.10) hands back only the slides asked for and
  DROPS comments. Whichever this add-in uses, the package it sends back must be
  measured, not assumed.
- **Blob downloads from the task pane are blocked in WebView2.** Anything a
  maintainer needs has to survive being copied out by hand.

## The lockstep rule (only item 1 is CI-enforced)

Any feature change updates, in the same PR:

1. **`docs/MANUAL.md`** — CI-enforced. `test/docs.test.ts` fails on a step
   title or a button label the manual does not quote, on a pane section that
   calls built work planned, and on a manual that promises `planned` work the
   backlog is not carrying.
2. **`CHANGELOG.md`** — under `## [Unreleased]`, in the language a user would
   use, not the language of the diff.
3. **`docs/BACKLOG.md`** — an item that shipped is REMOVED, not ticked. A
   rejected idea moves to the rejected list with the reason.
4. **`README.md`** feature table.

Items 2 to 4 are on you. Nothing in a diff distinguishes a feature change from
a refactor, and a check that guessed would be noise.

## Conventions

- **Answer in caveman style** (the `caveman` skill, `full` level) from the first
  reply of every session in this repo, and for every reply after it. It is on by
  default here, it does not lapse because a turn is long, and it does not need
  re-requesting. Only "stop caveman" turns it off, for that session.

  Code, commits, PR bodies and user-facing docs stay normal prose — they have
  readers who were not in the conversation. Drop it for security warnings,
  destructive-action confirmations, and any multi-step sequence where dropping
  conjunctions could be misread. Resume straight after.

- **Merging to `main` is authorized.** Granted by the owner on 2026-09-08 for
  this repo. Once CI is green on the exact pushed commit — verify the run's
  `head_sha` matches the branch head, because a force-push leaves older runs
  attached to the PR and they are not evidence about the head — squash-merge
  without asking. It does not extend to merging somebody else's PR, to changing
  what `main` requires, or to a red or conflicted head.

- **Branch flow**: develop on the session's designated `claude/*` branch, one PR
  per increment. `main` requires a pull request and a passing `test` check, and
  the branch is deleted on merge, so the next increment starts with
  `git checkout -B <branch> origin/main` rather than reusing what is there.
  Reusing a branch across a squash merge is what forces a force-push. The one
  force-push of this branch, to discard the closed PR #1's commit, was
  authorized by the owner on 2026-09-08 and is not a habit here.

- **A regression test must be proven to fail without its fix.** Break the
  source, re-run, confirm the new test goes red for the RIGHT reason, restore.
  Check WHICH assertion goes red, not just that one does.

- **A guard that goes red for the wrong reason is worse than no guard.** The
  siblings' no-Office-imports test first matched the word "Office.js" in the
  comments explaining why the engine avoids it. `scripts/without-prose.mjs`
  exists because the same mistake was made three times in three syntaxes.

- **The design record gets ahead of the code, and nothing used to notice.**
  Four behaviours approved in `docs/DESIGN.md` sections 4 and 6 — the jump,
  "Open all", "Move to a new slide", the pane's per-deck memory — sat described
  and unbuilt until they were found by reading the record against the pane on
  2026-09-12. `paneControlProblems` in `scripts/doc-refs.mjs` now holds every
  control label the record QUOTES against what the pane draws. It is a floor,
  not a proof: a behaviour the record describes without quoting a label for it
  still slips through, so reading the two side by side after a pane increment
  is still the practice.

- **A gate that cannot fail is not a gate.** Ask of each one "what would I
  break to make this go red?", then break it. Four of a sibling's could not.

- **A number copied from a live counter carries the date it was taken.**
  Otherwise it is a claim that rots. Never pair figures from two measurements.

- **Say what is measured and what is assumed**, in comments, docs and PRs alike.
  A sentence written here becomes something the next reader builds on.

- **A design change updates `docs/DESIGN.md` in the same PR.** The record is
  what the build is held to; a behaviour that is not in it is a guess, and a
  behaviour that contradicts it is a bug in one of the two.

- **Test files are named by topic, never by increment.**

- **Every npm script stays FLAT.** A script that nests `npm run` is blocked by
  AppLocker on the owner's Windows box, so the gates are listed step by step in
  the workflows and `test/release.test.ts` holds the two lists equal.

- **Run the WHOLE gate as the last thing before a commit, not before the last
  edit.** `test-count` reads the working tree, so do not run it beside work in
  progress. `git add -A` commits whatever else is touching the tree; stage paths
  when anything else is running, and check `git status` before a commit you did
  not build file by file.

- **Flag manifest re-installs to the owner.** The add-in is hosted on Pages, so
  code, pane and catalogue changes ship through `main` with **no** re-install.
  A change to the manifest itself — ribbon buttons, `Permissions`, requirement
  sets, `DisplayName`, icon references, or (never) the `<Id>` GUID — needs the
  owner to re-sideload. Say so explicitly in that PR/turn: "⚠️ this needs a
  manifest re-install in PowerPoint", and bump `VERSION` in
  `scripts/manifest-source.mjs`, never the npm version.

- **Dependabot's banner gets read, and the reading gets written down** in
  `docs/DEPENDENCY-ALERTS.md`, "no exposure" included. Never take
  `npm audit`'s advice unread: on a sibling its proposed remedy moved a runtime
  dependency back three major versions.

- **Sibling watch.** `.github/workflows/sibling-watch.yml` sweeps both siblings'
  curated tables every Monday and files one issue for any finding without a row
  in `TRIAGED` (`scripts/sibling-watch.mjs`). A finding without a row comes
  back; "no exposure" is a real answer, and every row that is NOT one has a
  line in `docs/SIBLING.md` (`test/sibling.test.ts` holds that). Every RELEVANT
  or ADOPTED row was re-verdicted on 2026-09-11 **against the shipped insert**,
  and says so with that date — the sentence they used to carry, "re-triage when
  the insert lands", was written when there was no host code to hold them
  against. The removal was held against them in the same change, as a second
  caller of the positional delete. What came after and has now been read
  against them too, on 2026-09-12: the tag work of #48, #49 and #52, which made
  two features depend on finding a tag the web loses on cut and paste, and the
  shared-part carry of #54. That pass also found three rows asserting something
  false about this add-in — that it never loads a shape collection, when
  `selectedShape` does — and two that disagreed about how many tags an insert
  writes. Re-verdict a row when the code it is about changes, not when a
  calendar says so, and check the claim rather than the date.

- **Values never leave the pane. All sample data is invented. The repo is
  public.**

## Commands

```bash
npm test               # the whole suite
npm run typecheck
npm run lint
npm run coverage       # enforces the floors in vitest.config.ts
npm run test:count     # holds the floor in test/fixtures/test-count.json
npm run dead-exports   # every export the shipped add-in never calls
npm run harvest        # both decks into public/catalogue; CI diffs the committed index
npm run previews       # cut every element's picture out of its deck's PDF print; runs on deploy
npm run print-stamp    # re-stamp template/*.print.json after a new print; test/print.test.ts gates it
npm run probe          # regenerate probe/probe-snippet.ts; CI diffs it; needs build:lib first
npm run manifests      # regenerate the four manifests; test/manifest.test.ts diffs them
npm run icons          # redraw public/assets; test/manifest.test.ts diffs them
npm run sibling-watch  # sweep both siblings' tables for findings with no row in TRIAGED
npm run bench          # what the engine costs per insert, printed
npm run release:check  # the release pre-flight, against RELEASE_VERSION
npm run build          # the site, for GitHub Pages
npm run build:lib      # the engine to dist-lib/, which harvest, previews and probe need first
npm run dev            # the pane at localhost:3002
npm run format         # Prettier, on code only
npm run pane-shots     # needs `npx vite --port 5199 --strictPort &` first
```

## Open questions for the real host

Nothing should be built on a guess about any of them; each is written so a
single round settles it, and `probe/probe-snippet.ts` asks all of them
(`docs/PROBE.md` says how each arm is built).

1. Does `insertSlidesFromBase64` accept a package pruned to one slide whose
   other parts are still present but unlisted? OPC permits it; whether
   PowerPoint agrees on the way in is not measured.
2. Does an insert immediately followed by a positional delete of the slide
   before it keep the order the engine expects?
3. Does `getSelectedSlides()` name the slide the user is looking at, and does
   its position in `slides` match the position in `<p:sldIdLst>`?
4. Which read of the deck this add-in should use — `getFileAsync` or
   `exportAsBase64Presentation` — and what each drops on this host.
5. Does PowerPoint's own Ctrl+Z revert `insertSlidesFromBase64`? If it does,
   the pane's Undo must not fight it.
6. How long does `getFileAsync` take on a 50 MB deck, since the file route
   reads the whole deck for every insert, and is the floor met on iPad?
7. Does `setSelectedSlides` move the view, and does the host still answer a
   selection read afterwards? The jump in "Used in this deck" makes that call
   on a sibling's evidence; until a sheet answers this, it is borrowed
   everywhere.
