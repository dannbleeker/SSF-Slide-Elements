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

**Today the scaffold and the package layer exist.** Nothing reads a deck from
PowerPoint yet and nothing splices; `docs/BACKLOG.md` is the order the rest
arrives in.

## Where things live

| directory | what it owns |
| --- | --- |
| `src/core/` | the engine, pure: `pptx/` is the package layer (`Pkg` over a .pptx as parts, relationships, content types and the slide list; `xml.ts`, `parts.ts`). The harvest and the splice arrive next |
| `src/host/` | the DECISIONS about talking to a host, all pure and all tested: `capability.ts` (the version floor), `errors.ts` (a raise as a bounded sentence) |
| `src/office/` | the Office.js CALLS, and nothing else. Every judgement is imported from `src/host` |
| `src/pane/` | `steps.ts` (which step, what the one button says, why it is blocked), `render.ts` (the DOM), `main.ts` (**the only file here allowed to touch Office.js**), plus the HTML and the SSF stylesheet |
| `scripts/` | the manifest generator and its rules, the icon drawer, the test-count floor, the release pre-flight, the pane audit, the sibling sweep and its `TRIAGED` table |
| `public/` | copied verbatim into `dist/`: the CNAME, the landing page, the support and privacy pages the manifests point at, the icons |

**`src/host` decides, `src/office` calls, and the architecture test holds both
directions.** An Office.js import in `src/host` makes a rule untestable; a rule
reimplemented inline in `src/office` looks tidier and rots quietly, so every
file there must import from `src/host` or `src/core`.

`src/office` is **not** in the coverage include list (`scripts/coverage-scope.mjs`),
deliberately. Pooling a well-tested engine with untestable host calls produces
one number that hides both.

**The pane is the one surface the suite cannot judge.** `pane-render.test.ts`
pins its behaviour in jsdom, which has no layout and no colour. `npm run
pane-shots` renders every state at 320 and 512 in both themes and MEASURES
overflow, contrast, focus rings, hit areas and axe; a weekly job runs it and
looking at the output is part of done for any change to the pane. The layout
rules: one orange element per view, one column, one primary control drawn last.

## What THIS host answered

Nothing yet. Every host fact in this repo is borrowed from a sibling, and the
first round against a real PowerPoint is the last item on the backlog. Until it
runs, nothing here should imply the host has been measured.

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
- **Without `targetSlideId`, `insertSlidesFromBase64` inserts at the FRONT.** A
  real run put 37 generated slides ahead of a title slide.
- **A queued call that raises nothing has not necessarily happened.** Confirm
  every destructive step with a second read; the deck DELTA is the evidence,
  never the absence of an error.
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

- **A gate that cannot fail is not a gate.** Ask of each one "what would I
  break to make this go red?", then break it. Four of a sibling's could not.

- **A number copied from a live counter carries the date it was taken.**
  Otherwise it is a claim that rots. Never pair figures from two measurements.

- **Say what is measured and what is assumed**, in comments, docs and PRs alike.
  A sentence written here becomes something the next reader builds on.

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
  line in `docs/SIBLING.md` (`test/sibling.test.ts` holds that). Every
  RELEVANT or ADOPTED row today says "re-triage when the insert lands", because
  there is no host code to hold it against; the host-handshake PR re-verdicts
  them.

- **Values never leave the pane. All sample data is invented. The repo is
  public.**

## Commands

```bash
npm test               # the whole suite
npm run typecheck
npm run lint
npm run coverage       # enforces the floors in vitest.config.ts
npm run test:count     # holds the floor in test/fixtures/test-count.json
npm run manifests      # regenerate the four manifests; test/manifest.test.ts diffs them
npm run icons          # redraw public/assets; test/manifest.test.ts diffs them
npm run sibling-watch  # sweep both siblings' tables for findings with no row in TRIAGED
npm run build          # the site, for GitHub Pages
npm run pane-shots     # needs `npx vite --port 5199 --strictPort &` first
```

## Open questions for the real host

Nothing should be built on a guess about any of them; each is written so a
single round settles it.

1. Does `insertSlidesFromBase64` accept a package pruned to one slide whose
   other parts are still present but unlisted? OPC permits it; whether
   PowerPoint agrees on the way in is not measured.
2. Does an insert immediately followed by a positional delete of the slide
   before it keep the order the engine expects?
3. Does `getSelectedSlides()` name the slide the user is looking at, and does
   its position in `slides` match the position in `<p:sldIdLst>`?
4. Which read of the deck this add-in should use — `getFileAsync` or
   `exportAsBase64Presentation` — and what each drops on this host.
