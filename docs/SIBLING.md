# What we know from the siblings, and how it gets here

Nothing this repo knows about the PowerPoint host was learned here. It was
learned by [SSF-Charts](https://github.com/dannbleeker/SSF-Charts) over its
real-host rounds against PowerPoint on the web, and carried into
[SSF-Merge](https://github.com/dannbleeker/SSF-Merge), which triaged every one
of those findings against its own package route and added three answer sheets
of its own. SSF Slide Elements held no host code at all when this file was
written, on 2026-09-08. It has since run its own probe on PowerPoint for the
web (four sheets under `docs/host-answers/`, 2026-09-10) and its own product
round on the same host (2026-09-11), so some of what follows is now measured
here rather than borrowed, and the rows say which. Everything about Windows,
Mac and iPad is still a recording from one of those two projects.

That is a debt, and it does not settle itself. This file is the single place
sibling-derived knowledge lives: what each finding is, what was done about it
here, and the rules that stop a borrowed fact quietly going bad.

## Why this file exists — the drift, as SSF-Merge measured it

SSF-Merge counted its own citations on 2026-08-27: **44 references to
SSF-Charts across 24 source, script and doc files**, every one hand-copied with
no link back. Sorting them found two kinds, and only one is a problem:

- **A single run's observation is durable.** "A by-id clean-up reported 45
  successful deletes having removed nothing" happened once, in one round, and
  will be true forever. Most citations are this. Leave them alone.
- **A live counter is a snapshot.** Five places quoted the number of archived
  rounds SSF-Charts had run. All five were right that morning, and the next
  round made every one of them wrong, with nothing anywhere to say so.

One had already drifted: a sentence pairing a batch count from one measurement
with a maximum from another, both correct separately and false together.

## The rule

**A number copied from a live counter carries the date it was taken.**

That turns a claim into a recording. "174 of 174" rots the moment another round
runs; "174 of 174, measured 2026-08-27" never becomes false, only older, and a
reader can judge that for themselves.

Two corollaries:

- **Where the number adds nothing, drop it.** "Every rung answered in every
  archived round" is stronger than "in all 174" and cannot decay.
- **Never pair figures from two measurements.** They read as one finding and
  are not.

`test/sibling.test.ts` holds the rule over `src/`, `scripts/` and `test/`: a
round count in a comment with no date within six lines fails the suite. It does
not read the docs, so this file and `CLAUDE.md` may quote the sentences that
break it.

## How a fact gets here

1. Something is learned in a sibling — a round, an upstream issue, a reverted
   fix.
2. It is triaged against **this** add-in's surface. The bar is SSF-Charts' own,
   and it is what keeps a report readable: not "this is interesting about
   Office.js" but **"if this were true, code in this repo would be wrong."**
3. It gets a row in `TRIAGED` (`scripts/sibling-watch.mjs`) — **including when
   the answer is "no exposure"**, which records that somebody checked. An
   untriaged item is indistinguishable from an unnoticed one.

Anything without a row is unanswered, and the sweep brings it back.

### The precondition nobody should have to rediscover

Step 2 needs **both sibling repositories checked out in the same session**. A
session holding only this repo cannot read SSF-Charts' tables, cannot verify a
citation against SSF-Merge's answer sheets, and cannot do any of this by hand.
It can read a sweep's output and nothing else. Say so when starting work that
depends on it.

### Where to read, in order of signal

SSF-Charts' `CLAUDE.md` is the fullest record and the worst feed: thousands of
lines of prose that diff into noise. The curated, keyed surfaces are what the
sweep watches and what a person should read first.

| Repo | Source | What it holds |
| --- | --- | --- |
| SSF-Charts | `scripts/office-js-watch.mjs` → `KNOWN_ISSUES` | Every office-js issue that project has triaged, with what it did about each. |
| SSF-Charts | `scripts/host-baseline.mjs` → `FAKE_BASELINE`, `KNOWN_DIVERGENCES`, `UNSTABLE_ANSWERS`, `PENDING_QUESTIONS` | What it knows about the host, keyed and worded for a diff. The pending register is empty by design between a probe's commit and its answering round. |
| SSF-Charts | `docs/BACKLOG.md`, the rejected list | Decisions not to do things, with reasons. Cheaper to read than to re-derive. |
| SSF-Merge | `scripts/sibling-watch.mjs` → `TRIAGED` | Its verdict on each SSF-Charts finding against the package route, which is this add-in's route too. Watched as a cross-check: a finding it triaged that has no row here is worth a look. |
| SSF-Merge | `CLAUDE.md` host rules; `docs/host-answers/` | Its own answer sheets from PowerPoint on the web, as prose. Not keyed, so not swept; the `merge:` namespace in `TRIAGED` is reserved for the day they are. |
| SSF-Merge | `docs/SIBLING.md` | The ledger this one is modelled on, with the drift measurement above. |

## The ledger

Every row here is a sibling finding this add-in **acted on**: a verdict of
`ADOPTED` or `RELEVANT` in `TRIAGED`. The rows marked `NO EXPOSURE` are
summarised after the table rather than listed, and `test/sibling.test.ts` fails
when a finding with any other verdict is missing from this file.

The rows were written on 2026-09-08, when no host code existed here, and
**re-triaged on 2026-09-11** against the insert that has since shipped. Four of
them had promised a defence the code never built — the deck read was to page
`getItemAt` at twenty and never do a collection load, and it does three
collection loads and no paging — and those four now say what the code actually
does instead. A row that reads as a description of this add-in and is not is
worse than no row at all, because the next reader builds on it.

| Finding | Where it was learned | What this add-in does about it |
| --- | --- | --- |
| A slide add whose sync never resolves though the slide lands (office-js#1650); a slide the run just added does not resolve by id, and a by-id clean-up once reported 45 deletes having removed nothing | SSF-Charts field rounds; SSF-Merge's `undo.ts` | **Adopted as doctrine.** The insert is one `insertSlidesFromBase64`; anything to remember goes into the package before it; the deck DELTA is the evidence, never the absence of an error. |
| A stale shape proxy answers `InvalidParam passed to GetItem(id)` (office-js#2903); waiting after `slides.add()` cost SSF-Charts 18 of 19 probe answers in one round | SSF-Charts, tried and reverted | **Adopted as doctrine.** Nothing here reads through a proxy after the insert, and waiting after an add is on the rejected list. |
| `addTextBox` deletes the SELECTED shape on the web (office-js#2775) | SSF-Charts' `dropShapeSelection` | **Relevant as a class, not a call.** Nothing here adds a text box, but an insert lands while the user may have something selected. SSF-Merge's probe asks whether a SLIDE insert survives a standing selection rather than assuming it. |
| A picture cannot be inserted while a shape is selected, and `setSelectedShapes([])` may never resolve (office-js#3698) | SSF-Charts' selection ladder | **Relevant as the same class.** The call itself is never made here; whether a slide insert is safe with a standing selection is what SSF-Merge's probe asks, unanswered as of 2026-09-08. |
| A collection load over ~50 items answers short (office-js#4272) | `ID_PAGE` in SSF-Charts' renderer; `ID_PAGE = 20` in SSF-Merge | **Relevant, and re-triaged 2026-09-11.** This row promised the sibling's paging, and the paging was never built. The deck read is a plain collection load, guarded instead by the deck's scalar count in the same batch: when the list is shorter than the count, no index is handed out. Both callers turn the list into an index and the removal after an insert is positional, so a wrong index is a wrong slide. |
| The web uppercases tag KEYS internally and needs the uppercased spelling to read them back (office-js#6079) | SSF-Charts' tag writer; SSF-Merge's keys are uppercase | **Relevant, and honoured.** Every insert writes `SSF_SLIDE_ELEMENT` and `SSF_SLIDE_ELEMENTS_CATALOGUE`, uppercase from the first day as this row asked. |
| A `SlideRange` id lacks the deck's `#suffix` (office-js#2474) | SSF-Charts' selection code | **Relevant.** Whatever names the slide the user is on matches by prefix and refuses two matches rather than guessing. |
| `Slide.exportAsBase64` omits modern comments and `ppt/authors.xml` (office-js#6867) | SSF-Charts' round evidence; SSF-Merge's open probe question | **Relevant, and it DECIDED the read.** SSF-Charts calls that API for a picture and marked it no exposure. This repo's own probe measured the presentation-level export dropping the comment part and `ppt/authors.xml` on the web on 2026-09-10, so `getFileAsync` is what the add-in reads with, and `exportAsBase64Presentation` is deliberately not called. |
| Shape tags are lost when a shape is cut and pasted on the web (office-js#3784) | SSF-Charts' triage; SSF-Merge's manual caveat | **Relevant as a limit, documented rather than guarded.** Nothing this add-in writes onto a shape may be something a later step depends on finding. |
| Inserted content may appear in the slide PREVIEW but not the main view without a refresh (office-js#6498) | SSF-Charts' visibility gate | **Relevant as a support answer.** A user reporting a missing element may be seeing this, and the deck delta will say it landed. |
| `PowerPoint.run` batching fails to load properties reliably after `context.sync()`, web only (office-js#6363) | SSF-Charts' central failure; SSF-Merge's `deckRead` probe | **Relevant, and re-triaged 2026-09-11.** The deck read does not batch across proxies: one `PowerPoint.run`, one collection load and one `getCount` queued before a single sync, both read after it, and nothing loaded across a proxy that outlived a sync. |
| The web forces a full presentation save on every `context.sync()`, read-only syncs included (office-js#6329) | SSF-Charts' `KNOWN_ISSUES`; no row in SSF-Merge | **Relevant.** Every host call this add-in makes is a sync, so they stay few and batched. |
| `getcount-populates-same-sync` — the count is right while the list is empty | SSF-Charts' `FAKE_BASELINE` | **Relevant, and re-triaged 2026-09-11.** The deck read does do a collection load, and this finding is why the scalar count rides in the same batch and wins the disagreement. |
| `getitemat-past-end` — what the host does past the end of a collection | SSF-Charts' `FAKE_BASELINE` | **Relevant.** No paging was built, so what it bounds is removing the replaced slide by index if the deck moved. |
| `which-end-a-short-read-drops` | SSF-Charts' `FAKE_BASELINE` | **Relevant.** A short read that is not a prefix makes a slide NUMBER wrong, not merely a list shorter, and this add-in addresses the slide the user is on by position. |
| `how-many-collection-reads-a-context-survives` | SSF-Charts' `PENDING_QUESTIONS`; SSF-Merge's paging loop | **Smaller than it was, re-triaged 2026-09-11.** There is no paging, so there is no loop: one `PowerPoint.run` per read, one collection load inside it, and no context accumulates reads either way. |
| `delete-then-lookup` — whether a deleted slide still resolves | SSF-Charts' `FAKE_BASELINE` | **Adopted as doctrine, re-verdicted 2026-09-11.** The replaced slide is removed by position, highest index first, and the deck is re-counted rather than the call believed. "Remove from N slides" is a SECOND caller of the same delete and runs it once per slide; each cycle is counted back before the next one starts. |
| `scratch-slides-returned` — whether a probe gets its slides back | SSF-Charts' positional sweep; SSF-Merge's triple-clamped undo | **Adopted as doctrine, re-verdicted 2026-09-11.** Any removal here is positional and clamped. The deck-wide removal can stop half-way — by design, at the first cycle the deck's own size does not confirm — and says how far it got, which is the opposite failure to 45 slides nobody asked for. |

### The rows that are no exposure

The other 54 rows — 20 issues and 34 host questions, counted 2026-09-08 — are
about surfaces this add-in never touches: adding shapes through the API,
grouping, creation ids, held shape and slide proxies, tags written through a
proxy, bindings, pictures, rasterising, layouts and masters, rotation and
presets. Each row in `TRIAGED` says which surface it is not, so a later reader
can see the check was made. They become exposure the moment anything here adds
a shape through the API, and that route is on the rejected list in
`docs/BACKLOG.md`.

### What reading it found in OUR code

Nothing on 2026-09-08, and there was nothing to find: the repo held no host
code at all when this ledger was written. Repeated on **2026-09-11** against
the host layer that has since shipped, it found one thing, and it was this
file. Four rows described a paged deck read that was never built, over code
that does the plain collection load those rows say is never done — which is
the live hazard in office-js#4272 rather than a wording slip. The code now
reads the deck's scalar count in the same batch and refuses an index when the
list and the count disagree, and the rows say so. SSF-Merge's own sweep found a
product defect this way, which is the argument for the read.

## What we learned that the siblings have not

Nothing about the host yet. Two things about a sibling's repository, found
while copying it, are worth carrying back to SSF-Merge:

- Its unified manifest names the 64px colour icon where the v1.17 schema asks
  for 192×192.
- Its `CONTRIBUTING.md` documents `npx vite --https.key … --https.cert …`,
  flags Vite 8 does not have; the dev certificate has to reach Vite through its
  config.

Both are corrected here and noted in the scaffold pull request for the owner to
carry across. SSF-Charts is read-only by the owner's instruction.

## Keeping this file honest

It is subject to its own rule. The citation count and the row counts above are
dated because they are counters; they will be stale rather than wrong. Re-count
before quoting either.

**The sweep runs weekly** (`.github/workflows/sibling-watch.yml`, Mondays at
08:41 UTC, after SSF-Charts' own sweeps at 06:17 and 06:41 and SSF-Merge's
watcher at 07:41, so a finding they triage that morning is in their tables
before this looks). `scripts/sibling-watch.mjs` reads the tables above and
reports anything with no row in its `TRIAGED` map — one issue, reopened and
updated, never one per sweep. Run it yourself with `npm run sibling-watch`.

Raw file reads only, never the GitHub API: all three repositories are public,
so it needs no token and runs in CI and in an agent session alike. It never
imports a sibling's code — a weekly job that executes a file fetched over the
network is a supply chain, not a sweep.

Ids are namespaced by whose finding it is, not where it was read:
`charts:issue:1650`, `charts:question:getitemat-past-end`. SSF-Merge's
`TRIAGED` keys are SSF-Charts' findings, so a finding in both tables is one row
carrying both. Two things SSF-Merge's reader could not do, both tested here:
read a quoted key that carries a colon (its own rows are spelled
`"issue:1650":`), and read a table that is declared and empty as empty —
SSF-Charts' pending register is empty by design, and SSF-Merge's watcher has
been red on it since 2026-09-02. An ABSENT table is still a broken sweep, for
every source.

Three exit codes, and the third is the one that matters: 0 when everything has a
row, 3 when something does not, and anything else when the sweep itself broke.
A table renamed upstream throws by name rather than matching nothing, because
"nothing new, every Monday, forever" is indistinguishable from a quiet week and
is exactly the failure this file exists to prevent.

**`TRIAGED` is the source of truth and this prose is downstream of it.** Every
reason opens with a verdict from a closed vocabulary — `NO EXPOSURE`, `ADOPTED`,
`RELEVANT` — and `test/sibling.test.ts` fails when a finding with any verdict
other than `NO EXPOSURE` has no line here.
