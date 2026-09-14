# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build. The open section below is in that order, so its top entry is what
to pick up next. The settled and rejected sections after it are there to stop a
closed question being re-opened by someone reading the code rather than the
record.

(The headings are matched literally by `test/docs.test.ts`, which slices from
the first occurrence of the open heading to the next one at that level — so
this paragraph must not spell either of them out.)

## Open

### Release and AppSource

**The only thing between this and shipping**, and most of it is written already.

Built and held by tests: the 300×300 store logo (`npm run icons`), the
descriptions and the validators' testing notes (`docs/LISTING.md`, held to the
manifests by `test/listing.test.ts`), the catalogue page on the site
(`public/catalogue.html`, generated at harvest), the privacy page naming the
browser storage, and the validators' test deck (`template/validators.pptx`,
authored by PowerPoint over COM rather than assembled here, held to its
description by `test/validators-deck.test.ts`).

Left, and every one of them needs the owner rather than this repo:

1. **A 1366×768 screenshot** of the pane beside a real presentation. Needs a
   screen.
2. **The listing name**, which waits on the same naming-policy answer SSF Merge
   waits on. Needs a decision, not a measurement.
3. **The Partner Center submission** of `manifest-prod.xml`, after 1 and 2.
4. **v0.1.0** on the releases page. `npm run release:check` is the pre-flight.

### Probe question 4, on a deck that has comments in it

The 2026-09-14 sheets could **not** answer it: the validators' deck carries no
comments and no `ppt/authors.xml`, so `exportAsBase64Presentation` had nothing
to drop and the sheet says "not asked" rather than passing. The 2026-09-10 and
2026-09-11 sheets are still the answer; this would confirm it on this build.

Two halves, and only the second needs the owner:

- **A deck with a comment and an `ppt/authors.xml` in it.** This repo can author
  one — `template/validators.pptx` was made the same way, by PowerPoint over COM
  — and nothing about it needs a host round.
- **One Script Lab round on the web** with that deck open. `docs/PROBE.md` says
  how; the 2026-09-14 round is the worked example.

Low cost, low urgency: it confirms an answer two older sheets already give,
rather than filling a hole. It is listed above the sweep only because it is the
last of the seven host questions not measured on this build.

### Widen what the mutation sweep changes

The one open item **this repo can finish on its own**, and the only one that is
code.

`scripts/mutants.mjs` runs six operators — boundary, boolean, negation,
fallback, guard, off-by-one — over `src/core`, `src/host` and `src/pane`. Every
survivor those six can find is now closed: 48 recorded equivalents, and the last
sweep left nothing else alive. So the next thing the sweep can teach us costs a
new operator, not another run.

Candidates, in the order they look worth trying: swapping a comparison's
operands, dropping an argument at a call, replacing a returned object's field
with its zero value, and exchanging `&&`/`||` chains' grouping. Each needs the
same treatment every operator here has had — a gate in `test/mutants.test.ts`
that is **proven to fail without it**, and a `FAST` row if it makes a file
expensive.

Deferred once already, on 2026-09-14, in favour of closing what the existing
operators had found. That work is done.

## Settled — do not re-open

### Mac and iPad will not be measured before release

The owner has **neither device** (2026-09-12), so a Mac round stopped being a
release requirement and both sit where iPad already sat: the pane degrades
honestly through the runtime floor check, the validators' testing notes disclose
that the publisher has measured neither, and **the validators' report is the
first measurement**. `docs/DESIGN.md` section 9 carries the reasoning and the
three ways an earlier measurement could still arrive — cheapest being a borrowed
device running the **Script Lab probe**, which needs no sideload of this add-in
at all. `test/listing.test.ts` holds the testing notes to disclosing exactly the
platforms with no answer sheet, so the disclosure cannot rot when a sheet is
filed.

This is a decision, not a pending task. Nothing is waiting on it.

### The host rounds that were being asked for have been run

- **The web** is measured three times over: two pairs of sheets from
  2026-09-10, a pair from 2026-09-14, read by `docs/PROBE.md`.
- **Windows** is measured: a pair of sheets from 2026-09-11 and a product round
  of four inserts and undos the same day, both read in `docs/DESIGN.md`
  section 15.
- **"Remove from N slides", end to end**, on the web on 2026-09-13, build
  `931bc1d`: the stamp onto three slides, the question, `Removed from 3
  slides.`, and the deck **read again** to check it rather than the footer
  believed. `docs/DESIGN.md` section 15 carries it. The same round settled the
  jump and "Move to a new slide".
- **Probe question 7** is answered, on the web on 2026-09-14: `setSelectedSlides`
  moves the view, puts the previous selection back, and leaves the host
  answering afterwards — 1,146 ms and 567 ms, twelve minutes apart. The jump
  stops being borrowed from SSF-Charts **on this platform**; Windows and Mac are
  still the sibling's.

### Question 6 is answered and needs no action

`getFileAsync` took **31,755 ms for a 0.05 MB deck** on 2026-09-14, and 14,523
ms later in the same run, against `exportAsBase64Presentation` at 550 ms for the
same deck — while the product round the day before saw 12 seconds for the same
work. That is not a size cost, so section 13's sixth question cannot be answered
by extrapolating from megabytes. **Nothing there is a reason to change the
engine; it is a reason not to trust a single timing.** `docs/DESIGN.md` section
15 carries it.

## Rejected — do not re-propose

- **Inserting shape by shape through Office.js.** Rejected before trying, on
  SSF-Charts' rounds: setting text through the API re-authors it
  (office-js#5858), `addTextBox` deletes the selected shape on the web
  (#2775), `setSelectedShapes` wedges the host (#3083, #3698), a shape proxy
  does not survive a `context.sync()`, and a run of eight charts took minutes.
  An element with a gradient, a table or an icon cannot be reproduced through
  the shape collection without losing something. The package route has none of
  those failure surfaces, and it is what the family has measured.
- **Mining the discarded first implementation** (commit `fbc0870`, PR #1,
  closed unmerged 2026-09-08). Rejected by the owner: it shipped a manifest
  version Office rejects, named scripts that did not exist, sat behind both
  siblings on every toolchain major, pointed its CNAME at a different subdomain,
  and measured nothing against a host. The increments above are rebuilt from
  the siblings, not from it.
- **Twelve pane ideas the owner turned down** on 2026-09-08, listed in the
  decisions log of `docs/DESIGN.md`: among them an authored tag vocabulary, a
  per-element description, a "New" chip, swapping the element already on the
  slide, Enter inserting the top hit, several new slides in one pick, authored
  sets, version-aware elements, a density toggle and a go-to-category dropdown.
