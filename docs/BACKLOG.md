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

The screenshot is no longer on this list. It shipped on 2026-09-15 as
[`docs/listing-screenshot.png`](listing-screenshot.png) — `docs/LISTING.md` is
the recipe, `scripts/listing-shot.ps1` retakes it, and `test/listing.test.ts`
holds it to 1366×768.

Left, and the first of them is the only thing anything else waits on:

1. **The listing name**, which waits on the same naming-policy answer SSF Merge
   waits on. Needs a decision, not a measurement. **Nothing else here can move
   until it is settled.**
2. **The Partner Center submission** of `manifest-prod.xml`, after 1.
3. **v0.1.0** on the releases page. `npm run release:check` is the pre-flight.

### Widen what the mutation sweep changes

The one open item **this repo can finish on its own**, and the only one that is
code.

`scripts/mutants.mjs` now runs **seven** operators — boundary, operands,
boolean, negation, fallback, guard, off-by-one — over `src/core`, `src/host`
and `src/pane`. Every survivor they can find is closed: 48 recorded
equivalents, and the last sweep of each left nothing alive.

**`operands` was added on 2026-09-14 and found nothing, which is the result.**
It swaps a comparison's operands — `rect.cx <= room.x` becomes
`room.x <= rect.cx` — which moves a comparison's SENSE where `boundary` only
moves its EDGE. 81 sites, **81 killed, 0 surviving**. Two were checked by hand
rather than trusted to the counter, because a mutant that fails to parse also
counts as killed: the overlap test in `cut.ts` went red on "paints out a part
that intrudes on the crop" (1 box expected, 0 produced), and the export-loss
comparison in `probe.ts` went red on "names what the export dropped" with the
verdict flipping `yes` to `no`. Both are behaviour, not syntax.

So the suite already distinguishes the direction of every comparison the
operator can reach, and **that lowers the expected value of the three
candidates left**, which were ranked on the same reasoning this one was:
dropping an argument at a call, replacing a returned object's field with its
zero value, and exchanging `&&`/`||` chains' grouping. Of those, only the
returned-field one asks a question the seven do not — an argument dropped at a
call is mostly a type error, which the type checker already refuses. Worth
doing when something else is not more valuable; not worth doing next simply
because it is here.

Each needs the same treatment every operator here has had — a gate in
`test/mutants.test.ts` that is **proven to fail without it**, and a `FAST` row
if it makes a file expensive. `--what <operator>` sweeps one operator alone,
which is what makes adding one cost minutes rather than a whole sweep.

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

### Probe question 4 is answered

On **PowerPoint for Windows, 2026-09-14**, on `template/probe-comments.pptx` —
the deck authored for it, carrying a modern comment and `ppt/authors.xml`.
`exportAsBase64Presentation` handed back **43 parts where `getFileAsync` gave
48**, dropping the comment part, `ppt/authors.xml` and the three
`ppt/webextensions/` parts. Both runs of the pair agreed, ten minutes apart.

That is the last of the seven questions to be measured by this repository's own
instrument. The engine reading with `getFileAsync` is no longer borrowed from
SSF-Merge's sixth sheet. `docs/DESIGN.md` section 15 carries the pair, together
with the two instrument rules it cost — an imported snippet must be TRUSTED
before Script Lab will run it, and the failure is invisible in the runner's
console.

The web has not been re-asked on this build and does not need to be: the drop is
the same defect (office-js#6867) the siblings measured there, and nothing in the
engine turns on the platform.

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

The Windows pair of the same day is the second reason. `getFileAsync` read a
0.04 MB deck in **62, 76 and 61 ms** there, and `exportAsBase64Presentation` in
28 and 30 ms. Whatever the web's 31,755 ms was, it is not a property of the
call.

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
