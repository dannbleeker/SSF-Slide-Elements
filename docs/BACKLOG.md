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

The listing name is no longer on this list. The owner settled it on 2026-09-16
as **`SSF Slide Elements`** — what the manifests always carried, now a decision
rather than an open question. It is recorded as a decision and NOT as a policy
clearance: nobody read it against certification policy 1100.7, and
`test/listing.test.ts` holds `docs/LISTING.md` to saying so.

Left:

1. **The Partner Center submission** of `manifest-prod.xml`. The owner's, and
   the only step that needs a Microsoft sign-in.

### A slide deleted mid-run is unmeasured on a host, and the message can misname it

Small, and open because two rounds on 2026-09-23 failed to settle it rather
than because anything is known to be wrong.

A run passes over a slide that has GONE — `indexOfSlide` answers nothing, the
cycle is skipped, `done` does not count it, and no position is deleted in its
place. `test/pane-wiring.test.ts` covers that. What no round has managed is to
watch a HOST do it: the first attempt deleted a slide the run had already
reached, and the second landed inside a cycle's count confirmation.

That second attempt found the thing worth writing down. While a cycle is
confirming its delete, the deck shrinking by one is indistinguishable to the
pane from its own delete having failed, so it stopped with *"the deck has a
slide too many: the copy was made but the original could not be taken away"*.
The deck had no extra slide — the user's own deletion accounted for the
difference — so the sentence named the wrong cause. It stops and tells the user
to look, which is the right direction to be wrong in, and nothing was lost.

Closing it means the cycle distinguishing "my delete did not land" from "the
deck changed underneath me", which the count alone cannot do; reading the ids
either side of the delete would, at the price of another read on every cycle of
every run. Worth doing only if it turns out to happen to anyone.

### The undo still aims at a position the user may have moved

**Narrowed on 2026-09-23, not closed.** The undo puts the user's original slide
back and then takes the rebuilt one away positionally. The window INSIDE the
undo is now guarded — for an "onto this slide" undo the insert half reads the
id at that index, and the delete half checks it is still there — but the window
from the INSERT to the Undo button being pressed is not, and cannot be closed
the same way.

Why not: the slide the undo deletes is one this add-in created, and `CLAUDE.md`
records that a slide the run just added does not resolve by id on the web. The
pane never held an id for it, so there is nothing to compare. A user who
inserts, drags slides around in the strip, and then presses Undo is aiming at a
position that has moved — and the count check cannot see it, because the undo
adds one slide and removes one whichever slide goes.

What would close it, and what it costs:

1. **Read the deck and check the slide at that index carries the add-in's
   tag.** Every inserted shape gets an `SSF_SLIDE_ELEMENT` tag written into the
   package before the insert, so the rebuilt slide is identifiable. This is
   certain, and it costs a whole `getFileAsync` on every undo —
   `src/host/timeout.ts` records 874 ms on a healthy web session and 40 s on a
   degraded one, on the one control a user presses when they want something
   taken back NOW.
2. **Disarm the undo when the selection moves.** Cheap, and wrong often: a user
   clicking about the deck loses an undo that was perfectly good.
3. **Leave it.** The failure needs a reorder between an insert and its undo,
   the undo is one deep, and the pane already says it cannot undo more.

This is a judgement about how much an undo may cost, which is the owner's. It
is written down rather than guessed at.

## Settled — do not re-open

### The mutation sweep has run every operator it was meant to

`scripts/mutants.mjs` runs **nine** operators — boundary, operands, boolean,
negation, fallback, guard, off-by-one, field, grouping — over `src/core`,
`src/host` and `src/pane`. Every survivor they found is closed, by a test or by
one of the **50** entries in the `EQUIVALENT` ledger. The one candidate never
built, dropping an argument at a call, is in the rejected list below with its
reason. A new operator is still welcome if it asks something these nine do not;
it needs a gate in `test/mutants.test.ts` proven to fail without it, and
`--what <operator>` sweeps it alone.

**`field`, 2026-09-23.** It empties each top-level field of a returned object
literal to `undefined`, asking whether any test READS a value where the others
ask whether it is computed right. 323 sites over 20 of the 35 files; the
sweep took 1 h 57 m and answered **265 killed by a test, 58 type-killed, 0
surviving, 0 hung**. The 58 are required fields — 42 of them `verdict` and
`detail` in `src/host/probe.ts`, 11 in `src/pane/steps.ts`'s outcomes — whose
presence the type checker guarantees and whose VALUE no test asserts at that
return. By the script's own rule that is caught by `npm run typecheck` in CI
and is not a hole; it is written down so that "the tests pin every probe
sentence" is not what anybody concludes from the clean result.

**`grouping`, 2026-09-24.** It keeps every `&&` and `||` and moves the
brackets. Prettier brackets every mixed chain, so the form that matters here is
the explicit one — an `||` chain in brackets as an operand of `&&` loses them,
and a bracketed `&&` chain inside an `||` has its last conjunct take the rest of
the chain. Only **6 sites** in the 35 files, which is what the backlog had
predicted about its value; and **3 of the 6 survived**, which it had not:

- `shown` in `src/pane/search.ts` showed every element of the picked category
  whatever was typed, and no test combined a picked category with a search
  that ruled out an element inside it.
- `slideShapes` in `src/core/splice/shapes.ts` and the paragraph pass in
  `src/core/splice/splice.ts` each lost their NAMESPACE test for all but one
  local name, so a foreign `extLst` or `endParaRPr` was treated as
  PresentationML's or DrawingML's own. Nothing held the namespace half of
  either rule.

Each now has a case proven to go red on that exact mutant; the other 3 are
type-killed. The same run found the ledger's stale check crying wolf — it
reported all fifty entries as stale on a grouping-only run, because it filtered
by file and not by operator — and that is fixed and held by its own case.

### The deck-wide removal is held over several cycles, and has no cap

**No cap on how many slides a removal touches**, and not for want of a number:
the owner asked on 2026-09-23 for a way to STOP a run instead, which is the
answer that needs no number. A Stop control sits in the footer while a run of
several cycles is going and takes effect between cycles, so the rest of the
deck is genuinely untouched (`docs/DESIGN.md` section 5).

**The loop is tested past its first cycle** since 2026-09-23. Until then every
removal case drove ONE cycle, because `deckWithStampOn` built its deck with the
real splice, which lists only the slide it rebuilt, so the fixture could not
put the element on two. It now puts the deck's own slide list back after each
splice — what the host's insert-then-remove does to the deck — and five cases
in `test/pane-wiring.test.ts` drive the loop over two and three slides: every
slide in order with the one between left alone, a first cycle that does not
land stopping the run, a second cycle that strands its copy after the first
finished, a slide deleted mid-run skipped while the one after it still runs,
and Stop between cycles.

Measured the same day: four breaks of `removeEverywhere` — the insert-count
`break` turned into a `continue`, the per-cycle id lookup read once before the
loop, the gone-slide `continue` turned into a `break`, and the Stop check
removed — each passed every one of the file's 100 earlier cases, and each is
caught by exactly one of the new five. The gap was real, not theoretical.

### The two library decks filed their elements differently — fixed and re-printed

Found 2026-09-23 by a finder over `scripts/harvest.mjs`, whose cross-deck gate
compares element KEYS only and so had never seen it: the 16:9 deck carried **11**
categories to the 4:3 deck's **10**, and **five elements** sat under a different
heading depending on which shape of deck was open. The pane's chips, the summary
count and `public/catalogue.html` all come from `library.categories`, which is per
size, so the library a user saw depended on their slide size.

The owner settled both content questions the same day: 16:9 over-divided, so its
`Hvide kasser med sorte overskrifter` heading went and those four elements fell into
`Hvide kasser`; and `[ KPI definition … ]` belongs under `One-page templates`, so the
4:3 deck's copy moved there. Both decks now harvest to the same 10 categories in the
same order, 106 elements each, **zero filed differently**, and `npm run harvest` no
longer prints its categorisation warning at all. Shipped in `5726c6a`.

**One measurement worth keeping, so nobody chases it again.** A 16:9 print taken
part-way through came back **1,365,499** bytes against a committed 2,186,758 — a 38%
drop for one slide fewer out of 109 — while 4:3 moved only −0.4%. That mattered
because `npm run previews` cuts every tile picture out of these PDFs by page, and the
gate checks SHA, slides and pages, never resolution. Re-printing both decks together
by one method does **not** reproduce it:

```
library-16x9.pdf   2,186,758 -> 2,185,156   (-0.07%, 109 -> 108 pages)
library-4x3.pdf    2,274,248 -> 2,274,097   (-0.01%, 107 -> 107 pages)
```

Old against new, both decks: 11 embedded images either side, max image width 604 and
600 unchanged, total image pixels identical to the pixel (1,673,959 and 1,004,954),
same filter mix, same full dimension multiset. Nothing was downsampled.

**What caused the 1.37 MB print is NOT established** — the files it came from could
not be examined — only that a COM print of these decks does not do it. Do not re-open
this to chase it; if it recurs, the thing to capture is the print itself, at the
moment it is made.


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
- **Probe question 7** is answered on the web and on Windows, both on
  2026-09-14: `setSelectedSlides` moves the view, puts the previous selection
  back, and leaves the host answering afterwards. On the web the call took
  998 ms and 1,146 ms on two sheets twelve minutes apart, the next selection
  read 590 ms and 567 ms; on Windows the call took 7 ms and the next read
  58 ms. The jump stops being borrowed from SSF-Charts on both platforms; on
  Mac and iPad it is still the sibling's.

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
- **A mutation operator that drops an argument at a call.** Rejected by the
  owner on 2026-09-23 with the `field` operator's plan: a dropped argument is
  almost always a type error, which `tsc` already refuses and which every
  survivor is already put through, so the sweep would spend its time producing
  type-kills.
- **Twelve pane ideas the owner turned down** on 2026-09-08, listed in the
  decisions log of `docs/DESIGN.md`: among them an authored tag vocabulary, a
  per-element description, a "New" chip, swapping the element already on the
  slide, Enter inserting the top hit, several new slides in one pick, authored
  sets, version-aware elements, a density toggle and a go-to-category dropdown.
