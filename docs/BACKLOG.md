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

The listing name is settled as **`SSF Slide Elements`** (owner, 2026-09-16).
Recorded as a DECISION and **not** as a policy clearance: nobody has read it
against certification policy 1100.7, and if a reviewer objects the answer is a
rename rather than a claim that it was cleared. `test/listing.test.ts` holds
`docs/LISTING.md` to saying so.

Left:

1. **The Partner Center submission** of `manifest-prod.xml`. The owner's, and
   the only step that needs a Microsoft sign-in.

### A slide deleted mid-run is still unmeasured on a host

What is left of a bigger item, and it is a flaky MEASUREMENT rather than
anything known to be wrong.

A run passes over a slide that has GONE — `indexOfSlide` answers nothing, the
cycle is skipped, `done` does not count it, and no position is deleted in its
place. `test/pane-wiring.test.ts` covers that. What no round has managed is to
watch a HOST do it: on 2026-09-23 the first attempt deleted a slide the run had
already reached, and the second landed inside a cycle's count confirmation.

A cycle costs about 110 ms here and its confirmation is a small part of that, so
a deletion fired blind lands in the awkward window perhaps one time in six. The
way to take it is a long run — 120 slides — with the deletion aimed well ahead
of the cursor and repeated until it lands in the clear, asserting that the
outcome reads one short of the total, that the deleted slide is absent, and that
nothing else moved.

The half of this that WAS a defect is fixed: a count disagreeing because the
USER had changed the deck used to be reported as "the deck has a slide too
many", sending somebody to look for a duplicate that did not exist. The cycle
now asks whether the slide it aimed at actually went, which the count cannot
say, and two cases in `test/pane-wiring.test.ts` — one per loop — hold it.

## Settled — do not re-open

### The undo's refusals are measured on both hosts

**Closed 2026-09-24**, Windows then the web, all four cases on each
(`docs/DESIGN.md` section 15). Every refusal left the deck identical and the
control restored the user's own slide, which is what makes the refusals mean
anything. Three things it settled:

- **The extra listing read costs 11–59 ms on the web**, against 11–32 ms on
  Windows. The record had carried it as unmeasured and as the thing the web
  would feel; it is not. The web's cost is the successful undo, 1074 ms
  against 108.
- **#155's sentence holds on the web**: after the manual's Ctrl+Z-twice the
  pane says the user's own slide is already back, rather than that the inserted
  slide is gone. Both are true; the useful one fires.
- **`undoAim`'s duplicate branch is unreachable** through PowerPoint's own
  Duplicate, on both hosts — the copy gets its own creation id. The branch
  stays, because a premise that fails must refuse rather than pick a copy, but
  it is DEFENSIVE and no host produces the state. Do not go looking for a way
  to trigger it.

Two limits, which are honest gaps rather than open work: the reorder was cut
and paste, because a synthetic drag cannot be delivered to an HTML5
drag-and-drop strip, so a hand drag is unmeasured on both hosts; and the web
deck was read through Office.js alone, the same library the product uses,
because the second source failed. Both are in section 15.

**Mac and iPad** stay unmeasured, and there the Undo falls back to the
count-checked positional one, which is what `docs/MANUAL.md` describes.

### The undo's positional aim, and what closed it

**Closed on 2026-09-24.** Kept here rather than deleted because the route out
was three separate PRs and the middle one is a host measurement nothing in the
code records.

The pane's Undo was positional: it took back the slide where its insert left it,
and nothing disarmed it when the user edited the deck. Four paths deleted one of
the user's own slides and reported "Undone.":

- an "as a new slide" insert, then PowerPoint's own Ctrl+Z, then the pane's Undo,
  which deleted the user's next slide;
- any slide added or deleted before the press;
- a count that moved and moved BACK — Ctrl+Z on a new slide, then a slide added
  in its place;
- a pure DRAG, which changes no count at all.

**The count check** closed the first two: the entry holds the count the insert
measured, and a press that finds a different one refuses, changes nothing and
disarms. **The user's own slide's id** closed the manual's Ctrl+Z-twice, which
leaves the count unchanged. Neither cost a host call.

**The creation-id check closed the last two**, and it needed a host answer
first. The engine writes a fresh `p14:creationId` into every rebuilt slide, and
the answer sheets showed a fixture slide's creation id coming back as the
`#suffix` of its Office.js id — but only in POSITIONAL reads. What the
`slides.load("items/id")` LISTING says about a slide `insertSlidesFromBase64`
has just added was unmeasured, and SSF-Charts had measured the two disagreeing
for a fresh `slides.add()` slide on the web. So the question went to a probe arm
(question 8) rather than into a guess, and both hosts answered yes on
2026-09-24: the listing carries the creation id straight away, agreeing with the
positional read, still agreeing later — on the web 3188 ms and a `getFileAsync`
later — and a creation id is unique in it. `docs/DESIGN.md` section 15 has both
sheets.

`undoAim` in `src/host/insert.ts` is the check: at the press, Undo proceeds only
when exactly one listed slide carries the rebuilt slide's creation id and it sits
at the position the delete is about to take. Gone, moved and duplicated all
refuse and name what was seen. It compares the `#suffix` and only the suffix,
never `sameSlideId`, because the prefix is the deck's and the deck reuses it.

**What is NOT closed, and is a fallback rather than a hole:** a host that hands
back no `#suffix` at all, a listing that does not answer inside its budget, and
an entry armed with no creation id. Each falls back to the count-checked
positional Undo, which is where this feature started, and on such a host a drag
is still Ctrl+Z's job — the manual says so. Mac and iPad have never been
measured (`docs/DESIGN.md` section 15 lists every host fact there as assumed),
and a refusal on a validator's first Undo is a worse trade than the drag it
would close.

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

Windows, 2026-09-14: `exportAsBase64Presentation` dropped the comment part,
`ppt/authors.xml` and the three `ppt/webextensions/` parts that `getFileAsync`
returned. That is why the engine reads with `getFileAsync`, and it was the last
of the seven questions to be measured by this repository's own instrument.
`docs/DESIGN.md` section 15 has the pair and the two instrument rules it cost.
The web does not need re-asking: it is office-js#6867, which the siblings
measured there, and nothing in the engine turns on the platform.

### The host rounds that were being asked for have been run

The web (2026-09-10 and 2026-09-14), Windows (2026-09-11), "Remove from N
slides" end to end on the web (2026-09-13), and probe question 7 on both hosts
(2026-09-14). `docs/DESIGN.md` section 15 carries every one with its numbers.
The jump is no longer borrowed from SSF-Charts on either platform; on Mac and
iPad it still is.

### Question 6 is answered and needs no action

`getFileAsync` took 31,755 ms for a 0.05 MB deck on the web on 2026-09-14 and
14,523 ms later in the same run, against 62, 76 and 61 ms for a 0.04 MB deck on
Windows. **That is not a size cost**, so section 13's sixth question cannot be
answered by extrapolating from megabytes — and it is a reason not to trust a
single timing, not a reason to change the engine. `docs/DESIGN.md` section 15
carries both sets.

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
- **Checking the undo by reading the whole deck for the add-in's tag.** Option 1
  of the undo item, rejected on 2026-09-24: a `getFileAsync` on every press
  costs 874 ms on a healthy web session and 40 s on a degraded one (measured
  2026-09-10), and it is not even certain — the `SSF_SLIDE_ELEMENT` tag holds
  the element's catalogue id, so any slide carrying the same element passes.
- **Disarming the undo when the selection moves.** Option 2, rejected the same
  day: Office.js offers no reorder event, only a selection change, which fires
  on every click and so throws away good undos, and nothing measured says a
  Ctrl+Z fires it — which is the path that lost a slide.
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
