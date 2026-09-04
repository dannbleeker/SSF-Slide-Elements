# What is proven, and what is assumed

This project's siblings learned the hard way that reasoning about a host is not
the same as measuring it, and that a sentence written here becomes something the
next reader builds on. So this file says plainly which claims have been tested
and which have not.

## Proven, in the suite, on every commit

- **The harvest reads the library deck correctly.** 1 cover, 10 dividers, 98
  elements, no slide unaccounted for. Every content slide carries exactly two
  placeholders (`title`, `sldNum`) and 677 non-placeholder shapes between them.
- **Every element splices into a real deck and produces a valid slide.** 98 of
  98, checked for dangling relationships, missing parts, undeclared content
  types and duplicate shape ids, with the whole relationship graph walked
  transitively. This gate has failed for real twice — at 0 of 98 — and both
  failures were engine bugs:
  - `<mc:AlternateContent>` branches were being given different shape ids,
    splitting one shape into two. The library's own cover slide is such a shape.
  - 41 think-cell shape tags (`THINKCELLSHAPEDONOTDELETE`) were classified as
    slide furniture and left behind, so the spliced slide named 41
    relationships that did not exist.
- **The picker.** Search, grouping, counts, truncation and escaping, in jsdom.
- **The pane's layout.** No horizontal overflow and no spilled label at 320 and
  512, in both themes, in a real browser (`npm run pane-shots`). This caught a
  real defect the suite could not see: the card label's `-webkit-line-clamp`
  computed to an inert `flow-root` and hard-clipped text mid-line.

## Assumed, and NOT yet measured against a real PowerPoint

Everything in this section is a reasoned choice, not a reading. Each is written
so a single round in a real host settles it.

1. **`insertSlidesFromBase64` accepts a package pruned to one slide.** The
   package keeps every other slide's PARTS and removes only their entries from
   `<p:sldIdLst>`, which OPC permits. Unreferenced parts are legal; whether
   PowerPoint agrees on the way in is not measured.

   This route was chosen specifically to AVOID `sourceSlideIds`, whose id format
   this project has not measured and whose documentation is inconsistent about
   the version it arrived in. If pruning turns out to be the wrong bet,
   `sourceSlideIds` is the alternative and it needs its own round.

2. **The insert lands after `targetSlideId` and the original can then be
   removed.** SSF-Merge has measured that an insert WITHOUT a target lands at the
   FRONT, and that a target works. What is unmeasured here is the ordering of an
   insert immediately followed by a positional delete of the slide before it.

   The order is deliberate — insert first, prove the deck gained a slide, and
   only then remove — so the failure mode is a duplicated slide the user can
   delete, never a lost one. That is by construction, not by measurement.

3. **`getSelectedSlides()` names the slide the user is looking at, and its
   position in `slides` matches the position in `<p:sldIdLst>`.** The engine
   addresses slides positionally, on the sibling's hard-won rule that this host's
   ids look like `256#3561048925` and that both a position and an id are strings.
   Whether the two orderings agree is assumed.

4. **Shape geometry from `occupiedOn` is in points.** Office.js documents it as
   points and the engine converts by 12700. A wrong unit here does not break
   anything — it only makes the placement avoid the wrong region — which is why
   it is not guarded.

5. **How long a 1.5 MB round trip takes.** The budgets in `src/host/timeout.ts`
   are guesses shaped by the sibling's, not measurements. A library deck is small;
   a user's deck may not be.

## The rule this file exists to keep

**One answer sheet is not evidence about a host — it is evidence about that host
in that minute.** When a round is run, record what it answered and on which
build, and do not promote a claim from this section to the one above it on a
single sample. Where two explanations fit an observation, add the question that
separates them rather than reasoning about which is likelier.
