# SSF Slide Elements — project memory

A library of ready-made slide elements for PowerPoint, as an Office.js task pane
add-in. Part of the SSF family; SSF-Merge is where the package layer and most of
what this repo knows about the host came from, and SSF-Charts is the record of
what shape-by-shape drawing costs.

## Architecture in one paragraph

A .pptx is a zip of XML parts. The insert happens **in the file**: read the open
presentation, splice the element's markup into the shape tree of the slide the
user is on, reduce the package to that one slide, hand it to PowerPoint through
one `insertSlidesFromBase64`, and remove the slide it replaced. Source and
destination are the SAME DECK, so the spliced slide carries the deck's own
layout, master and theme and there is no theme drift to reconcile. Nothing is
drawn shape by shape, so none of SSF-Charts' per-shape failure surfaces exist.

`src/core` is pure TypeScript with **zero Office imports**, enforced by
`test/architecture.test.ts` in both directions.

## The three signals the library deck already carries

This is the design decision everything else rests on, and it was MEASURED
against the real deck rather than chosen:

- **The LAYOUT says what a slide is for.** `Frontpage` is the cover (1 slide),
  `Sektionsadskillelse` opens a section (10), `Kun titel` is an element (98).
  No slide in the deck is unaccounted for.
- **The TITLE placeholder is the name** — of a section, or of an element.
- **A PLACEHOLDER is furniture; everything else is the element.** Every one of
  the 98 content slides carries exactly two placeholders, `title` and `sldNum`,
  and 677 non-placeholder shapes between them. No exceptions.

So adding an element is: open the deck, draw it, push. There is no sidecar file
and no naming convention. **Do not add one** without a reason that survives the
question "what does the deck already say?"

`LayoutRoles` in `harvest.ts` names the layouts rather than hard-coding the
Danish ones, so a second library in another language is a config line.

## Where things live

| directory | what it owns |
| --- | --- |
| `src/core/pptx/` | `pkg.ts` (zip, parts, rels, content types, slide list), `xml.ts` (one parser everywhere), `parts.ts` (what a relationship is CALLED, in one place) |
| `src/core/catalogue/` | `harvest.ts` (library deck in, catalogue out), `types.ts` (the index/payload contract) |
| `src/core/element/` | `splice.ts` (the insert, and the four rewrites), `verify.ts` (is it still a slide PowerPoint will open) |
| `src/core/preview/` | `svg.ts` — the thumbnail, drawn from geometry rather than rasterised |
| `src/host/` | the DECISIONS: `capability.ts`, `placement.ts`, `undo.ts`, `timeout.ts`, `errors.ts` |
| `src/office/` | the Office.js CALLS, and nothing else. Every judgement imported from `src/host` |
| `src/pane/` | `render.ts` (the picker's DOM), `main.ts` (the only file here allowed to touch Office.js) |
| `template/` | the library deck. The authoring surface, and the input to the harvest |
| `public/catalogue/` | generated, committed, shipped. `public/` because that is what Vite copies verbatim into `dist` |

## The four rewrites a splice performs

Every one of them has a failure that opens as "PowerPoint found a problem with
this file", with nothing to say which. `verify.ts` maps one finding to each, so
a red result names the step rather than leaving somebody to bisect.

1. **Part names** — the element's `ppt/media/image3.png` and the target's are
   different pictures with the same name.
2. **Nested relationship targets** — a chart's `.rels` names its workbook by
   path, stale the moment the workbook is renamed. No visible symptom until
   somebody opens the chart's data.
3. **Relationship ids** — `rId3` means different things in two slides.
4. **Shape ids** — a collision does not always break the file, which is worse.

## Gotchas, all of them paid for

- **`<mc:AlternateContent>` branches share one `<p:cNvPr id>` deliberately.**
  `<mc:Choice>` and `<mc:Fallback>` carry the same shape and only one is live.
  Renumbering them apart splits one shape into two that PowerPoint disagrees
  with itself about — and the library's own COVER SLIDE is such a shape, so this
  was wrong on the first thing anybody would have inserted onto. `renumberShapeIds`
  maps old id to new id, so a repeat inside one shape stays a repeat. `verify.ts`
  skips `<mc:Fallback>` when counting for the same reason.

- **`<p:tags>` under a shape's `<p:nvPr>` is that SHAPE's data, not the slide's.**
  Classified as slide furniture at first, which left 41 dangling relationship ids
  on one slide. They are all `THINKCELLSHAPEDONOTDELETE` — this library was built
  with think-cell, and carrying them is what keeps a chart editable after it
  lands. Shape tags travel; the slide's own layout and notes do not.

- **An unmapped relationship id is REMOVED, never left.** A slide naming a
  relationship its `.rels` does not define is exactly what PowerPoint calls a
  damaged file. `rewriteRelIds` drops the attribute, and takes a `<p:tags>` with
  it because that element is nothing but its reference. The harvest now carries
  everything a shape can name, so this should never fire — it is the invariant,
  not the plan.

- **A name ending in a digit is not a duplicate.** The first duplicate-name
  warning regexed the generated id for a numeric suffix. Half this library's
  names end in a number ("…med checkliste på 4"), so it fired on all of them and
  buried the one real duplicate: **slides 58 and 59 have the same title**, and
  the 4:3 deck shows 59 should read "5 kasser". Ask the counter, never the id.

- **`-webkit-line-clamp` is not reliable here.** The card label computed to
  `flow-root` — blockified — so the clamp was inert and the text was hard-clipped
  mid-line with no ellipsis, in the BUILT pane only. It looked exactly like the
  minifier dropping `-webkit-box-orient`, and it was not. Truncation is done in
  the STRING (`clampName`), which behaves the same in every WebView PowerPoint
  might embed and is visible to jsdom, where a CSS clamp is not.

- **Read a shape's transform from its OWN properties, never by descending.** A
  group's children each carry one in the group's child coordinate space; the
  first `<a:off>` a descendant search finds belongs to the group only by luck of
  document order. Same rule for `children` vs `elements` in `xml.ts`.

- **The preview must honour the group coordinate transform.** Roughly half the
  library's elements are groups, and a preview that ignores `chOff`/`chExt` piles
  every group's contents in the top-left corner — which reads as a rendering bug
  rather than the maths one it is.

## Conventions

- **Branch flow**: develop on the session's designated `claude/*` branch; after
  each merge, reset it onto `origin/main` (`git checkout -B <branch> origin/main`)
  — never stack on merged history. One PR per increment.
- **Every npm script stays FLAT.** A script that nests `npm run` is blocked by
  AppLocker on the owner's Windows box, so `build:lib` and the script that needs
  it are run as two commands, never as one.
- **A regression test must be proven to fail without its fix.** Both engine bugs
  above were found by the splice sweep going red at 0 of 98; that is the standard.
  Check WHICH assertion goes red, not just that one does — the architecture guard
  passed for a day against five clean files because it was matching the word
  "PowerPoint." in prose.
- **Looking at the pane is part of done.** `npm run pane-shots` renders and
  MEASURES it at 320 and 512 in both themes. jsdom has no layout and no colour,
  so a rule about how the pane looks is invisible to the suite.
- **Say what is measured and what is assumed.** `docs/UNPROVEN.md` is the ledger.
  The file engine is proven on every commit; the HOST HANDSHAKE is not yet
  proven against a real PowerPoint at all, and nothing in this repo should imply
  otherwise until a round says so.
- **When two explanations fit the evidence, measure — do not reason.** The label
  clamp cost three wrong fixes to a minifier that was innocent; one
  `getComputedStyle` settled it.

## Commands

```bash
npm test               # the whole suite, including the 98-element splice sweep
npm run typecheck
npm run lint
npm run build:lib && node scripts/harvest.mjs   # regenerate the catalogue
npm run build          # the site
npm run pane-shots     # needs `npx vite preview --port 4178 --strictPort &` first
```

## Backlog

- **A real-host round.** Everything in `docs/UNPROVEN.md` section 2, in one
  session, on one build, recorded.
- **The 4:3 library.** The owner has a second deck (100 slides) with no divider
  layout at all — it is the older version. `fitScale` already handles a 16:9
  element going into a 4:3 deck; the deck itself needs its sections adding before
  it can be harvested.
- **The redraw fallback.** The agreed design is file surgery with a shape-by-shape
  fallback for hosts too old for the file route. Only the surgery half is built.
  `src/host/capability.ts` already answers the question that would select it.
- **Slide 59's title**, which duplicates 58's. The owner's deck, the owner's call.
