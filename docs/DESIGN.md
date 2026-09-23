# SSF Slide Elements: the design record

This is the design the pane and the insert are built to. It was settled with
the owner over one day of demos (2026-09-08), decision by decision, and every
decision below carries its date. The package layer, the harvest (section 3)
and the host probe (section 13) are built; `docs/BACKLOG.md` is the order the
rest arrives in.

Two rules keep it honest. A design change updates this file in the same PR, so
the record and the code never disagree. And everything here is **assumed**
until the host-probe increment has measured it; section 15 says which is which.
Nothing in the design should be built on a guess that a single round against
PowerPoint would settle.

## Contents

1. [What the user does](#1-what-the-user-does)
2. [What an element is](#2-what-an-element-is)
3. [The library](#3-the-library)
4. [The pane, top to bottom](#4-the-pane-top-to-bottom)
5. [Where an element lands](#5-where-an-element-lands)
6. [Inserting](#6-inserting)
7. [The options behind the gear](#7-the-options-behind-the-gear)
8. [Search](#8-search)
9. [Accessibility and platforms](#9-accessibility-and-platforms)
10. [Loading and failure states](#10-loading-and-failure-states)
11. [Performance](#11-performance)
12. [What AppSource certification needs](#12-what-appsource-certification-needs)
13. [Open questions for the host](#13-open-questions-for-the-host)
14. [Build order](#14-build-order)
15. [Measured and assumed](#15-measured-and-assumed)
16. [Decisions log](#16-decisions-log)

## 1. What the user does

1. Opens the pane from the ribbon button **Slide elements**. The pane shows the
   library that matches the open deck's slide size, grouped by category, with
   search and tags at the top.
2. Hovers an element: the tile grows a little and the name goes bold. After a
   third of a second a preview card opens with the element at full width and a
   small slide showing where it will land, next to what the slide already has.
3. Clicks it. The element lands **onto the current slide** or **as a new slide
   after it**, per a setting behind the gear. Insert first, prove the deck grew,
   then remove the replaced slide (the host rules in `CLAUDE.md`).
4. The footer says what happened, with the deck's slide count before and after,
   and offers **Undo**.

## 2. What an element is

- **A library slide is one element**: everything on it that is not layout
  chrome (title, footer, slide number). Grouped or not: 42 of the whole-slide
  elements in the 16:9 deck have no group at all (tables, matrices, the
  one-pagers), so the harvest cannot depend on grouping.
- **A collection slide yields one element per top-level shape, and the slide's
  title is their category.** The owner's collection slides give _Markers_ and
  _Stamps and labels_. There was a third, _Flowchart shapes_, until the owner
  had it removed on 2026-09-16 — one slide in each deck, ten part elements on
  it, and the category went with the slide that carried its heading. The
  _Icons_ collection is gone too, and not by a decision of its own: it held the
  scales and the waste bin, the waste bin left with Flowchart shapes, and the
  owner had the scales removed the same day — a SHAPE deletion off the Stamps
  and labels slide rather than a slide deletion, since seven other elements
  share that slide. A collection slide is
  marked by a line `SSF: ét element pr. figur` in the notes of its category
  heading slide, inherited by the slides under it, overridable per slide.
  Instruction text on such a slide (the "good rules for flowcharts" box) goes
  into the slide's notes, not onto the slide.
- **Names.** A slide element is named by its title. A part is named by its own
  text with brackets stripped, else by its PowerPoint shape name when that is
  not a generic `Gruppe N`, else by the slide title numbered. The deck's TITLES
  stay Danish and are the **keys**; `template/names.en.json` carries the English
  name for every key, one file per locale later, and the harvest fails on a key
  with no name. The pane shows only the English name: no Danish tooltip, no
  per-element description.
  **A PART's key is English**, because a part is keyed by its own text and that
  text was translated on 2026-09-11: eleven keys moved with it, `Fortroligt` to
  `Confidential` and so on, and `names.en.json` moved with them. Two consequences
  worth knowing. An element's `id` is a slug of its key, so those eleven ids
  changed — a favourite or a "Used in this deck" tag written before that date
  names an id the catalogue no longer has. And `tagsFor` matches the key against
  a word list that was written in Danish; a rule a part can reach needs both
  spellings, which is why it now reads `dokument|document` and why four elements
  had quietly lost their `meeting` tag before it did.
- **Placeholder text in the library is English**, in both decks, done
  2026-09-11: 76 distinct strings across 3,393 paragraphs, so what lands on the
  slide reads in the pane's language. Per-locale text comes later the way names
  do. Left in Danish, deliberately: slide titles (the keys), and the 4:3 deck's
  276 off-slide guidance shapes, which no element carries. Left alone as well:
  the deliberate filler — lorem ipsum, `xxxx`, a Wingdings tick — which is meant
  to read as obviously fake in any language.
  Slide titles are never shown because the splice skips `title` and `ctrTitle`
  placeholders when it takes the library slide's shapes
  (`src/core/splice/shapes.ts`), for a new slide as well as onto the current
  one — so a Danish title cannot reach a user's deck by either route.
- **A shape PowerPoint does not draw is not part of an element.** A top-level
  shape marked `hidden="1"` in the Selection Pane is skipped by the harvest,
  the way layout chrome and an empty placeholder are. Not a tidiness rule:
  think-cell parks an invisible OLE frame at the slide origin on every slide it
  has touched, and the 4:3 library deck has been through it — 41 of its 106
  slides carry one. Measured on the committed catalogue on 2026-09-23, before
  the rule existed:
  - it joined the element's BOX, which is a union, so **42 elements** came out
    anchored at x=0.0002 and about 93% of the slide wide, against the same
    element in the 16:9 deck at x=0.0573 and 88%. The box is what the landing
    places from, what the preview crops to, and the frame `authored` rebases
    from, so all three were computed for a rectangle nearly the size of the
    slide;
  - it was serialised into the MARKUP, so **41 of the shipped 4:3 elements**
    put think-cell's frame into the user's deck on every insert;
  - and it dragged its payload: **49 OLE binaries** published under
    `4x3/parts/ppt/embeddings/` and copied into the user's presentation with
    the element. After the rule: 8, which are the owner's own charts, and the
    4:3 deck's carried parts fall from 215 to 122.

  The check is on the TOP-LEVEL shape only, which is where the harvest decides
  what content is. A hidden shape inside a group the owner drew is the owner's
  artwork and is carried as authored.
- **An element that comes in several sizes is one tile with a stepper.** A run
  of elements that differ only by one count (process flows with 1 to 6 boxes,
  hierarchies with 2 to 5 boxes, matrices with 2 to 5 rows, and so on) is
  derived from the name pattern: a number that stands alone before a word,
  never inside "2×2" or "1-2-3", "box" the one irregular plural, and the
  key-figure flows one run by an explicit rule. The stepper names what it
  counts ("boxes 1 2 3 4 5 6"). The pane never uses the word "family"; that
  word is for this document. The 16:9 deck has twelve such runs, so 106
  elements show as 73 tiles: the runs cover 45 elements, so 106 − 45 + 12.
- **Categories** are the heading slides (a slide with a title and no content).
  **Tags** are derived from names; there is no authored tag vocabulary.
- An off-slide shape (x at or beyond the slide's right edge) is never part of
  an element.

## 3. The library

- **Two decks, one per slide size**, authored by the owner and committed under
  `template/`: `library-16x9.pptx` and `library-4x3.pptx`. Both carry the same
  106 keys, and the harvest refuses a key that is not in both, so no element
  is ever scaled from the other size. A deck that is neither 16:9 nor 4:3 (A4,
  16:10, a custom size) borrows the nearest library, scaled to fit, and the
  line under the pane's header says so.
- **Previews are PowerPoint's own rendering.** There is no PowerPoint in CI, so
  the owner prints each deck to PDF and commits the print beside the deck; the
  harvest cuts every element's preview from it. Every deck change therefore needs
  a re-print. A part is cut to its box with 3% of air, the boxes of neighbouring
  parts painted white, and a rotated part masked to its rotated frame. The parts
  are the same objects in both decks, so their cuts are shared.
  The geometry of that is `src/core/catalogue/cut.ts`, pure and decided in
  tests rather than by judging a rendered picture: 3% of the element's OWN size
  on each side, clamped so a crop never reaches off the page; a neighbour
  painted out only when it is a part, on the same slide, and actually intrudes
  on the crop, clipped to it; and the mask as the four corners of the unrotated
  frame turned about its centre. That frame is why an element now carries
  `rotation` — `box` is the rotated EXTENT, and the extent cannot be un-rotated
  back into the frame it came from. Two elements per deck have one: the stamps,
  at −29.06° and 35.02°.
- **Both decks are on one colour palette, since 2026-09-11.** They were not, and
  nothing had noticed: the 16:9 deck (made 2021 from the owner's own template)
  was on the stock Office scheme, while the 4:3 deck still carried **"07 Blå"**,
  the palette of the company it was made at in 2013 — its `docProps` still names
  that company. So **79 of the 117 element pairs rendered in different colours
  in the two sizes** — 117 being what the libraries held on 2026-09-11, before
  the Flowchart shapes category was removed; the figures in this paragraph are
  that measurement and are not restated against a later library. The same box
  was Office orange at 16:9 and light blue at
  4:3, the same rule Office blue and grey. The 4:3 deck's `<a:clrScheme>` was
  replaced with the 16:9 deck's, and nothing else in the package was touched —
  one part changed of 568, verified part by part by SHA-256, and the result
  opened by PowerPoint itself with its 108 slides and the new accent. After it,
  96 of the 117 pairs resolve to identical colours and the remaining 21 differ
  only in how many whites and blacks two hand-drawn twins hold.
  The decks' notes and handout themes still carry the old palette. Deliberate:
  neither is harvested and neither reaches a user's deck, and an edit to a deck
  that cannot be opened here is not made for tidiness.
  **The 4:3 print was retaken the same day**, without a screen, through
  PowerPoint COM (`SaveAs(pdf, 32)`; `ExportAsFixedFormat` cannot be bound on
  this build at any arity). A print taken by a different route is a print to
  check rather than trust, so it was compared with the dialog print of the same
  deck page by page: 108 pages either way at the same page size, the outer ring
  of every page inked exactly as before — which is what rules out a stray
  "frame slides" — 78 pages rendering pixel-identical, and the 29 that moved
  being the ones the palette touched. The sidecar records which route took it
  and `test/print.test.ts` holds that field to the two spellings.
- **Each print carries a sidecar naming the bytes it was taken from**, and that
  is the gate — `template/library-16x9.print.json` and its 4:3 twin, holding the
  deck's SHA-256, the print's own SHA-256, the slide and page counts, the date,
  the PowerPoint build and the export settings used.
  Page count against slide count was the original check and it is not enough on
  its own: **110 pages match 110 slides whatever the pages are OF.** A print of
  yesterday's deck, of a corrected copy, or of a different deck the same length
  all pass it, and the previews then come out cut from the wrong file with
  nothing to say so. That is not hypothetical here — the first pair of prints
  was taken from slash-corrected COPIES while the committed decks would not
  open, and was thrown away rather than committed for exactly this reason.
  The rules are in `scripts/print-provenance.mjs`, the stamping in
  `npm run print-stamp`, and `test/print.test.ts` holds both. It was proven to
  fail the way a gate must: a byte appended to the committed print gives
  "the print has changed since it was stamped … re-stamp it", and a byte
  appended to the committed deck gives "the deck has changed since the print was
  taken … re-print it".
  **Both prints are committed**, `template/library-16x9.pdf` (109 pages,
  2,186,758 bytes) and `template/library-4x3.pdf` (107 pages, 2,274,248 bytes),
  re-taken on **2026-09-16** through COM after the Icons slide left the library
  (section 14, item 4) and stamped into `template/*.print.json`, which is what
  `test/print.test.ts` holds them to. Those figures supersede the 2026-09-11
  dialog print — 110 and 108 pages — which section 15 keeps as the record of
  that round. Section 14's items 2 and 3 have not happened, so both prints need
  retaking again after that pass, and the cut IS downstream of them now:
  `src/core/catalogue/cut.ts` and `npm run previews` take every tile's picture
  out of exactly these two files, so a deck edited without a re-print and a
  re-stamp is caught by the gate rather than quietly cut from the wrong page.
  `*.pdf` is
  declared `binary` in `.gitattributes` for the same reason `*.pptx` is — the
  repo's `* text=auto eol=lf` would otherwise leave a print to git's binary
  heuristic.
- **Both decks open in desktop PowerPoint. They did not until 2026-09-11, and
  why is worth keeping.** Measured on Windows that day (PowerPoint
  16.0.20326.20132, Microsoft 365 Current Channel, x64): both decks were refused
  with "PowerPoint found a problem with content", offering **Repair**. The cause
  was two `<Override>` entries per deck in `[Content_Types].xml` whose `PartName`
  was missing its leading `/` — `ppt/notesSlides/notesSlide2.xml` and
  `notesSlide3.xml` in the 16:9 deck, `notesSlide13.xml` and `notesSlide14.xml`
  in the 4:3 deck. OPC requires an absolute part name, and `Package.Open` refused
  both decks with "Part URI must start with a forward slash." Nothing else about
  either package was wrong: the zip was intact, no relationship dangled, and
  every part carried a content type.
  **Accepting the Repair was not the way out**: on a copy it opened the deck and
  dropped ten of the fifteen `ppt/embeddings/oleObject*.bin` parts, and
  `ppt/changesInfos/changesInfo1.xml` with them. The four slashes were added
  instead, and nothing else: all 375 and 568 parts kept the bytes they had,
  `[Content_Types].xml` is identical once the leading slashes are normalised
  away, and the harvested index did not move — version `1641fe687794` before and
  after. Both decks then opened at 110 and 108 slides with no prompt.
  **If a deck stops opening again, ask `System.IO.Packaging.Package.Open`
  first**: it names the violating part, where PowerPoint's dialog names nothing
  and Repair reports only what it removed.
- **Boxes.** The box of a rotated shape is its rotated extent, not the
  unrotated frame the XML gives (the owner's stamps are rotated 29° and 35°). A
  table's box is the sum of its columns and rows, not its frame's `ext`, which
  PowerPoint ignores when it draws the table. Landing and cropping both depend
  on this.
- **Colours.** The library is authored in theme colours (accent 2 alone is used
  166 times across the 16:9 deck's slides and 174 across the 4:3 deck's, on 45
  and 44 slides; counted from the committed decks on 2026-09-11, where the
  figure this line carried before — 139 — matched neither). A theme-mapped colour follows the destination deck's theme by
  itself when the markup keeps it as a theme reference; the colour switch
  (section 7) pins them to the library's values instead.
- **The catalogue is static.** `npm run harvest` reads the decks into the
  catalogue the pane ships as files on Pages: an index (`catalogue.json`,
  about 160 KB: names, categories, boxes, sizes and landing per element for
  both sizes, and a content hash as its version) plus one file per element with
  its markup, relationships and the parts it carries, and the parts themselves.
  The index is committed and CI fails when it no longer matches the decks; the
  element files and parts (about 16 MB of generated JSON across both sizes,
  measured 2026-09-11) are built on every
  deploy and never committed. Thumbnails and previews join them with hashed
  names (section 11). Adding an element is editing the deck, re-printing, and
  merging a PR. The site also gets a catalogue page generated
  at harvest, every element with picture and name, for browsing outside
  PowerPoint and for the AppSource screenshots.

## 4. The pane, top to bottom

- **Header**, blue, with the SSF tick above the name. The tick (the small orange
  bar above a heading) is the design system's signature and the one orange in
  the pane: orange in one place at a time, or it stops meaning anything.

  **The build stamp used to sit at its right, and no longer does** (2026-09-15).
  It is on the root element as `data-build` instead. The reason it was drawn has
  not gone away — PowerPoint caches the pane's HTML for about ten minutes, so a
  round opened too soon after a deploy tests code the host never fetched and
  reads as a clean run of the wrong build, and this project used the stamp twice
  on 2026-09-14 to avoid exactly that. What changed is that the owner asked for
  it out of the AppSource screenshot, and `docs/LISTING.md` forbids retouching
  the picture, so the stamp had to leave the pane rather than leave the image. An
  attribute is invisible to a user and to a capture and still there for devtools,
  for a support request, and for a driver reading the DOM. **"Report a problem"
  still carries the same value in its prefilled URL**, which is the route for a
  user who cannot open devtools. Hover rings, the star, the ghost frame and the coach
  marks are blue or navy. Under the header, only when the deck's size is
  neither 16:9 nor 4:3: "4:3 library, scaled to A4 slides".
- **Search**, one field, placeholder "Search", `/` focuses it, Esc clears it.
- **Tags**, one line until opened with the chevron at its right; a picked tag
  moves to the front so it stays visible when the line is closed. The chevron
  only shows when there is a second line.

  **Built 2026-09-23, having been described here since the record was written.**
  Until then nothing drew a chevron: the line's one-row clip was lifted by
  `state.gear` — the OPTIONS panel — so it unfolded as a side effect of an
  unrelated control and could not be opened deliberately at all. Worse, a
  constant capped the line at the first TWELVE tags, and both committed
  libraries carry twenty-four, so half the vocabulary never reached the DOM and
  could not be used as a filter. The line now draws every tag and the CSS clips
  it to one row.

  Whether there is a second line is answered by a COUNT (`TAGS_PER_ROW` in
  `render.ts`), not by layout: the pane cannot measure whether the line wraps,
  and `npm run pane-shots` — which needs a browser — is the instrument that
  would check it at 320 and 512. The count is deliberately low, so the chevron
  is offered whenever it is needed and at worst offered once when it was not.
- **The gear**, beside the search: insert target, shapes as one group or loose,
  colours, "Report a problem", "Browse the catalogue on the site" (section 7).
- **Used in this deck**: the library elements already in the deck, each with
  the slide numbers it is on; a number jumps to that slide. Read from the tags
  the add-in writes at insert.

  Built 2026-09-11, with two deviations the record now carries:

  - **It reads on request, not on open.** The section starts as one line — "See
    what this deck already uses" — and reads when clicked. Reading it means
    reading the user's WHOLE presentation, and how long that takes on a fifty
    megabyte deck is section 13's sixth open question, still unanswered. Asking
    costs one click; reading on open would spend an unmeasured cost on everybody
    who never looks at this list. Once read, an insert updates the list rather
    than re-reading the deck, and an undo takes its entry out again — the pane
    knows exactly what it just put where.
  - **The slide numbers are links, and the pane says "Slide N" only when it saw
    the host there.** Built 2026-09-12 on `setSelectedSlides` (PowerPointApi
    1.5), on SSF-Charts' measurement rather than this repo's: that sibling ships
    the same call and its archive from 2026-08-13 to 2026-09-04 holds 2,429
    selection-ladder rungs on PowerPoint for the web with none silent and none
    refused (section 15, "Borrowed"). The call the family saw wedge the host is
    `setSelectedShapes`, which is still never made. After the call the pane
    reads the selection back in the same batch, and only a read-back naming the
    slide produces "Slide N"; a host that answers another slide, or nothing, or
    raises (office-js#3552: desktop throws while the notes pane has focus) gets
    "click slide N in the strip" and no claim. Below 1.5 the numbers stay text.
    `goToByIdAsync`, the older route, was rejected on the tracker's record
    (#2595, #2631, #567). Probe question 7 measures the call directly on the
    next round of any platform.

  Three states, and the pane says which: never asked, asked and empty ("Nothing
  from the library is in this deck yet"), asked and answered. An id the current
  catalogue cannot name is kept and shown as "an element from an older version
  of the library" — eleven ids changed when the part keys were translated
  (section 2), and a row silently dropped would make the deck look emptier than
  it is.
- **Favourites** (a star in the tile's corner, remembered per machine) and
  **Recent** (the last six inserts) sit above the categories.
- **Categories** start collapsed, with "Open all" beside the count; a search or
  a tag opens what it finds; category headers stick while scrolling.

  **Each heading carries a mark that turns when it opens** (2026-09-16). The
  heading had always been a button with `aria-expanded`, so a screen reader was
  told it opened and a sighted user was not: bold text with a hairline under it
  and no reason to think it was pressable. The mark is a real glyph, not a
  drawn triangle, because the pane is checked in forced colours every week and a
  border trick vanishes there. The row answers hover, too.

  **On a deck the pane has never seen, the FIRST category opens** (2026-09-16).
  Before it, the first sight of a new deck was a search box, a row of tags and a
  column of shut headings — not one element on screen, and nothing saying a
  heading opens. The first rather than the biggest: the order on screen is the
  library's own, and "the top one is open" is a rule a user can see. Nothing
  opens when Favourites or Recent already put tiles on that screen. Whether it
  is a first visit is answered by `storage.ts` from whether anything was ever
  written for the deck, NOT from an empty list of open categories — those are
  two different things, and reading them the same would re-open the top category
  every time somebody closed it. The first category is the first of the library
  that MATCHES the deck, which is not the one the pane starts from: the pane
  draws the 16:9 library while the deck is still being measured, and the two
  libraries do not carry the same categories in the same order. So the choice is
  made again when the real library arrives (2026-09-16). It went in choosing
  once, against the provisional library, and nothing noticed because both real
  libraries happen to begin with the same category — luck, not a rule, and on a
  library that differed the first screen would have gone back to the column of
  shut headings this exists to prevent.

  Built 2026-09-12, with one detail the record did not have: the control is
  WITHHELD once every category is open, and while a search or a tag is on,
  because in both states it would be a click that changes nothing. There is
  no "Close all" beside it — one control is what was approved, and a category
  closes by its own header. Tiles are
  two across, three from 400 px. A tile is the element's picture and name; a
  sized element carries its stepper; a part already in the deck carries
  "Remove from N slides" — which appears only once the deck has been read, since
  before that the pane does not know what is in it (section 6 has the rest).
- **The preview card** opens after a third of a second of hover or focus: the
  element at full width, its name, one line saying where it lands. At 512 px and
  wider it docks beside the list and hides no tiles, the list giving up a gutter
  for it and only while one is open.
  **The grey boxes for what the slide already holds are built** (2026-09-11),
  and they are a SNAPSHOT rather than a live read. Everything measured against
  the user's slide — those boxes, and where the splice says an element landed —
  is in the USER's slide size, never the library deck's: the two are different
  numbers exactly when a deck borrows the nearest library, which is the case
  nobody tests on. The card's little slide is drawn in the user's own shape for
  the same reason, so the element's own frame is the approximate one there,
  which is what the borrowed line under the header already says. They come out of the deck
  read the pane already does when it opens — the same one that measures the
  slide size, so they cost nothing extra — and they are stamped with the slide
  they were read from. On another slide the card draws none: the boxes answer
  "will this land on top of something", and an answer about a different slide is
  a wrong answer rather than a missing one. They keep up with an insert and an
  undo without re-reading, because the splice says where the element landed, and
  "See what this deck already uses" refreshes them. Re-reading the whole
  presentation on every slide change is the cost section 13's sixth question has
  not measured.

  Two things this says differently from how it was approved, both on 2026-09-11
  and both from building it:
  **It pins to the BOTTOM under 512 px, not over the top of the list.** It must
  not DISPLACE the list, and that is the load-bearing part: a card in the flow
  pushes the tiles down by its own height, which slides the tile out from under
  the cursor that opened it, so the pointer lands on a different tile and a third
  of a second later the card is showing the wrong element. It was built in the
  flow first and the shots showed exactly that. Overlaying the TOP would cover
  the search box and the tags, which are the two controls a user is most likely
  to be reaching for while browsing, so it overlays the bottom.
  The grey boxes for what the destination slide already has were the half of
  this that waited: they need the slide's own shapes, and that was called host
  work no round had done. It turned out not to need a host call at all — they
  are read out of the FILE the insert already reads, and they shipped on
  2026-09-11 as the snapshot described above. The ghost frame for the landing
  is drawn from the catalogue and never needed one either.
- **Footer**: the last outcome with the measured slide count, then the actions
  (**Move to a new slide** when a whole-slide element landed on a slide that
  already had content, **Undo (n)**), then a line with the current
  settings that opens the gear — and shuts it again, since it is the same
  disclosure control as the ⚙ above the list. Both carry `aria-expanded` and
  point at the panel by id, which the footer's did not until 2026-09-23: a
  screen-reader user pressing it heard a plain button whose own name comes from
  the settings and so does not change, while a panel opened at the top of the
  pane outside their reading position with nothing announced.
  Built in that order, which is this list's:
  measured on 2026-09-12, the row wraps to two lines at 320 px with Undo alone
  on the second whichever way round the first two go, so the order is the
  record's and nothing more.
- **First open ever**: three coach marks (hover to preview, click to insert,
  Undo and the gear), dismissed once and remembered per machine.
  Built 2026-09-11 as three LINES OF ONE PANEL above the search, not three
  callouts pointing at the controls. The pane is 320 px at its narrowest, where
  three floating callouts would cover the very tiles, footer and gear they were
  pointing at — and a coach mark that hides what it is describing teaches
  nothing. The panel says the same three things and leaves the pane readable.
- **The pane reopens where you left it**: search, tags, open categories,
  scroll position, picked counts and settings are kept per deck across pane
  closes and reloads, in the browser's storage on the user's machine.

  Built 2026-09-12, with four things decided in the building:

  - **Two buckets, and which half goes where.** Favourites and the first-run
    flag are per MACHINE — a star is a statement about the library and the coach
    marks are "dismissed once", not once per deck — and everything about how the
    library was being READ is per deck. Recent went per deck too, which this
    record did not say either way: it exists so the thing just used is easy to
    reach again, and "just used" is a fact about a deck.
  - **The deck is told apart by a HASH of its URL, not by the URL.** Only
    equality is ever asked, and a SharePoint path can name a client, a project
    or a person, so there is no reason for one to sit in storage where anything
    else on the origin could read it back. The query string and the fragment are
    dropped first: a OneDrive URL for one file is not one string across
    sessions, and a key that moved with `?web=1` would remember nothing.
  - **No URL means the per-machine bucket**, which is the case for a deck that
    has not been saved. Two unsaved decks then share one memory. Forgetting
    outright was the alternative, and losing a search on every close is the
    behaviour this feature exists to remove.
  - **"Picked counts" is `chosen`**, which the pane also uses for the keyboard
    cursor. It is written only by the operations that already persist, so what
    comes back is the step last INSERTED rather than wherever the keyboard was
    left.
  - **The scroll position** followed on 2026-09-12, and is the one item here
    that is not a field of the state. It is deliberately kept out of it: an
    offset in the state means a re-render per scroll event, and re-rendering the
    whole list while it moves under the user's finger is the one thing this
    must not cost. So it is a module variable, written to storage on a 250 ms
    trailing timer — a pane torn down inside that window loses a few pixels of
    where the user was, which is the only thing this stores — and put back once,
    after the first draw that has TILES in it. Not merely a browse step: the
    loading screen is one short paragraph, and scrolling that to 800 px leaves
    the user looking at nothing. Once, too: a later draw is the user's own
    doing, and re-scrolling them to where they were an hour ago is the pane
    fighting them. A scroll of their own before the restore has happened stands
    down the restore for the same reason. The tiles' pictures arriving later
    move nothing, because `.tile-img` is absolutely positioned inside a box the
    tile has already reserved — which is what makes one restore land where the
    user left off rather than approximately.

## 5. Where an element lands

The landing is decided per collection slide, so the deck decides it, with one
exception for width.

- A **stamp or a label** lands top-right, with a 10% margin because a rotated
  stamp's visible ellipse pokes past its frame.
- A **marker, a flowchart shape or an icon** lands at the cursor. An add-in
  cannot see the mouse on the canvas, so "at the cursor" means the selected
  shape's position (`getSelectedShapes`, read-only, measured safe by the
  siblings), else the slide's centre; the user drags from there. With a shape
  selected, a **marker wraps it**: sized to the shape with a little air and
  centred on it, unless the shape is bigger than about a third of the slide,
  because ink scales its stroke and a wrapped large box turns into a blob; then
  the marker lands centred on the shape at its authored size.
- A **part wider than half the slide** (the breadcrumb bar, the kicker box)
  lands where it sits in the library.
- A **whole-slide element** lands relative to the destination slide's layout:
  below its title placeholder, scaled into its body area when it would not fit,
  and left where it sits in the library when it already fits. A library
  authored for the SSF layout would otherwise overlap a customer deck's taller
  title.
- A part ignores the insert target: it always lands on the slide the user is
  on. A stamp or a label with several slides selected lands on every selected
  slide.

  **Built 2026-09-23, and NOT "in one insert", which is what this line said
  from the day it was written and until the day it was built.** That cannot be
  done: `insertSlidesFromBase64` puts every slide of its package CONTIGUOUSLY
  after one `targetSlideId` — `CLAUDE.md` records a real run that put 37
  generated slides ahead of a title slide — so rebuilt copies of slides 2, 5
  and 9 would arrive as a block and the deck's own order would be gone. Keeping
  the order means aiming each copy at its own slide, and that is one insert
  each.

  So it is `removeEverywhere`'s shape with a different payload: one deck read
  for the whole run, one insert-then-positional-delete cycle per slide, each
  confirmed by the DELTA before the next starts, the positional delete guarded
  by reading the target id back, and the first step that cannot be confirmed
  stops the run with the footer saying how far it got. The slides are worked
  through in ascending order, which matters because each cycle is net zero on
  the slide count — so a later slide is still at the index this code computed
  only once the earlier ones have been put back.

  **The pane's Undo is disarmed and the footer says so.** It is one insert deep
  and positional, so it cannot take back several; a button silently gone is
  worse than a sentence. PowerPoint's own Ctrl+Z reverts an insert (question 5,
  measured on the web and on Windows), and the sentence names it. Unlike a
  removal this is not asked first, because it adds rather than takes away.

  **A cursor-landing part loses its cursor here**, and that is the only sane
  answer: "at the cursor" is the selected SHAPE's position, and with several
  slides selected there is no one shape to take. It falls back to the middle of
  the slide, which is what this section already specifies for nothing selected
  — so two of the ten parts per library, the two markers, land centred on every
  selected slide rather than around a shape.

  **It says which slide it is on and how far through it is**, rather than one
  sentence for the whole run: the pane locks itself, not PowerPoint, and a run
  over many slides is minutes.

  **And it can be STOPPED.** A Stop control sits in the footer while a run of
  several cycles is going, carrying the same numbers the notice does, and it is
  the one control left enabled while the pane is busy — everything else is
  disabled deliberately, one thing at a time, but the thing this interrupts can
  hold the pane for minutes and the only other way out is closing the task pane
  mid-edit. Added 2026-09-23, on the owner's ask.

  **The stop takes effect BETWEEN cycles, never inside one.** A cycle is an
  insert of a rebuilt slide followed by a positional delete of the original,
  and stopping between those two leaves the deck one slide longer carrying
  both — the stranded state this whole path exists to avoid. So the cycle in
  flight finishes, which is what lets the outcome say "Stopped after 2 of 9
  slides. The rest are as they were." and mean it. A stop is NOT a failure: it
  does not mark the outcome as the user's to finish by hand, because there is
  nothing to check.

  Nothing caps the run instead — a cap would need a number, and stopping it is
  the answer that needs none.

  One selected slide, or a host that will not say, runs the ordinary
  single-slide insert with its ordinary Undo — `stampTargets` answers the empty
  list under two slides, which is what hands that path back. A whole-slide
  element is untouched by any of this: three selected slides would be three
  copies of a slide the user asked for once.

  **Not measured on a real host.** The cycle shape is the one the removal has
  run on PowerPoint for the web (2026-09-13); this payload has not, and
  `selectedSlides` — the read behind it — has never been asked of a real
  PowerPoint for more than its first item.

## 6. Inserting

- **One insert at a time.** The pane locks while an insert runs, the tile says
  "Inserting…", and the footer reports the deck's slide count before and after:
  "12 → 13 slides" for a new slide, "12 → 13 → 12 slides, slide 4 replaced" for
  an insert onto the slide, because in the file that is insert a rebuilt copy
  after the slide and remove the original. Two inserts 0.4 s apart killed a
  sibling's tab; the lock is that rule made visible. The delta is the evidence,
  never the absence of an error.
- **Undo** goes **one deep**, positional and count-checked, never by id — with
  one id read that IS checked: for an "onto this slide" undo the restored
  original is aimed by the id at that index, and the same index is read back
  before the rebuilt slide is deleted, because the two are the same slide. That
  closes the window inside the undo. The window from the insert to the button
  being pressed stays open and cannot be closed by an id: the slide the undo
  deletes is one this add-in created, and a slide the run just added does not
  resolve by id on the web, so no id for it was ever obtainable.
  `docs/BACKLOG.md` carries what would close it and what that costs.

  There was an **Again** beside it, repeating the last insert. It went on
  2026-09-16: the element it repeats is the first tile in **Recent**, drawn a
  few lines up the same pane, and clicking that tile does the same thing. Two
  controls for one action, in the place with the least room — at 320 px the row
  wrapped to two lines. The one with the picture on it stayed.

  Ten deep was the decision on 2026-09-08 and it did not survive the build.
  Taking back an insert that landed ONTO a slide means putting the replaced
  slide back, and the only way to put a slide into a deck is to hand PowerPoint
  a package containing it — so ten steps of history is ten copies of the user's
  presentation held inside a task-pane WebView, which is the memory that killed
  a sibling's run. One step is what the pane can honestly hold. PowerPoint's own
  **Ctrl+Z reverts an insert**, measured on the web on 2026-09-10
  (`docs/host-answers/`), so the deeper history already exists and the pane's
  job is to stay out of its way rather than to duplicate it badly.
- **A whole-slide element onto a slide with content** gets the footer offer
  "Move to a new slide" beside Undo. **Empty content placeholders** on the
  slide are removed when a whole-slide element lands, so no "Click to add
  text" ghost sits behind it; the title placeholder stays; Undo puts the
  placeholder back.

  Built 2026-09-12, with three things decided in the building:

  - **What "with content" counts.** `contentCount` in
    `src/core/catalogue/boxes.ts`, reported by the splice as `report.held` out
    of the bytes it already has open, so the offer costs no second deck read
    and no host call. It is `occupiedBoxes`' reading with two differences, each
    because the question is different: the slide's own TITLE does not count,
    since a whole-slide element lands below it and the insert never removes it,
    and a slide holding nothing but its title is the ordinary destination
    rather than a crowded one; and no rectangle is required, since a body
    placeholder the user has typed into is content wherever the layout puts it.
    An empty placeholder still does not count, because the insert takes it
    away.
  - **The move is an undo followed by a second insert**, not a third
    operation. So it inherits both of their guarantees — the undo positional
    and count-checked, the insert reading the deck fresh and proving the delta
    — and adds no new way for the deck to end up somewhere neither can
    describe. When the undo does not work the move stops there with the undo's
    own sentence: a second copy beside the first is the one outcome a user
    asking to MOVE something cannot have meant.
  - **It is offered exactly as long as Undo is**, because it needs Undo to
    work. It stays on screen and greys out while an insert runs rather than
    vanishing: it is about the insert that just happened, which is still the
    last one whatever the pane is doing. A control that disappears and comes
    back under the cursor is worse than one that greys.
- **Multi-shape elements land as one group** (gear option, default on), so the
  user moves them as one and ungroups when editing.
- **Right-click** (long-press on touch) on a tile offers the other insert
  target for that one insert, without touching the setting.

  Built 2026-09-11. Three things decided in the building:

  - **It opens on `contextmenu`**, which is the right mouse button and also the
    keyboard's own menu key and Shift+F10 — so it is not a mouse-only feature by
    accident. The event is cancelled only where a menu of ours opens: right-click
    in the search box still gets the browser's menu, with paste in it.
  - **A part offers nothing.** A stamp or a marker ignores the insert target and
    always lands on the slide the user is on (section 5), so a menu offering it
    "as a new slide" would promise something the engine does not do — the same
    reason the preview card's landing line refuses to say it.
  - **It belongs to a TILE, not to an element.** One element is drawn in
    Favourites, in Recent and in its own category, so a menu keyed by element id
    opened on all three at once — found by rendering every combination of the
    pane's optional state rather than by using it, because it needs Recent to be
    holding the very element the menu is open on. The same is true of the
    question before a removal.
  - **The menu is anchored to the TOP of its own tile**, not to the pointer. The
    pane is 320 px at its narrowest, where a menu at the cursor hangs off the
    edge as often as not; and the tile's name is at its bottom, so a menu there
    hides which element it belongs to. Nothing about where the pointer was
    reaches the pane's state, which is also what lets the shot audit draw it.
  - **A question does not outlive its tile, an insert included.** The question
    is drawn on the tile and nowhere else, and while one is open the Remove
    button is suppressed on every tile — so a question whose tile has gone
    leaves the pane in a state nothing on screen describes. Every change that
    can take a tile off the screen drops it: search, tags, category, star,
    clear, chip, and the insert, which rewrites Recent through `remember` and
    so can drop the oldest entry's tile out from under a question opened there.
  - **One thing open at a time, refused at the OPENING end.** A right-click
    while a removal question is up leaves the browser's own menu alone. The
    pane used to cancel the event and set the menu anyway, while the renderer
    refused to draw one over a question — so the gesture did nothing at all,
    and the menu it had set arrived later out of nowhere, on the redraw after
    the Escape that answered the question. The two tiles can never be the same
    tile, since a question is only ever on a part and a menu only ever on a
    whole-slide element, so this is a rule about the pane, not about one tile.
- **Deck-wide stamps.** A stamp already in the deck can be removed from every
  slide it is on with one click, found by the tag written at insert. The manual
  says that shape tags do not survive cut and paste on the web.

  Built 2026-09-11, and it is the only thing this add-in does that takes
  something OUT of somebody's deck. Four things follow from that, all decided in
  the building:

  - **It is asked, not done.** The tile's button opens a question naming the
    slides — "Take Confidential off slides 2, 5 and 9?" — and saying the pane
    cannot undo it, because it cannot: Undo is one insert deep (above), and an
    undo of this would mean holding a copy of the deck per slide touched. Escape
    answers the question "no".
  - **One cycle per slide, each confirmed before the next starts.** The same
    insert-then-remove sequence an insert uses: hand over a package holding the
    rebuilt slide, prove the deck grew, take the original away, prove it shrank
    back. The first step that cannot be verified stops the run, and the footer
    says how far it got — "Removed from 2 of 3 slides. The rest are as they
    were."
  - **One deck read for the whole run.** A cycle only rewrites the slide it
    targets, so every package is built from the bytes read at the start.
  - **The positional delete is guarded by reading the target's id back.** The
    insert aims by id and survives a reorder; the delete aims by POSITION, and
    that position came from the read at the start of the run. The pane locks
    itself rather than PowerPoint, so a user can drag a slide in the strip
    across the whole of it — and `removeSlideAt` is
    `slides.getItemAt(index).delete()`, which takes whatever is at that
    position now. The count cannot catch it: a cycle adds one slide and removes
    one, so the check agrees whichever slide went. So the id at that position
    is read back after the insert and compared with `sameSlideId`; a mismatch
    leaves the copy standing and stops the run, which is the failure
    `CLAUDE.md` asks for — a duplicate the user can delete rather than a slide
    they have lost. Added 2026-09-23. `insert` grew this guard in #127 and the
    several-slide stamp was built with it; this path, the only one that takes
    content OUT of a deck and the one that runs the most of these, had
    neither.
  - **What it can reach is what this add-in tagged.** The shapes are found
    through `readShapeTags`, which keys on the tag NAME — a user's own shape
    carries no tag of ours, and another add-in's tags in the same folder are not
    ours either. The tag parts and pictures those shapes pointed at are left as
    orphans rather than swept: deciding what else in somebody's deck still needs
    a part is the one class of mistake that produces a file PowerPoint calls
    damaged.

  **Not measured.** No round has run this against a real PowerPoint. The
  mechanism is the insert's, measured on the web and on Windows, but a SEQUENCE
  of them has not been, and section 15 says so.
- **Tags at insert**: every inserted shape or group carries the element's key
  and the catalogue version in a shape tag, written in the package before the
  insert because a slide the run just added does not resolve by id.
- Not doing: swapping the element already on the slide for another; several
  new slides in one pick; authored sets; naming inserted shapes in the
  selection pane; version-aware elements that flag a newer design.

## 7. The options behind the gear

- **Insert**: onto this slide, or as a new slide. Whole-slide elements only.
- **Shapes**: as one group, or loose.
- **Colours**: "This deck's theme" (default), where theme-referenced colours
  follow the destination deck, or "As in the library", where every colour is
  pinned to the value it had in the library deck. Explicit colours stay either
  way.

  Built 2026-09-11, and four things about it were decided in the building:

  - **The library's colours come from the harvest, one map per SIZE.** A
    `<a:schemeClr val="accent1"/>` means nothing without a theme, and the theme
    part is in the library deck, which the pane never sees — so the harvest
    resolves the twelve slots and the four names the master's `<p:clrMap>`
    redirects, and the catalogue carries the answer. A deck whose elements are
    spread over two themes is REFUSED rather than resolved: one map cannot be
    right for both halves, and a wrong map would pin every element to colours
    no slide in the library ever had, invisibly.
  - **Both library decks are on one palette**, since 2026-09-11. They were not:
    the switch is what found it, and section 3 records what was done. An element
    pinned in one size is now pinned to what it is in the other, and
    `test/colours.test.ts` holds the two decks' maps EQUAL rather than holding
    either to a list of hexes.
  - **A carried chart is pinned too.** The library's one chart states 17 scheme
    colours of its own, and a switch that rewrote an element's shapes but not
    the chart inside it would leave the two disagreeing about which deck they
    belong to.
  - **`phClr` is never pinned, and neither is a name the theme has no colour
    for.** The placeholder colour is a style's argument rather than a theme
    slot, so there is nothing here to resolve it to; an unknown name is left as
    a scheme colour and goes on following the destination. Better a colour that
    moves than a colour invented by this code. A colour transform —
    `<a:lumMod>`, `<a:alpha>` — is carried across onto the pinned value, since
    a pin that kept the colour and lost the shade would be wrong twice.
- **Report a problem**: opens the support page in the browser with the build
  stamp, host and platform prefilled, so a report is usable.
- **Browse the catalogue on the site**: opens the catalogue page.
- External links open a browser window from the pane
  (`Office.context.ui.openBrowserWindow`, probed at runtime; a plain new-tab
  link is unreliable on desktop) and never navigate the pane itself.

  Both built 2026-09-11, and three things about them were decided in the
  building:

  - **What may go in the address is an ALLOWLIST, in `src/host/links.ts`.**
    Three values — a build that looks like a commit, a host and a platform that
    are on Office's own lists — and no way to add a fourth. This is a privacy
    rule before it is a formatting one: the pane is running inside somebody's
    presentation, the support page is on the open web, and a URL is the least
    private thing there is. The support page holds the same two lists from the
    other end and shows nothing it does not recognise, which is what stops a
    crafted link putting a word of its choosing on the page somebody in trouble
    lands on. `test/security.test.ts` holds the two spellings together.
  - **The site is the one the PANE was served from**, never a production
    address written into the code — so a dev build links to the dev origin, and
    nothing in the pane can send a user to a site the add-in did not come from.
  - **The links are buttons.** An `<a href>` in a task pane either does nothing
    or navigates the pane away from itself, and a user whose pane has become a
    web page has to close and reopen it to get back. When the host opens
    nothing at all — no `openBrowserWindow` and a blocked `window.open` — the
    pane says so and names the address, because a click that silently does
    nothing is the version the user cannot work around.

  Not yet measured: whether `openBrowserWindow` is there and behaves on each
  host. The requirement set is probed rather than assumed and `window.open` is
  the fallback, so both positions are covered — but which one runs where is a
  question for the next round on each platform.

## 8. Search

Matches the English name, the Danish key (and later every locale's name), the
category and the tags. While searching,
the categories that have hits appear as chips with counts to narrow the search,
and a sized tile greys out the counts that do not match. A query with no hits
offers "Did you mean …" from the nearest names. Not doing: Enter inserting the
top hit.

**The words a search matched are marked inside the name**, on the tile and on
the preview card, built 2026-09-23. `highlight()` in `src/pane/search.ts` cuts
a name into matched and unmatched runs — every occurrence of every query word,
part-words included, touching runs joined — and `render.ts` draws the matched
ones as `<mark class="hit">` through `textContent`, so a name still cannot
become markup and a screen reader still hears the tile's "Insert <name>".

- It looks at the NAME alone. A word that matched only the Danish key, a tag or
  the category marks nothing, because none of those is on screen; the tile
  still shows.
- A name whose lower case is a different length ("İ") is left unmarked rather
  than marked in the wrong place.
- Not on "Used in this deck", which the search does not filter — a mark there
  would claim a filter that is not applied. Not on "Did you mean" either, where
  it could never mark anything: those chips appear only when nothing matched.
- Its own ground (`--mark`), bold, never orange: the tick is the view's one
  orange. In forced colours it takes the system's `Mark`/`MarkText` — which
  Chromium gives a `<mark>` by itself, measured the same day, EXCEPT inside the
  chosen tile, whose `forced-color-adjust: none` is inherited and left the mark
  white on light blue at 1.21:1 until the stylesheet named the pair.
  `npm run pane-shots` now measures contrast in forced colours wherever the
  page has taken its colours back, which is what caught it.

Before that date this section said the highlight was built when nothing drew
one; no user-facing document had promised it.

## 9. Accessibility and platforms

- **Keyboard**: Tab into the list, arrows move between tiles, Enter inserts,
  `/` focuses search, Esc closes a menu, the preview or the search in that
  order. Focus draws the same ring as hover. A live region announces every
  outcome.

  **The arrows belong to the TILES, and to the search box on the way out of
  it.** Pressed anywhere else they are left to the browser, which is what
  scrolls the list. They used to be taken everywhere: the handler ran
  `arrowTo(key, tiles.indexOf(activeElement), n)` whatever the focus was on,
  `indexOf` answers -1 for anything that is not a tile, and `arrowTo` clamps
  -1 to 0 — so an arrow pressed on the gear, a category heading, a tag chip,
  the size stepper, the star or the primary button was cancelled and threw the
  focus to the first tile at the top of the list. In a 320 px pane that also
  took away the only key a mouse user has for scrolling it.

  **Closing something puts the focus back on the control that opened it.** The
  tile menu goes back to its tile, the removal question to that tile's Remove
  button, the gear to whichever of its two controls was pressed. Without it
  each of those landed on `<body>`, because the control the focus was on is
  inside the surface being closed and the redraw removes it — the one case
  `draw`'s restore deliberately hands to the browser's fallback. The fallback
  is right for a tile a search filtered away, which has no owner to go back to;
  a dismissed surface has exactly one, and it is still on screen.

  **None of it worked until 2026-09-22**, and the reason is a property of the
  pane worth stating rather than a slip: `render` empties `#pane` and builds
  fresh elements, so a redraw destroys whatever holds the focus. Focus landing
  on a tile marks it chosen, marking it chosen redraws, and the redraw removed
  the button that had just been reached — so Tab could not get past the first
  tile and every arrow after that landed on tile 0, because the pane's own
  `indexOf(document.activeElement)` was -1. `draw` restores the focus across
  the render now, the way it already restored the search caret.

  **And `draw()` calling `focus()` re-enters the pane's own focus handling.**
  `focus()` raises `focusin`, which is the same event a user arriving at a tile
  raises, so the restore looked like a fresh focus: it marked the tile chosen
  (a redraw, calling itself) and armed the preview's third-of-a-second timer on
  every redraw that had a tile focused. That second half is the one that hid:
  it turned a one-shot into a cadence, and a late draw painted over a completed
  insert — outcome, Undo and all — while the insert itself had plainly run. So
  the restore is flagged and the focus handler returns under it. Anything that
  focuses on render has to account for this, or it arms timers and redraws in a
  loop.
- **Touch**: a tap on a tile INSERTS, and nothing depends on hover. A
  long press opens the same menu the right button does, and the click the
  lifting finger produces is swallowed — without that the menu was closed by
  its own gesture and the element inserted onto the current slide, which is the
  target the menu exists to override (found 2026-09-23, reproduced in jsdom).

  This line read "the first tap on a tile shows the preview, the second
  inserts" until 2026-09-23. **That was never built** — `onClick` inserts on
  the click every tap produces, and nothing in that path consults
  `state.previewing`. It is written down as what the pane does rather than
  built, on the owner's decision of the same day, because neither iPad nor
  Windows touch has an answer sheet and a two-tap rule is a change to the core
  interaction on a platform this repo cannot measure. The preview is still
  reachable on touch: it opens on focus, which a tap gives the tile. Revisit it
  in a round where a real touch host is to hand.
- **The removal question takes the focus and is announced.** Opening it removes
  the Remove button from every tile — the control that opened it — so there is
  nothing for the redraw's focus restore to put the focus back on, and it fell
  to `<body>`: a keyboard user had to Tab from the top of the document to reach
  a confirmation they had opened one keystroke earlier, and a screen-reader user
  was told nothing at all. It is the one action here that takes content out of
  the deck, and the one the pane says it cannot undo, so it is also the one that
  may not open silently. The question itself is the announcement, because it
  names the element and the slides and says the pane cannot undo it. Found on
  2026-09-23 by replaying `draw`'s own focus logic in jsdom.
- **Windows high-contrast mode**: the pane follows forced colours; rings, chips,
  tiles and the tick stay visible. **Measured from 2026-09-12**, when
  `pane-shots` gained a forced-colours pass over every state: until then this
  was a promise nothing rendered, and the tick was not keeping it — a 26x3
  block whose only visual is its background, erased the moment backgrounds are
  forced. It is painted `CanvasText` there now, so it survives in the user's
  own palette rather than in SSF orange, which is what asking for forced
  colours means. The audit refuses any decoration that goes the same colour as
  the thing behind it.
- **Office themes**: light and dark, from `Office.context.officeTheme`. The host
  hands over a COLOUR rather than a name, so the pane decides: BT.601 perceived
  brightness of `bodyBackgroundColor`, and below half of 255 is dark
  (`src/host/theme.ts`). A colour it cannot read leaves `data-theme` unset and
  the stylesheet's `prefers-color-scheme` fallback in charge, which is the right
  answer outside a host — every time the pane is opened in a browser to look at
  it. Written down here on 2026-09-12, when the arithmetic moved out of
  `main.ts`: the threshold had been a number in the one file the coverage floor
  exempts, and the record said only that there were two themes.
- **Pane sizes**: usable at 320 px, three columns from 400 px, docked preview
  from 512 px, and the list stays scrollable with the footer visible in a
  400 px tall pane. `pane-shots` measures these states.
- **Platforms**: PowerPoint on the web, Windows, Mac and iPad. A task pane is
  offered wherever PowerPoint runs and the manifest cannot exclude a platform,
  so **all four are in whether measured or not**.

  The owner has **neither a Mac nor an iPad** (2026-09-12), so both are in the
  same position, and it is not the position iPad alone was in: a Mac round used
  to be a release requirement in section 12 and no longer is. Three ways a first
  measurement can still arrive, cheapest first — a **borrowed device running the
  Script Lab probe**, which needs no sideload of this add-in and answers all
  seven questions in one paste; a borrowed device with the manifest sideloaded,
  through OneDrive on iPad and `~/Library/Containers/…/wef` on Mac; or a
  real-device cloud, if PowerPoint installs and licenses there.

  Until one of those happens the **validators' report is the first measurement**
  for both, which is a deliberate trade and not an oversight. What makes it
  survivable: the pane degrades honestly through the runtime floor check rather
  than a `<Requirements>` element, touch is first class, and the testing notes
  say plainly that the publisher has measured neither — so a finding on Mac or
  iPad reaches us as a first measurement rather than as a surprise.
  `test/listing.test.ts` derives that disclosure from which answer sheets exist,
  so the day a Mac sheet is filed the notes have to stop calling Mac unmeasured.

## 10. Loading and failure states

The pane never shows an empty or broken screen, and every message says what
happened and what to do.

- "Loading the library…" with the header visible at once.
- "The library did not load" with a retry, for offline or Pages down.
- The floor message when the host is below PowerPointApi 1.2, naming the fix:
  a current Microsoft 365, PowerPoint 2021, or PowerPoint on the web.
- "The insert was refused", "The insert did not confirm: the deck still has N
  slides, nothing was changed", and "The deck grew by one but the copy could
  not be removed: delete slide N by hand". A call can raise and still have done
  the work, so every one of these is written from the measured delta.

  A deck that SHRANK gets its own: "The insert did not confirm: the deck has N
  slides where it had M. Check the deck before inserting again." It used to
  fall into the no-op sentence above, which names the count from BEFORE — so
  over a deck of 11 it read "the deck still has 12 slides, nothing was
  changed", a count the deck does not have and a claim the delta refutes. It
  needs no misbehaving host: the pane locks itself and not PowerPoint, and on
  the web the insert and its confirming count take seconds, so a user deleting
  a slide in that window produces it. The pane cannot know whether the insert
  also landed, so the sentence stops at the two counts it took.

  A deck REORDERED while the insert ran gets its own too, and it is the one
  sentence here that deliberately names no slide number: "The insert landed, but
  the deck was reordered while it ran, so the copy was left in place: N → M
  slides. Both your slide and the copy are there; delete whichever you do not
  want." The index the removal uses is read before the host calls and used after
  them, up to `BUDGET.insert` later, and the pane locks itself rather than
  PowerPoint — so a drag in the thumbnail strip moves the slides while changing
  no count, which is the one thing `mayRemove` looks at. The insert survives it,
  because it aims by `targetSlideId`; the removal does not, because it aims by
  position. So the id at that index is read back and compared before anything is
  deleted, and a mismatch — or a read that does not answer — leaves the copy
  standing. That is the house rule made good: the failure mode is a duplicate
  the user can delete rather than a slide they have lost. It names no number
  because the positions this code holds are exactly the ones that just went
  stale, and naming one off a stale index is what sent a user to delete their
  own content before.
- The deck-wide removal runs an insert-then-delete cycle per slide, and its two
  stopping points do NOT share a sentence. When the insert never lands the deck
  is untouched and "Removed from N of M slides. The rest are as they were — try
  again, or take them off by hand." is true. When the insert lands and the
  delete does not, it is false: that slide's original is still there with the
  element on it and an element-free copy sits beside it. That case says
  "Removed from N of M slides, and the deck has a slide too many: the copy was
  made but the original could not be taken away. Check the deck before trying
  again — trying again would add another." It names no slide number, for the
  reason the reorder sentence above gives, and it withdraws the invitation to
  retry because each failed cycle strands another copy.
- A read-only or protected deck, and a deck the host will not hand over
  (`getFileAsync` on an unsaved deck on the web, to be measured).

## 11. Performance

The demo page weighs 6.4 MB because its previews are inlined; the real pane must
be usable in about two seconds on the web. Catalogue JSON first (names,
categories, boxes, no pictures), tile thumbnails about 300 px wide as WebP
(roughly 8 KB each) loaded lazily per open category, the full preview picture on
demand when the card opens, all with long-lived hashed filenames on Pages so a
release never serves a stale catalogue.

**What the engine costs on a big deck, measured 2026-09-11** by
`npm run bench` (`scripts/bench-engine.mjs`), on this machine, in Node — so
these are the engine's numbers and not the host's:

| deck | open | "Used in this deck" | a slide's boxes | one insert |
| --- | --- | --- | --- | --- |
| 40 slides, 0.1 MB | 1 ms | 20 ms | 0 ms | 66 ms |
| 200 slides, 0.3 MB | 3 ms | 86 ms | 1 ms | 134 ms |
| 60 slides, 20 MB | 37 ms | 23 ms | 0 ms | 114 ms |
| 120 slides, 45 MB | 81 ms | 45 ms | 0 ms | 228 ms |

Three readings, and the third is why the table is here at all:

- **Weight, not slide count, is what a deck costs.** 200 slides of XML open in
  3 ms; 45 MB of pictures takes 81. The tag sweep behind "Used in this deck" is
  the one thing that scales with SLIDES, at about half a millisecond each, and
  reading what a slide holds is free.
- **An insert on a 45 MB deck is a quarter of a second of engine time.** The
  host's own two calls — `getFileAsync` and `insertSlidesFromBase64` — are not
  in that number and are still unmeasured (section 13, question 6).
- **It was four seconds until the same day**, and every one of those seconds was
  BASE64. Office hands the deck over as base64 and takes it back the same way,
  and JSZip's own encoder and decoder are pure JavaScript: 1.6 s to decode 45 MB
  and 2.5 s to encode it, against 27 ms and 24 ms for the platform's own. The
  package layer now converts either side of JSZip rather than through it
  (`src/core/pptx/base64.ts`), which is an eighteen-fold difference on the
  largest deck anybody is likely to open — and it was invisible until somebody
  measured, because every deck in the suite is a few hundred kilobytes.

  The memory that buys it is worth knowing: a package handed back as base64
  exists briefly as both bytes and characters, so a 45 MB deck peaks around
  105 MB. That is the same shape the old path had, and it is why the pane holds
  ONE deck copy for undo rather than ten (section 6).

## 12. What AppSource certification needs

Checked against Microsoft's commercial-marketplace certification policies for
Office add-ins (section 1120) and the Partner Center submission form, assuming
the XML manifest is submitted and the listing is free.

Covered by the design or the release plan: a stable GUID and a version bump per
submission; an HTTPS-only origin, no localhost, no `AppDomains` because nothing
navigates the pane; `ReadWriteDocument` only; no sign-in and no data collection;
**requests only to the add-in's own origin** — the pane fetches its catalogue
from the site it is served from, which is why `SECURITY.md` says "it sends
nothing anywhere" rather than "no network calls"; the support and privacy pages
on the site, the privacy page naming the browser storage (the gear's settings,
favourites, the last six inserted, and which categories were left open) and
that nothing leaves the machine; **our own terms of use and a licence page**
(2026-09-16, replacing Microsoft's standard EULA — the standard text is written
for an add-in that might do anything, so it could say neither that the
presentation never leaves the pane nor that an insert REWRITES one slide of the
file it is given, and Microsoft's own submission rules forbid a listing URL
pointing at the repository, so "the licence is in the repo" was not an answer
either; `termsOfUseUrl` lives in the JSON manifests only, so the move needed no
re-sideload, checked by hashing the XML either side of it); the publisher
**StruktureretSundFornuft ApS** as a company — the owner settled the spelling on
2026-09-16, when the sibling SSF Merge's terms page turned out to carry the
`ApS` and this project's manifests did not, and a publisher display name that
does not match the registered one is a certification question nobody wants to
answer twice; first-run guidance; keyboard, focus rings,
live region, high contrast, 320 px; the runtime floor check with a plain message
instead of a `<Requirements>` element; Office.js from the official CDN; the
store logo 300×300, at least one screenshot 1366×768 (the docked 512 px view
serves), descriptions, testing notes and a validators' test deck; one measured
round on web and Windows, both taken. **Mac and iPad are NOT release
requirements** — section 9 says why, and the testing notes disclose it.

**The listing NAME is settled: `SSF Slide Elements`** (2026-09-16, the owner).
It is what the manifests have always carried, so nothing in the build moved;
what changed is that it stopped being an open question.

**Recorded as a DECISION, not as a policy clearance**, and the distinction is
the point. Section 12 was read against the certification policies' section 1120
and the submission form; **nobody has read this name against policy 1100.7**,
which is the question the sibling SSF Merge was held on. If a reviewer objects
to it, the answer is a rename — not a claim that it was cleared, which is what a
record saying "checked" would licence somebody to make.

## 13. Open questions for the host

Each is written so a single round settles it. The probe that asks them is a
Script Lab snippet, not a pane: `docs/PROBE.md` says why, and how each question
is put, and the sheets are filed under `docs/host-answers/`. **All seven are
answered on the web and on Windows**, and section 15 reads every sheet.
Questions 1 to 6 were asked before the splice and the picker were built;
question 4 was settled last, on the Windows pair of 2026-09-14 against
`template/probe-comments.pptx`. Question 7 arrived after them, with the jump,
and was answered on both platforms on 2026-09-14. Mac and iPad have had no
round at all, so every answer below is borrowed there.

1. Does `insertSlidesFromBase64` accept a package pruned to one slide whose
   other parts are still present but unlisted?
2. Does an insert immediately followed by a positional delete of the slide
   before it keep the order the engine expects?
3. Does `getSelectedSlides()` name the slide the user is looking at, and does
   its position in `slides` match the position in `<p:sldIdLst>`?
4. Which read of the deck this add-in should use, `getFileAsync` or
   `exportAsBase64Presentation`, and what each drops on this host.
5. Does PowerPoint's own Ctrl+Z revert `insertSlidesFromBase64`? If it does,
   the pane's Undo must not fight it.
6. How long does `getFileAsync` take on a 50 MB deck, since the file route reads
   the whole deck for every insert, and is the floor met on iPad?
7. Does `setSelectedSlides` move the view, and does the host still answer a
   selection read afterwards? Added 2026-09-12 with the jump in section 4: that
   call is the one selection WRITE the pane makes, and it was made on
   SSF-Charts' web measurement until the sheets of 2026-09-14 answered it here,
   on the web and on Windows. On Mac and iPad the jump is still borrowed.

## 14. Build order

Everything approved ships in v1; there is no v1/v2 split. The order is
`docs/BACKLOG.md`: harvest → host probe → splice → picker → host handshake →
release and AppSource. Four deck edits belong before the library is final, and
each one ends in a fresh PDF print of both decks:

1. **English placeholder text** — DONE 2026-09-11, section 2.
2. **The rules text into notes** — already true where it was checked: the
   "gode regler for flowcharts" box is in the notes of the 16:9 deck's flowchart
   slide and on no slide. The 4:3 deck's equivalent guidance is the 276
   off-slide shapes, which are out of every element already. Nothing found that
   still needs moving; re-check before calling it closed.
3. **The collection marker line** — `SSF: ét element pr. figur` is in the notes
   of one heading slide per deck. Whether the slides under it inherit it, or
   whether the other collection headings need their own, is unsettled.
4. ~~**The Icons slide**~~ — **settled on 2026-09-16, and not by answering it.**
   It held the scales and the waste bin, "which are neither" marker, stamp nor
   flowchart shape, and the record said what was wrong without saying what the
   fix was. The waste bin left with the Flowchart shapes category and the owner
   had the scales removed the same day, so there is no Icons collection left to
   specify. Every part in the library is now a marker, a stamp or a label.

## 15. Measured and assumed

**Measured, on PowerPoint for the web, 2026-09-10.** Two pairs of answer sheets
under `docs/host-answers/`; `docs/PROBE.md` says what each reads as.

- A package pruned to **one listed slide** is accepted and lands exactly that
  slide, with the other slides' parts still in the zip. The splice is built on
  it: the user's whole deck goes back with one slide added and every other slide
  unlisted, which is the cheap route section 6 hoped for.
- An insert of the deck's **own** bytes adds **no** master under either
  formatting option (1 → 1), where a foreign fixture deck adds one (1 → 2). This
  is why the rebuilt slide is cloned from the user's own deck rather than
  assembled from the library's.
- Insert-after-target followed by a **positional delete** keeps the order the
  engine expects, and a slide the run has just added IS accepted as a
  `targetSlideId`.
- `getSelectedSlides` names the slide the user is on, and the API's order is the
  file's `<p:sldIdLst>` order.
- `exportAsBase64Presentation` **drops** the deck's comment part and
  `ppt/authors.xml`, so the insert reads with `getFileAsync` instead.
- PowerPoint's own **Ctrl+Z reverts an insert**, which is why the pane's Undo is
  one deep (section 6) rather than fighting it.
- The SPLICE ITSELF was run against this host on 2026-09-10, three inserts, the
  real engine bundled into a Script Lab snippet: a four-shape element onto a
  slide, an element carrying an embedded object and a picture (five parts
  copied) onto a slide, and one as a new slide. All three landed, and the deck
  went 3 → 4 → 3, 3 → 4 → 3 and 3 → 4. Reading the deck took 0.5 to 1.9
  seconds, the splice 76 to 179 ms, the insert 0.6 to 1.4 seconds.
- **A comment survives its slide being rebuilt, and a new slide must not carry
  one.** A modern comment on the web is anchored from the slide's own
  extension list, so a clone keeps it — which is right for "onto this slide"
  and was wrong for "as a new slide": the first round put the same comment on
  two slides. Section 6's new-slide rule now drops comments the way it already
  dropped notes, measured again afterwards on the same host.

  **Both halves of that rule were spelling-dependent, and both were wrong for
  the other spelling** (found 2026-09-22 by reading, and reproduced in
  `test/splice.test.ts` before either was changed). The clone kept whatever the
  MARKUP named, and only a modern comment is named there:

  - A **classic** `ppt/comments/commentN.xml`, which PowerPoint 2016 and 2019
    write and which any deck not yet upgraded still carries, is named by its
    relationship alone. So the rebuild dropped it, and the sentence above —
    the one thing this add-in must not get wrong — did not hold for it. The
    clone now keeps comment relationships of both spellings and leaves the
    decision about a NEW slide to the one place that makes it.
  - The new-slide rule removed the comment RELATIONSHIP and left the modern
    comment's anchor in the slide's extension list. That anchor then named a
    relationship that was gone — and, because deleting a relationship frees its
    id, the next one the insert added took it, so the anchor came out resolving
    to this add-in's own tag part. `blank()` now takes the reference out with
    the relationship, which is the discipline `clone.ts` already applied and
    the one place that deleted without it.

  Neither is measured on a host: this container has no PowerPoint. Both are
  held by the package the engine hands over, which is what the host reads.
- **"As a new slide" emptied only the placeholders it knew how to empty, and
  carried the rest of the user's content over** (found 2026-09-22 by reading,
  and reproduced in `test/splice.test.ts` before it was changed). The rule is
  "the clone keeps its placeholders, emptied; everything else goes", and
  `blank()` read a placeholder as any top-level shape carrying a `<p:ph>`. A
  table dropped into a content placeholder is a `<p:graphicFrame>` and a
  picture in a picture placeholder is a `<p:pic>`: both carry the `<p:ph>`,
  neither has a `<p:txBody>`, so the emptying pass stepped over them and left
  them whole. A new slide could therefore arrive carrying the user's own
  figures, under an element placed as if the slide were empty. A placeholder is
  now kept only when it is a `<p:sp>`, which is how an EMPTY placeholder is
  spelled — the layout's own prompt box, kept whether or not it has a
  `<p:txBody>` to empty, which `test/splice-malformed.test.ts` already pinned
  and which is what caught a first fix that removed those too. A placeholder
  spelled any other way is one the user has FILLED, so it goes with the rest of
  the content and PowerPoint draws the layout's prompt in its place. Both
  shipped routes into `target: "new"` reach it: the
  tile menu's "Insert as a new slide" and the footer's "Move to a new slide",
  the second of which is OFFERED on `held > 0`, the very condition such content
  creates. Not measured on a host, for the same reason as the pair above.
- `getFileAsync` answered a 34 KB deck in 874 ms on a healthy session and took
  40 seconds for 40 KB on one that had been through a session-timeout reload.
  Both are facts about a minute rather than about the host. Worse again on
  2026-09-11, and worth the number: a 65 KB deck **exceeded the 180-second
  budget** on a browser session that had been open for hours, and answered in
  **2.1 seconds** on the same deck immediately after a page reload. An insert
  refused for that reason left the deck untouched and said so, which is the
  behaviour section 6 asks for — but the first thing to try when a read is slow
  is a reload, not a larger budget.
- **The tags an insert writes SURVIVE `insertSlidesFromBase64`, and they are
  what "Used in this deck" and "Remove from N slides" ARE read from** — both
  shipped on 2026-09-11, on the strength of this measurement.
  Measured on the web on 2026-09-11 by reading the throwaway deck back out of
  PowerPoint with `getFileAsync` after a session of inserts: seven
  `SSF_SLIDE_ELEMENT` tag parts, each paired with its
  `SSF_SLIDE_ELEMENTS_CATALOGUE` version, referenced from the shapes through
  `<p:custDataLst>` on four different slides with every relationship intact.
  Nothing in the two features above therefore rests on an unmeasured host
  capability. The same read found **think-cell's own `THINKCELLSHAPEDONOTDELETE`
  tags sharing `ppt/tags/` with ours**, numbered around them — which is the
  case `nextTagNumber` exists for, met in the wild rather than in a fixture.
- **The deck's slide count lags an insert by up to about three seconds.**
  Measured on the web on 2026-09-11, polling `slides.getCount()` every 300 ms
  through a real insert: 2.8 seconds at the old value, then the new one. One
  round in three lost an undo to it — the restoring insert landed, the count
  read straight afterwards said it had not, and the undo stopped between putting
  the user's slide back and removing the rebuilt one, leaving six slides where
  five belonged and saying so. Every size that decides something is now read
  again on a backoff until the deck agrees.
- **The whole product was run on this host on 2026-09-11**, from the pane rather
  than from a snippet: a tile clicked in the picker, the element onto the slide
  the user was on, and Undo afterwards. The insert reported `5 → 6 → 5 slides,
  slide 1 replaced`, and slide 1 came back from the Undo with the same slide id
  and the same two shapes, by id and by name, that it carried before —
  `2:Title 1` and `9:White box, 1 large` — with the inserted triangle gone and
  the other four slides untouched. The shape inventory was read through the
  pane's own Office.js, which is the only place on the web it exists.

**Measured, on PowerPoint on Windows, 2026-09-11.** PowerPoint 16.0.20326.20132,
Microsoft 365 Current Channel, x64; the probe reports the platform as `PC` and
the host as `16.0.20326.20132`, PowerPointApi up to 1.10. Two answer sheets
under `docs/host-answers/`, and the whole product run from the pane afterwards.

**Every one of the six questions answered the way the web answered it.** The
package route is not web-specific:

- A package pruned to **one listed slide** is accepted and lands exactly that
  slide, under both prunings, with the other slides' parts still in the zip.
- An insert of the deck's **own** bytes adds **no** master under either
  formatting option (1 → 1); the foreign fixture deck adds one (1 → 2), the
  same asymmetry the web showed.
- Insert-after-target then a **positional delete** keeps the order, and a slide
  the run has just added IS accepted as a `targetSlideId`.
- `getSelectedSlides` named the slide that was clicked — slide 2 of 4 — and the
  API's order is the file's `<p:sldIdLst>` order.
- `exportAsBase64Presentation` **drops** the comment part and `ppt/authors.xml`
  here too (51 parts in, 46 out; it also drops the three `ppt/webextensions/`
  parts). So the insert reads with `getFileAsync` on Windows for the same
  reason it does on the web.
- PowerPoint's own **Ctrl+Z reverts an insert**: the first run left its tagged
  slide, one Ctrl+Z on the canvas took the deck from 5 slides back to 4 with
  slide id 260 gone, and the second run found the marker, found no tag, and
  left nothing behind.

**Every timing difference went the other way from the worry — Windows is
faster.** The two workarounds the code carries are not load-bearing here, and
stay because the web still needs them:

- **The slide count does not lag an insert.** Measured twice, polling
  `slides.getCount()` every 300 ms through an `insertSlidesFromBase64` whose
  promise was timestamped: the new count was already being returned **240 ms
  and 247 ms before the call resolved**. The web sat at the old value for 2.8
  seconds. The backoff in `src/host/timeout.ts` therefore does NOT need raising
  for Windows. (A first attempt at this measurement polled through the pane's
  own insert, which goes 4 → 5 → 4 by design; that run could not tell a lagging
  count from the real transient and was discarded.)
- **A selection read straight after an insert does not hang.** 3 ms, against
  the four seconds the pane allows a glance; 4 ms again at +1 s and +4 s.
- **Reads are fast and scale.** `getFileAsync` returned a 14.13 MB deck in
  2816 ms and 3036 ms (4.6 to 5.0 MB/s); `exportAsBase64Presentation` did the
  same deck in about 1050 ms (13.3 MB/s). The web's healthy reading was 34 KB
  in 874 ms and its degraded one 40 KB in 40 seconds. At 4.8 MB/s a **50 MB
  deck reads in about 10 seconds** on this machine, which is the first real
  evidence for question 6 — the web never had a deck big enough to ask it.
- The whole probe took **8.1 and 9.4 seconds** against a 4-slide, 14.8 MB deck,
  with no call reaching its 120-second budget. The web's second pair spent 302
  seconds and had an insert time out and land anyway.

**The whole product was run from the pane on Windows on 2026-09-11**, build
`6c906b4`, sideloaded from a trusted shared-folder catalogue:

- The line under the header names the slide you are on and follows a click.
  **Three clicks a quarter-second apart settle on the last one clicked**, twice
  over — slides 4, 2, 3 ended on "Slide 3." and slides 3, 4, 1 on "Slide 1.",
  each agreeing with what PowerPoint itself reported.
- **Four insert-and-undo rounds held**, on different slides, under both insert
  targets, each with a payload the slide did not already carry so that a
  working undo and a broken one could not look alike. Judged on the shape
  inventory — slide id, and every shape's id and name — never on the count:
  a triangle onto a slide holding `2:Title 1 | 3:Text Placeholder 2`
  (`4 → 5 → 4 slides, slide 1 replaced`); a white box onto a slide that also
  held `5:Picture 4` (slide 3 replaced); a hierarchy **as a new slide** after
  slide 2 (`4 → 5 slides`, nine shapes, removed again by Undo); and a draft
  stamp onto slide 4 (slide 4 replaced, arriving as `6:Group 4`). After every
  Undo the deck was identical to before, slide id by slide id and shape id by
  shape id.

**What PowerPoint keeps when it SAVES, measured the same way on 2026-09-11.**
Three decks built by this engine, opened through COM with no window, saved as
`.pptx`, and read back part by part. This answers a question `remove.ts` had
written down as open — the removal deliberately leaves orphans behind, and
whether they cost the user anything turns entirely on this.

- **Every orphan is dropped.** A deck after three insert-then-remove cycles
  carried 73 parts, 9 slide parts and **18 orphaned tag parts**; saved, it
  carried 39 parts, **1** slide part and **0** tag parts, and went from 55,072
  bytes back to 37,556 — the deck it started as. So the 3.5–5.5 KB an
  insert-then-remove cycle leaves in this engine's own package is reclaimed the
  moment the user saves, and the decision to leave orphans rather than walk them
  costs a user nothing.
- **Identical pictures are merged.** A deck carrying four byte-identical copies
  of one 29 KB `.emf`, one per insert, saved with **one** — 85,080 bytes to
  49,149. PowerPoint deduplicates media itself, so the sharing in `carry.ts` is
  about the package this add-in builds and ships on every insert, not about the
  size of the file the user ends up with.
- **A part name of ours is accepted, and renamed.**
  `ppt/media/ssf-93397c1904353bcf-29552.emf` opened without a repair prompt and
  came back as `ppt/media/image1.emf`. The content-addressed name lives only
  until PowerPoint rewrites the file.
- **Tag parts that are still REFERENCED survive a save** — four inserts, four
  tag parts in, four out, with the shapes still carrying their `<p:tags>`. The
  web measurement above says the tags survive `insertSlidesFromBase64`; this
  says they also survive an ordinary save on the desktop.

**The whole product was run from the pane on PowerPoint for the WEB on
2026-09-13**, build `931bc1d` — the code the pane itself prints under its
header — with `manifest-prod.xml` sideloaded into `template/validators.pptx` on
OneDrive. This is the round the backlog had been asking for, and it settles
three separate things.

- **"Remove from N slides", end to end, which no round had exercised.** The
  Confidential stamp onto all three slides (`3 → 4 → 3 slides, slide N
  replaced`, once per slide); "See what this deck already uses" answering
  *Confidential stamp — slides 1, 2 and 3*; the question *Take Confidential
  stamp off slides 1, 2 and 3? The pane cannot undo this.*; and then **`Removed
  from 3 slides.`** Checked by READING THE DECK AGAIN rather than by believing
  the footer: no rows, no Remove buttons, "nothing from the library is in this
  deck yet", and the deck still three slides. The read behind it settled in 12
  seconds; the three confirmed cycles finished inside 4.
- **The jump moves the deck, and does NOT wedge the host.** From `Slide 1 of 3`,
  clicking the `2` in "Used in this deck" left PowerPoint on `Slide 2 of 3`, and
  an insert straight afterwards ran normally (`3 → 4 → 3 slides, slide 3
  replaced`). That was the PRODUCT's behaviour rather than an answer sheet; the
  sheet arrived the next day and agrees with it (below). `CLAUDE.md`'s rule
  about `setSelectedSlides` wedging the web host's selection subsystem is a
  sibling's recording that neither the round nor the probe reproduced.
- **"Move to a new slide"** took a white box that had landed on slide 2 and made
  it its own slide: `3 → 4 slides.`
- **The element stamped was `confidential`**, one of the 23 that arrive as a
  group the owner drew — the class the ungroup fix had missed — so that fix is
  host-verified rather than engine-verified.
- **One refusal, and it was the right one.** A second insert fired immediately
  after the first answered *PowerPoint would not say which slide you are on, so
  nothing was inserted. Click a slide and try again.*, and nothing landed. A
  keypress to move slide and a retry worked. The selection is not always
  readable in the moment after an insert has replaced a slide; the pane neither
  guessed at it nor dropped the insert silently.

**Measured, on PowerPoint for the web, 2026-09-14 — the probe's own sheets, and
the one that closes question 7.** A pair under `docs/host-answers/`
(`…T09-26-03-063Z.json` and `…T09-38-04-756Z.json`), run from Script Lab against
`template/validators.pptx` and read by `scripts/read-answers.mjs`.

- **Question 7, answered here for the first time on any platform.**
  `setSelectedSlides` moved the view to the slide asked for in **1,146 ms**, put
  the previous selection back, and **the next selection read answered in 567
  ms** — the host keeps working afterwards. Measured twice, 12 minutes apart
  (998 ms / 590 ms on the first sheet), which is what makes it a measurement
  rather than an anecdote. Section 13's paragraph on the jump stops being
  borrowed from SSF-Charts for the web.
- **Question 5 the same day.** The slide the first run left behind was gone at
  the second run's start, with its marker still in place — so PowerPoint's own
  Ctrl+Z reverts an insert, and the pane's Undo must not fight it. This is the
  desktop answer confirmed on the web by the instrument rather than by eye.
- Questions 1, 2 and 3 answered as before: both prunings land exactly one slide,
  insert-then-positional-delete keeps the order, a just-added slide is accepted
  as `targetSlideId`, and the API's order is the file's `<p:sldIdLst>` order.
- **Question 4 was NOT asked**, and the sheet says so rather than passing: the
  validators' deck carries no comments and no `ppt/authors.xml`, so the export
  had nothing to drop. The 2026-09-10 and 2026-09-11 sheets were still the
  answer, and a deck with comments is what a re-run needed — which is what the
  pair below was run on.
- **Question 6 is the uncomfortable one.** `getFileAsync` took **31,755 ms for
  0.05 MB**, and 14,523 ms at the end of the same run, against
  `exportAsBase64Presentation` at 550 ms for the same deck. That is not a size
  cost — it is a 50 KB deck — so section 13's sixth question cannot be answered
  by extrapolating from megabytes, and the read the pane pays for "Used in this
  deck" is bounded by something other than bytes. The product round the day
  before saw 12 seconds for the same work, so the number moves a lot between
  runs. **Nothing here is a reason to change the engine yet; it is a reason not
  to trust a single timing.**

**A pair on PowerPoint for Windows, 2026-09-14** — host `16.0.20326.20144`,
`PowerPointApi` through 1.10 — run on `template/probe-comments.pptx`, the deck
authored for exactly this: three slides, a modern comment on slide 3 and the
`ppt/authors.xml` that comes with one. Sheets
`2026-09-14T13-58-44-243Z.json` and `2026-09-14T14-08-57-305Z.json`.

- **Question 4 is answered, for the first time on any platform by this
  repository's own instrument on a deck that could answer it.**
  `exportAsBase64Presentation` handed back **43 parts where `getFileAsync` gave
  48**, and the five it did not carry over were
  `ppt/comments/modernComment_102_A61DA0E.xml`, `ppt/authors.xml`, and the three
  `ppt/webextensions/` parts. So the drop SSF-Merge's sixth sheet found on the
  web reaches the presentation-level call on Windows too, and the engine reading
  with `getFileAsync` is measured here rather than borrowed. Both runs agreed,
  ten minutes apart.
- **Question 6, against the web's anomaly of the same day.** `getFileAsync` read
  the 0.04 MB deck in **62, 76 and 61 ms** across the pair, and
  `exportAsBase64Presentation` in **28 and 30 ms**. The web had taken 31,755 ms
  for a deck of the same size hours earlier. Whatever that was, it is not a
  property of the call, which is the second reason not to answer question 6 by
  extrapolating.
- **Question 5 on Windows, with the keystroke measured on both sides.** The
  first run left its tagged slide (3 → 4, confirmed through COM), one Ctrl+Z
  took the deck back to 3, and the second run found its own marker from the
  first and the tagged slide gone. Ctrl+Z reverts an insert here as it does on
  the web.
- **Question 7 on Windows**: `setSelectedSlides` moved the view in **7 ms**, put
  the previous selection back, and the next selection read answered in 58 ms.
  The jump is now measured on both platforms this project ships to and borrowed
  on neither.
- Questions 1, 2 and 3 as everywhere else: both prunings land exactly one slide,
  insert-then-positional-delete keeps the order, a just-added slide is accepted
  as `targetSlideId`, and the API's order is the file's `<p:sldIdLst>` order.
  Inserting a slide on the deck's own master added **no** master, under
  `KeepSourceFormatting` and `UseDestinationTheme` alike.

**Two rules about the INSTRUMENT, both paid for on 2026-09-14.**

- **An imported Script Lab snippet does not run until it is TRUSTED, and the
  failure is silent where you are looking.** The runner's console says "There
  are no logs to display" and the deck does not change; the reason sits inside
  the runner's sandboxed `user-snippet` iframe, reading "Untrusted Snippet — in
  order to run … you must first trust it in the editor". The editor shows a
  "Would you like to trust this snippet?" bar with a **Trust** button. This cost
  about fifty minutes across the web and Windows, and on the web it was
  misdiagnosed as a wedged runner. **Read the iframe, not the console strip.**
- **The ribbon's Run does nothing once the runner is open**, on Windows as on
  the web. The runner's own "Last updated … ago" label is a BUTTON and is the
  re-run, which is what the second of a pair needs.

The rest of this section is about the DECKS and the print rather than Office.js.

- **Neither committed deck opened, until the slashes went in.** Both were
  refused with "PowerPoint found a problem with content", offering Repair, in
  the UI and through COM alike (`Presentations.Open` raised `0x808D1001`).
  Section 3 carries the cause and what was done. The decks were not edited to
  FIND it: every experiment ran on copies, and both files were byte-identical to
  `HEAD` until the fix was made deliberately, on its own, afterwards.
- **The fix cost the catalogue nothing.** Harvesting the corrected decks
  produces a catalogue index byte-identical to the one harvested from the
  originals, version `1641fe687794` either way — predicted before the change was
  made and confirmed against the decks now committed. So it did not invalidate
  `public/catalogue/catalogue.json` and needed no re-harvest.
- **`ExportAsFixedFormat` cannot be called through automation on this build.**
  Every arity, from PowerShell and from VBScript alike, raises
  `DISP_E_TYPEMISMATCH`; the method is present on the type and refuses to bind.
  So a print cannot be driven from code with its options named. The committed
  prints went through the **File → Export dialog by hand** instead, which is
  what section 3 says to do and is the only route on this build where the
  switches can be both set and read back. `SaveAs(path, ppSaveAsPDF)` works and
  takes PowerPoint's defaults; on these two decks it is equivalent, since they
  carry no comments, no ink and no hidden slides, so every non-default switch
  the print needs is a no-op — but it cannot show which options were used, only
  what came out.
- **Both decks printed correctly on 2026-09-11, through the dialog.** This is
  the record of that round, not of the committed prints: the pair described here
  was superseded on 2026-09-16 by a COM re-print of both decks, after the Icons
  slide left the library, and section 3 carries the committed figures (109 and
  107 pages). Taken through
  **File → Export → Create PDF/XPS → Options** from the decks at their committed
  paths, with Range all, Publish what Slides, Frame slides off, Include hidden
  slides on, Include comments off, Include ink off, and Optimise for Standard —
  each setting read back off the dialog before publishing, because the Options
  dialog resets to its defaults between presentations and does NOT remember what
  the previous publish used. The 16:9 deck gives 110 pages for 110 slides at
  960×540 pt (2,220,422 bytes); the 4:3 deck 108 pages for 108 slides at
  720×540 pt (2,432,141 bytes). Page counts were read from the PDF page tree and
  again through `Windows.Data.Pdf`, against slide counts read from the decks'
  own `ppt/slides/slideN.xml` parts and `<p:sldId>` entries. Sampled pages are
  rendered slides with no frame, and each one's slide-number footer matches its
  page number. Both decks were byte-identical to `HEAD` afterwards.
  An earlier pair of prints, taken from corrected copies before the fix was
  committed, was deliberately NOT committed: a print whose provenance does not
  match the deck the harvest checks it against would pass the slide-count check
  while being cut from the wrong file.

Measured against the committed catalogue on 2026-09-23: the 16:9 library has
106 named elements, 10 of them parts of two collection slides (103 and 104),
twelve runs of sizes over 45 members, 62 whole-slide elements carrying no group;
the stamps are rotated 29° and 35°; a table's frame is narrower than the table
PowerPoint draws. The 4:3 library is the same 106, 10 and 12, over collection
slides 102 and 103. The figures this line used to give — 118 elements, 21 parts,
four collection slides, 42 without a group — were taken before the removals of
\#105 and \#107 and were never restated; 118 was the count in
`template/names.en.json`, not in the catalogue, and nothing here reproduces 42
under any definition of "group" the record states.

**Borrowed, dated, and read back rather than trusted:** `setSelectedSlides`,
the one selection write this add-in makes (the jump in section 4). No sheet of
this repo's measures it. SSF-Charts' self-test archive does, on PowerPoint for
the web: 2,429 selection-ladder rungs over about 347 rounds between 2026-08-13
and 2026-09-04, every rung answered, none silent, none refused; its wedge
finding is `setSelectedShapes([id])` on build `55011a3`, and `setSelectedSlides`
went silent only downstream of that. Windows and Mac: no sibling has measured
the call; the tracker has office-js#3552 (desktop throws while the notes pane
has focus). The pane therefore reads the selection back after every jump and
claims nothing it did not see. Probe question 7 was the arm that would turn this
paragraph into a measured one, and it did, on the web on 2026-09-14: the call
moved the view to the slide asked for in 1,146 ms, put the previous selection
back, and the next selection read answered in 567 ms. **On the web this is now
measured here rather than borrowed.** Windows and Mac are still the sibling's
and the tracker's.

**A run of several cycles, and Stop, measured on Windows on 2026-09-23** against
`090e97c` — the pane read its own `data-build` off the document before anything
was run, so the build under test is not inferred from the deployment. Host
16.0.20326.20158. Every figure below is a COM inventory of the deck, slide id by
slide id, either side of the action; the pane's own sentence is quoted as a
claim and the inventory is what settles it.

- **A stamp across three selected slides landed on exactly those three.** Each
  took a new `SlideID` (257 → 262, 261 → 263, 260 → 264), which is three
  insert-then-remove cycles rather than three shapes added in place, and the
  three unselected slides kept their ids and their shape counts. The footer
  said "Stamped 3 slides" and three is what the deck gained. This is the
  sequence the paragraph below used to list as assumed.
- **Stop stops, and the rest really are as they were.** 59 slides selected,
  Stop pressed at the fourth cycle: the pane said "Stopped after 5 of 59
  slides. The rest are as they were", and the deck agreed exactly — five slides
  changed, **55 untouched by id and by shape count**, the deck still 60 long.
- **A cycle costs about 150 ms here.** From the pane's own progress labels,
  cycles 1 to 5 of that run landed at 341, 557, 689, 820 and 943 ms; a
  three-slide run is over in **392 ms** end to end. That prices the Stop
  control honestly: on a small deck on Windows there is well under a second in
  which to press it, and it earns its place on big decks and on the web, where
  a cycle is seconds rather than milliseconds.
- **The two libraries now agree, on a real host, on both slide sizes.** A 16:9
  deck and a 4:3 deck opened against the same build offered the same ten
  categories in the same order with the same counts, a bare `73` on the count
  line, and `KPI definition` under One-page templates on both.

**A transition survived "as a new slide", and the gate could not see it.**
Measured on Windows on 2026-09-23 against `380700f`, on a deck whose slide 2
carried both a transition and an entrance animation: the new slide came back
with **no `<p:timing>`** — correct — and **with the transition**. The reason is
the spelling. `blank()` walked the direct children of `<p:sld>`, and PowerPoint
writes a modern transition wrapped:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="p14"><p:transition p14:dur="2000">…</p:transition></mc:Choice>
  <mc:Fallback><p:transition>…</p:transition></mc:Fallback>
</mc:AlternateContent>
```

so the direct child is `mc:AlternateContent` and the transition is a
grandchild. The suite's fixture used the bare spelling, which is why every gate
was green over it — a fixture that cannot fail the way the real thing fails.
`dropTimingAndTransition` now handles both, drops the wrapper only when
emptying it leaves nothing, and `test/splice.test.ts` holds both spellings in
both directions. A wrapped `<p:timing>` is covered too; that half is defensive
rather than measured, and the code says so.

**One `Ctrl+Z` does not take back an "onto this slide" insert; two do.**
Measured the same day, and it does not contradict question 5 above — that
answer is about a BARE `insertSlidesFromBase64`, which is one operation. This
add-in's insert is two: the insert, then the positional delete of the slide it
replaced. So PowerPoint's undo stack holds two entries for it, and one press
reverts only the delete:

| | deck |
| --- | --- |
| after the insert | 256, **259**, 258 — three slides |
| after one Ctrl+Z | 256, **257**, **259**, 258 — **four slides** |
| after two Ctrl+Z | 256, 257, 258 — as it started |

A user who presses it once is left holding both their original slide and the
rebuilt one.

**Measured again on 2026-09-23 against `faf101e`**, on a deck whose every slide
carries its own label, which is what makes the second reading conclusive: after
one press the deck held **two slides both labelled `SLIDE-05`** — the user's
original (2 shapes) and the add-in's rebuilt copy (3 shapes) — and after the
second it was back to one, byte-identical to before the insert. Two decks, two
readings, so `docs/MANUAL.md` now says so.

**A mid-run reorder: the guard fires, and nothing is lost.** Same round, 39
slides selected for a stamp and one slide dragged from position 30 to position
5 while the run was going. The run stopped and said *"Stamped 26 of 39 slides,
and the deck has a slide too many: the copy was made but the original could not
be taken away."* The deck bore that out exactly: **no label lost**, one label
(`SLIDE-28`) present twice — the original beside its rebuilt copy — and the
moved slide left unstamped. That is `stillThere` refusing the positional
delete, which is what section 6 asks of it.

**That was the guard's limit, and it has since been closed.** The guard compared
the slide it aimed at with the slide now at that index, INSIDE one cycle. A
reorder that landed BETWEEN cycles moved nothing during a cycle, so nothing was
caught and the run carried on against indices that had shifted under it:
measured on 2026-09-23 on a deck of identical slides, a run reported "Stamped 59
slides" where 58 slides gained one, and the slide that had moved was the one
without it. Nothing was lost — the deck kept its length and its slides — but the
count was one too high and one selected slide was silently skipped.

**Both runs now carry slide IDS.** `slideIds` reads the deck's ids in order,
`indexOfSlide` answers where one of them currently sits, and every cycle asks it
twice: once to find the slide it is about to rebuild, and once after the insert
to find the original it is about to remove. The positions handed to the ENGINE
are unchanged and must be — they index this run's own byte snapshot, which does
not move — but nothing touches the live deck by a position it was given in
advance any more.

Three behaviours fall out of it, and the third is why this was worth doing to a
destructive loop:

- a slide that **moved** is followed, so an ordinary drag costs the user
  nothing where it used to strand a copy they had to find and delete;
- a slide that has **gone** answers nothing, so the cycle is skipped, `done`
  does not count it, and no position is deleted in its place;
- the count at the end is the number of slides that actually changed.

`test/pane-wiring.test.ts` drives a reorder from inside a cycle and between two
cycles, over a fake deck that has a real order to be reordered; both cases were
run red against the positional code first, on the assertion about which index
was deleted.

**Verified against a real PowerPoint on 2026-09-23**, against `bd91527`, on the
40-slide deck whose every slide carries its own label. The property the fix
promises does not depend on when the drag lands, which is what makes it
checkable without hitting a 150 ms window:

- **A stamp over 39 selected slides, with a slide dragged from position 30 to
  position 5 while it ran.** Thirty-nine labels gained **exactly one** stamp —
  every selected slide, the dragged one among them. `SLIDE-01`, the one slide
  not selected, was untouched; nothing gained two; nothing was lost; and the
  pane's "Stamped 39 slides" matched the deck. The same scenario against the
  positional code reported 59 where 58 had gained one.
- **A removal over the same 39 slides, with a slide dragged from 28 to 3 while
  it ran.** All thirty-nine stamps gone, the deck still forty slides, **no
  label lost**. This is the path that takes content out of a deck.
- **Regression in the same round**: insert onto a slide (`40 → 41 → 40`); the
  pane's Undo returning the deck **identical to baseline by id, label and shape
  count**; Stop at cycle 6 of 39 leaving exactly `SLIDE-02`…`SLIDE-07` changed
  and 34 untouched; "Used in this deck" naming exactly those six; and a removal
  putting every slide back to the shape count it started with.

**What the round did NOT settle.** The skip path — a slide DELETED while the
run is going, which must be passed over rather than deleted into — was not
cleanly measured. Two attempts were defeated by timing: the first deletion
landed after the run had already reached that slide, and the second landed
inside a cycle's count confirmation, where the deck shrinking by one is
indistinguishable to the pane from its own delete having failed. The run then
stopped and said "the deck has a slide too many", which is the conservative
answer and tells the user to look — but the deck had no extra slide, so the
sentence names the wrong cause. Nothing was lost in either attempt, and the
unit cases in `test/pane-wiring.test.ts` do cover the skip; what is missing is
the host's own word for it.

**Assumed**: every host fact above on **Mac and iPad**, where no round has been
run — and, from 2026-09-12, where none is planned before release: the owner has
neither device, so the validators' report is the first measurement for both
(section 9). A SEQUENCE of insert-then-remove cycles is **no longer assumed in
either direction**. Stamping: the round of 2026-09-23 above put a stamp on three
slides and then on fifty-nine, and the deck was read slide by slide either side
of both. Removing: the round against `bd91527` the same day took a stamp off 39
slides with a slide dragged mid-run, and the deck came back forty slides with
every stamp gone and no label lost. What is still unmeasured on a host is a
slide DELETED mid-run, which the section above records. Also assumed: the two-second budget in section 11, and the certification
reading in section 12. Windows is no longer assumed — the section above is its
round — but one Windows machine is one machine, and the "50 MB in about ten
seconds" figure is an extrapolation from a 14 MB deck, not a measurement of a
50 MB one.

## 16. Decisions log

All 2026-09-08, all the owner's, in the order they were taken.

| Decision | Verdict |
| --- | --- |
| Package route, one insert, insert first then remove | approved |
| Elements from the owner's deck; collection slides split one element per shape | approved |
| Stamps land top-right, markers and flowchart icons at the cursor, centre when nothing is selected | approved |
| Recent kept (last six); collection marker as a notes line; names per locale, English only; colours as a switch; categories collapsed with Open all | approved |
| No Danish tooltip; everything English; format-only elements authored in the other deck (both v2 decks approved) | approved |
| Publishing on AppSource, publisher StruktureretSundFornuft ApS as a company, web + Windows + Mac at launch | approved |
| Sizes with a stepper; keyboard operation; preview card; favourites; parts ignore the target and "Move to a new slide"; options behind the gear with sticky search and headers; one insert at a time | approved |
| Authored tag vocabulary; a "when to use" sentence per element; a "New" chip | rejected |
| No "family" word in the pane; collection split by slide title with an Icons slide; tags one line | approved |
| Stamp every selected slide; remove all of a kind; fit to the destination layout; insert as one group; search matches Danish; any slide size; the pane reopens where you left it; first-run guide | approved |
| Swap the element on the slide; Enter inserts the top hit | rejected |
| Search bar says "Search"; the SSF tick as the one orange; real renderings on the demo slide; stamp and flowchart cuts fixed | approved |
| Marker wraps the selected shape; right-click for the other target; Again; empty placeholders removed; "Used in this deck"; forgiving search; category chips while searching; Undo ten deep with Ctrl+Z; touch; high contrast; catalogue page; Report a problem; all library text English in v1; docked preview at 512 px | approved |
| Several new slides in one pick; authored sets; version-aware elements; names and alt text on inserted shapes; density toggle; go-to-category dropdown | rejected |
| iPad included at launch; the owner cannot measure it, so the first measurement is a borrowed iPad, a device cloud or the validators | approved |
| The design record into the repo as a docs-only PR; everything in v1; host probe before the splice and the picker | approved |
| The host probe as a Script Lab snippet with every verdict a tested pure function, rather than a probe pane that would need hosting and a re-install first; one slide left behind on purpose for the Ctrl+Z question | decided in the build, 2026-09-08 |
| The splice sends the user's WHOLE deck back with one slide added and every other slide unlisted, rather than removing the others properly — licensed by probe question 1, and the rebuilt slide is cloned from the user's own deck so the insert adds no master (question 1b) | measured in the build, 2026-09-10 |
| Undo one deep instead of ten, because putting a replaced slide back means holding a copy of the deck, and PowerPoint's own Ctrl+Z was measured to revert an insert | decided in the build, 2026-09-10 |
| A tile draws the element's landing as a diagram rather than a picture, because the PDF prints section 3 cuts previews from are not committed yet | decided in the build, 2026-09-10 |
| The pane fetches its catalogue from its own origin, so SECURITY.md's "no network calls" became the stronger "sends nothing anywhere": one named file, GET only, no absolute URL, no request options | decided in the build, 2026-09-10 |
| The probe tells a second run from a first by a marker in the document settings, written before the slide it leaves so the user's Ctrl+Z still lands on the insert; the slide alone could not, because Ctrl+Z is what removes it (web round, 2026-09-10) | decided in the build, 2026-09-10 |
| The colour switch resolves the library theme at HARVEST time, one map per size, and refuses a deck with two themes — rather than carrying the theme part into the user's package, which would make every insert add a theme the user did not ask for | decided in the build, 2026-09-11 |
| The store listing is written into the repo (`docs/LISTING.md`) and held to the manifests by a test, rather than typed into Partner Center at submission time — and the screenshot, the validators' test deck and the listing name are left as the owner's, because a composited screenshot or a deck built by this repo's own code would be a picture of something that does not exist | decided in the build, 2026-09-11 |
| The 4:3 deck was re-themed to the 16:9 deck's colour scheme rather than the other way round: the 16:9 deck is the owner's own 2021 template and the 4:3 deck's palette came from the company it was authored at in 2013. The change is the deck's, so the committed print no longer belongs to it and the print gate says so until it is re-printed | decided in the build, 2026-09-11 |
| The gear's two external links carry an allowlist of three values and open the site the PANE was served from, as buttons rather than anchors — and the support page reads the same allowlist back, so a crafted link can put a build code on that page and nothing else | decided in the build, 2026-09-11 |
| "Used in this deck" reads the deck when the user asks rather than when the pane opens, and its slide numbers are text rather than links — the read is the sixth open question's unmeasured cost, and the jump is a host call no round has made | decided in the build, 2026-09-11 |
| The jump built after all, on `setSelectedSlides`, on SSF-Charts' dated web measurement (2,429 rungs, 2026-08-13 to 2026-09-04, none silent) rather than this repo's, with the selection read back on every click so the pane claims only what it saw; text below PowerPointApi 1.5; `goToByIdAsync` rejected on office-js#2595, #2631 and #567; probe question 7 added to measure the call on the next round | owner: build what can be built on the siblings' and the tracker's research, 2026-09-12 |
| Right-click opens on `contextmenu` rather than a mouse-only handler, offers nothing on a part, and anchors to the top of its own tile rather than to the pointer — so the keyboard reaches it, the pane never promises a landing the engine does not do, and no coordinate reaches the state | decided in the build, 2026-09-11 |
| The preview card's grey boxes are a snapshot stamped with the slide it was read from, drawn only while the user is still on that slide, rather than a read per slide change — and they are read out of the FILE, so they need no host capability the insert does not already use | decided in the build, 2026-09-11 |
| "Remove from N slides" asks before it removes, runs one confirmed cycle per slide, stops at the first step it cannot verify and says how far it got — and reaches only shapes this add-in tagged | decided in the build, 2026-09-11 |
| A rectangle measured on the user's slide is divided by the USER's slide size, never the library deck's — they differ exactly on a borrowed library, which is the case that would never have shown up in testing | fixed in the build, 2026-09-11 |
| A removal re-reads which slides carry the element from the deck it is about to change, rather than trusting the list the question was asked about | fixed in the build, 2026-09-11 |
| Anything anchored to a tile — the right-click menu, the question before a removal — is keyed by TILE rather than by element, because one element is drawn in up to three lists at once | fixed in the build, 2026-09-11 |
| Base64 is converted by the platform rather than by JSZip, on both sides of the package layer — measured as the single largest cost in the insert path, 4.1 s to 0.23 s on a 45 MB deck | fixed in the build, 2026-09-11 |
| A carried MEDIA part is named after its own content — a fingerprint of its bytes and their length — so the second insert of the same element finds its picture already in the package and points at it instead of copying it again. Measured: four inserts of `markeringer-1` left four byte-identical copies of one 29 KB `.emf` and cost 11.6 KB each; they now cost 1.9 KB and leave one. It is the PACKAGE this add-in ships on every insert that shrinks, not the user's saved file — section 15 measured PowerPoint merging identical pictures itself on save. Derived from the content rather than found by searching, because a search means decompressing every picture in the user's deck on every insert. Media only: a chart or an embedded workbook is a document, and two charts sharing one workbook would mean editing one edits both | fixed in the build, 2026-09-11 |
| An insert stamps the shapes INSIDE any group that lands, as well as the group, because ungrouping is one gesture and it destroys the group and its tag together — measured: a five-shape element ungrouped went from one use to not in the deck at all. The reader stops at a tagged shape, so while it is a group the answer is unchanged; a shape with no `<p:nvPr>` is skipped rather than refused, since failing an insert that works today is the worse trade. About 250 bytes a shape The first version of it asked whether THIS code had made the group, which left the 23 elements that are drawn as a group already — the stamps among them — behaving the old way; a sweep over every element in both libraries is what found that. | decided in the build, 2026-09-11 |
| The tag sweep goes into the user's own groups rather than reading the top level of the slide only: grouping an element with a shape of your own is one gesture, and it made "Used in this deck" answer that the element was not in the deck while it sat on the slide. A removal takes the tagged shape out of that group, leaves the user's own shape beside it, and takes the group too only when the removal is what emptied it | fixed in the build, 2026-09-11 |
| "Move to a new slide" is an undo followed by a second insert rather than a third operation, so it inherits the positional count-checked undo and the delta-proving insert and adds no new failure of its own; it stops at a failed undo rather than leaving a second copy; and "a slide that already had content" is `contentCount`, which does not count the slide's own title or an empty placeholder the insert removes — reported by the splice out of bytes it already holds, so the offer costs no second deck read | decided in the build, 2026-09-12 |
| The pane's per-deck memory is keyed on a hash of `Office.context.document.url` rather than on the URL, with the query and fragment dropped so the key survives a session; no URL — an unsaved deck — falls back to the one per-machine bucket rather than forgetting; favourites and the first-run flag stay per machine, and Recent goes per deck | owner: key it on the deck's URL, guarded, 2026-09-12 |
| The scroll position is kept OUT of the pane's state — an offset in the state is a re-render per scroll event — and written on a 250 ms trailing timer; it is put back once, after the first draw that has tiles in it, and stood down by any scroll the user makes first | decided in the build, 2026-09-12 |
| The owner has neither a Mac nor an iPad, so a Mac round stops being a release requirement and joins iPad's position: the validators' report is the first measurement for both, the pane degrades through the runtime floor check rather than a manifest requirement set, and the testing notes disclose it — derived by `test/listing.test.ts` from which answer sheets exist, so a filed Mac sheet forces the disclosure to drop Mac | owner: no access to a Mac or an iPad, 2026-09-12 |
| The Scales comes out of both library decks — a SHAPE deleted from the Stamps and labels collection slide, not a slide deleted, because seven other elements share it. Identified by the box the catalogue records rather than by an index or a shape name, and refused unless exactly one shape matched. It was the last of the _Icons_ pair after the waste bin left with Flowchart shapes, and the only part whose key was Danish — `Stempler og lignende 1`, the numbered fallback for a part with no text of its own. The libraries go to 106 elements | owner: remove the Scales, 2026-09-16 |
| The Flowchart shapes category and its ten part elements come out of both library decks: one slide each (105 at 16:9, 104 at 4:3), carrying the category's own heading, so the category goes with it. The libraries drop from 117 elements to 107 and the 16:9 deck from twelve categories to eleven. A deck that already uses one keeps it — they are ordinary shapes once inserted — and "Used in this deck" still names it as an element from an older library rather than dropping the row | owner: delete Flowchart shapes, 2026-09-16 |
| The build stamp moves off the header and onto the root element as `data-build`, rather than being deleted or painted out of the screenshot: the AppSource image may not be retouched, and the stale-cache diagnostic it exists for is worth keeping wherever it can be read — devtools, a support request, a driver over CDP — while being invisible to a user and to a capture. "Report a problem" still prefills it | owner: take it out of the listing shot, 2026-09-15 |
| Search marks the words it matched inside the name on the tile and the preview card, and NOT on "Used in this deck" (which the search does not filter) or "Did you mean" (which can never hold a match) | owner: approved with the plan, 2026-09-23 |
