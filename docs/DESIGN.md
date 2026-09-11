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
   and offers **Undo** and **Again**.

## 2. What an element is

- **A library slide is one element**: everything on it that is not layout
  chrome (title, footer, slide number). Grouped or not: 42 of the whole-slide
  elements in the 16:9 deck have no group at all (tables, matrices, the
  one-pagers), so the harvest cannot depend on grouping.
- **A collection slide yields one element per top-level shape, and the slide's
  title is their category.** The owner's three collection slides give
  _Markers_, _Stamps and labels_ and _Flowchart shapes_; a fourth, _Icons_,
  holds the scales and the waste bin, which are neither. A collection slide is
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
- **An element that comes in several sizes is one tile with a stepper.** A run
  of elements that differ only by one count (process flows with 1 to 6 boxes,
  hierarchies with 2 to 5 boxes, matrices with 2 to 5 rows, and so on) is
  derived from the name pattern: a number that stands alone before a word,
  never inside "2×2" or "1-2-3", "box" the one irregular plural, and the
  key-figure flows one run by an explicit rule. The stepper names what it
  counts ("boxes 1 2 3 4 5 6"). The pane never uses the word "family"; that
  word is for this document. The 16:9 deck has twelve such runs, so 117
  elements show as 84 tiles: the runs cover 45 elements, so 117 − 45 + 12.
- **Categories** are the heading slides (a slide with a title and no content).
  **Tags** are derived from names; there is no authored tag vocabulary.
- An off-slide shape (x at or beyond the slide's right edge) is never part of
  an element.

## 3. The library

- **Two decks, one per slide size**, authored by the owner and committed under
  `template/`: `library-16x9.pptx` and `library-4x3.pptx`. Both carry the same
  117 keys, and the harvest refuses a key that is not in both, so no element
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
  in the two sizes**: the same box was Office orange at 16:9 and light blue at
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
  **Both prints are committed**, `template/library-16x9.pdf` (110 pages,
  2,218,863 bytes) and `template/library-4x3.pdf` (108 pages, 2,307,900 bytes),
  re-taken on 2026-09-11 after the placeholder text went English, from the decks
  as they stand. Section 14's remaining deck edits have not happened, so both
  need retaking again after that pass, and the cutting code does not exist yet
  either — nothing downstream is holding a stale cut. `*.pdf` is
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

- **Header**, blue, with the SSF tick above the name and the build stamp at the
  right. The tick (the small orange bar above a heading) is the design system's
  signature and the one orange in the pane: orange in one place at a time, or it
  stops meaning anything. Hover rings, the star, the ghost frame and the coach
  marks are blue or navy. Under the header, only when the deck's size is
  neither 16:9 nor 4:3: "4:3 library, scaled to A4 slides".
- **Search**, one field, placeholder "Search", `/` focuses it, Esc clears it.
- **Tags**, one line until opened with the chevron at its right; a picked tag
  moves to the front so it stays visible when the line is closed. The chevron
  only shows when there is a second line.
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
  - **The slide numbers are text, not links.** The jump is a host call no round
    has made, and a control that might do nothing is worse than a sentence that
    says where the element is. It comes back with the round that can verify it.

  Three states, and the pane says which: never asked, asked and empty ("Nothing
  from the library is in this deck yet"), asked and answered. An id the current
  catalogue cannot name is kept and shown as "an element from an older version
  of the library" — eleven ids changed when the part keys were translated
  (section 2), and a row silently dropped would make the deck look emptier than
  it is.
- **Favourites** (a star in the tile's corner, remembered per machine) and
  **Recent** (the last six inserts) sit above the categories.
- **Categories** start collapsed, with "Open all" beside the count; a search or
  a tag opens what it finds; category headers stick while scrolling. Tiles are
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
  **The grey boxes for what the destination slide already has are NOT built.**
  They need the pane to read the slide's shapes, which is host work that cannot
  be verified without a round against a real PowerPoint. The ghost frame for the
  landing is drawn from the catalogue and needs no host at all, so that half is
  there.
- **Footer**: the last outcome with the measured slide count, then the actions
  (**Move to a new slide** when a whole-slide element landed on a slide that
  already had content, **Again**, **Undo (n)**), then a line with the current
  settings that opens the gear.
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
  slide in one insert.

## 6. Inserting

- **One insert at a time.** The pane locks while an insert runs, the tile says
  "Inserting…", and the footer reports the deck's slide count before and after:
  "12 → 13 slides" for a new slide, "12 → 13 → 12 slides, slide 4 replaced" for
  an insert onto the slide, because in the file that is insert a rebuilt copy
  after the slide and remove the original. Two inserts 0.4 s apart killed a
  sibling's tab; the lock is that rule made visible. The delta is the evidence,
  never the absence of an error.
- **Undo** goes **one deep**, positional and count-checked, never by id.
  **Again** repeats the last insert on the current slide or selection.

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
category and the tags. Matches are highlighted in the names. While searching,
the categories that have hits appear as chips with counts to narrow the search,
and a sized tile greys out the counts that do not match. A query with no hits
offers "Did you mean …" from the nearest names. Not doing: Enter inserting the
top hit.

## 9. Accessibility and platforms

- **Keyboard**: Tab into the list, arrows move between tiles, Enter inserts,
  `/` focuses search, Esc closes a menu, the preview or the search in that
  order. Focus draws the same ring as hover. A live region announces every
  outcome.
- **Touch**: the first tap on a tile shows the preview, the second inserts;
  nothing depends on hover.
- **Windows high-contrast mode**: the pane follows forced colours; rings, chips,
  tiles and the tick stay visible.
- **Office themes**: light and dark, from `Office.context.officeTheme`.
- **Pane sizes**: usable at 320 px, three columns from 400 px, docked preview
  from 512 px, and the list stays scrollable with the footer visible in a
  400 px tall pane. `pane-shots` measures these states.
- **Platforms**: PowerPoint on the web, Windows, Mac and iPad. The manifest
  cannot exclude iPad, so iPad is in whether measured or not; the owner has no
  iPad, so the first iPad measurement is a borrowed device with the manifest
  sideloaded through OneDrive, a real-device cloud if PowerPoint installs
  there, or the validators' report. Until then the pane degrades honestly on
  iPad through the floor check, and the testing notes say the publisher has not
  measured it.

## 10. Loading and failure states

The pane never shows an empty or broken screen, and every message says what
happened and what to do.

- "Loading the library…" with the header and build stamp visible at once.
- "The library did not load" with a retry, for offline or Pages down.
- The floor message when the host is below PowerPointApi 1.2, naming the fix:
  a current Microsoft 365, PowerPoint 2021, or PowerPoint on the web.
- "The insert was refused", "The insert did not confirm: the deck still has N
  slides, nothing was changed", and "The deck grew by one but the copy could
  not be removed: delete slide N by hand". A call can raise and still have done
  the work, so every one of these is written from the measured delta.
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
that nothing leaves the machine; Microsoft's standard EULA; the publisher
StruktureretSundFornuft as a company; first-run guidance; keyboard, focus rings,
live region, high contrast, 320 px; the runtime floor check with a plain message
instead of a `<Requirements>` element; Office.js from the official CDN; the
store logo 300×300, at least one screenshot 1366×768 (the docked 512 px view
serves), descriptions, testing notes and a validators' test deck; one measured
round on web, Windows and Mac, and iPad as section 9 says.

**Not checked by anyone yet: the listing NAME.** Section 12 was read against the
certification policies' section 1120 and the submission form; nothing here has
been read against the naming policy. The sibling SSF Merge is held on exactly
that question — whether policy 1100.7 permits its name — and "SSF Slide
Elements" is the same publisher and the same shape of name, so the answer there
decides the answer here. It is the owner's to settle before a submission, not a
build task.

## 13. Open questions for the host

Each is written so a single round settles it. The probe that asks them is a
Script Lab snippet, not a pane: `docs/PROBE.md` says why, and how each question
is put. The owner runs it on web, Windows and Mac before the splice and the
picker are built, and the sheets are filed under `docs/host-answers/`.

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
4. **The Icons slide** — the least specified of the four. Section 2 says the
   Icons collection "holds the scales and the waste bin, which are neither"
   marker, stamp nor flowchart shape; it says what is wrong and not what the
   fix is. This one needs a decision before it can be done.

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
- **Both decks print correctly, and the prints are committed.** Taken through
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

Measured in the demo and the print: the 16:9 deck has 118 named elements, 21 of
them parts of four collection slides, twelve runs of sizes, 42 whole-slide
elements without a group; the stamps are rotated 29° and 35°; a table's frame is
narrower than the table PowerPoint draws.

**Assumed**: every host fact above on **Mac and iPad**, where no round has been
run. Also assumed, and newly so: that a SEQUENCE of insert-then-remove cycles
behaves the way one does. "Remove from N slides" (section 6) runs one per slide,
and while the single cycle is measured on the web and on Windows, a run of them
is not — it is the one built feature whose mechanism has never been exercised
end to end against a host, and the next round on any platform should put a stamp
on three slides and take it off again. Also assumed: the two-second budget in section 11, and the certification
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
| Publishing on AppSource, publisher StruktureretSundFornuft as a company, web + Windows + Mac at launch | approved |
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
| Right-click opens on `contextmenu` rather than a mouse-only handler, offers nothing on a part, and anchors to the top of its own tile rather than to the pointer — so the keyboard reaches it, the pane never promises a landing the engine does not do, and no coordinate reaches the state | decided in the build, 2026-09-11 |
| The preview card's grey boxes are a snapshot stamped with the slide it was read from, drawn only while the user is still on that slide, rather than a read per slide change — and they are read out of the FILE, so they need no host capability the insert does not already use | decided in the build, 2026-09-11 |
| "Remove from N slides" asks before it removes, runs one confirmed cycle per slide, stops at the first step it cannot verify and says how far it got — and reaches only shapes this add-in tagged | decided in the build, 2026-09-11 |
| A rectangle measured on the user's slide is divided by the USER's slide size, never the library deck's — they differ exactly on a borrowed library, which is the case that would never have shown up in testing | fixed in the build, 2026-09-11 |
| A removal re-reads which slides carry the element from the deck it is about to change, rather than trusting the list the question was asked about | fixed in the build, 2026-09-11 |
| Anything anchored to a tile — the right-click menu, the question before a removal — is keyed by TILE rather than by element, because one element is drawn in up to three lists at once | fixed in the build, 2026-09-11 |
| Base64 is converted by the platform rather than by JSZip, on both sides of the package layer — measured as the single largest cost in the insert path, 4.1 s to 0.23 s on a 45 MB deck | fixed in the build, 2026-09-11 |
