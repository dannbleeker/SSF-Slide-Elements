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
  not a generic `Gruppe N`, else by the slide title numbered. The deck stays
  Danish and its titles are the **keys**; `template/names.en.json` carries the
  English name for every key, one file per locale later, and the harvest fails
  on a key with no name. The pane shows only the English name: no Danish
  tooltip, no per-element description.
- **Placeholder text in the library is English in v1**, in both decks, so what
  lands on the slide reads in the pane's language. Per-locale text comes later
  the way names do. Slide titles stay Danish because they are the keys and are
  never shown.
- **An element that comes in several sizes is one tile with a stepper.** A run
  of elements that differ only by one count (process flows with 1 to 6 boxes,
  hierarchies with 2 to 5 boxes, matrices with 2 to 5 rows, and so on) is
  derived from the name pattern: a number that stands alone before a word,
  never inside "2×2" or "1-2-3", "box" the one irregular plural, and the
  key-figure flows one run by an explicit rule. The stepper names what it
  counts ("boxes 1 2 3 4 5 6"). The pane never uses the word "family"; that
  word is for this document. The 16:9 deck has twelve such runs, so 117
  elements show as 85 tiles.
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
  harvest cuts every element's preview from it and fails when the print's slide
  count does not match the deck's. Every deck change therefore needs a re-print.
  A part is cut to its box with 3% of air, the boxes of neighbouring parts
  painted white, and a rotated part masked to its rotated frame. The parts are
  the same objects in both decks, so their cuts are shared.
- **Boxes.** The box of a rotated shape is its rotated extent, not the
  unrotated frame the XML gives (the owner's stamps are rotated 29° and 35°). A
  table's box is the sum of its columns and rows, not its frame's `ext`, which
  PowerPoint ignores when it draws the table. Landing and cropping both depend
  on this.
- **Colours.** The library is authored in theme colours (accent 2 alone is used
  139 times). A theme-mapped colour follows the destination deck's theme by
  itself when the markup keeps it as a theme reference; the colour switch
  (section 7) pins them to the library's values instead.
- **The catalogue is static.** `npm run harvest` reads the decks into the
  catalogue the pane ships as files on Pages: an index (`catalogue.json`,
  about 160 KB: names, categories, boxes, sizes and landing per element for
  both sizes, and a content hash as its version) plus one file per element with
  its markup, relationships and the parts it carries, and the parts themselves.
  The index is committed and CI fails when it no longer matches the decks; the
  element files and parts (about 14 MB, generated JSON) are built on every
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
- **Favourites** (a star in the tile's corner, remembered per machine) and
  **Recent** (the last six inserts) sit above the categories.
- **Categories** start collapsed, with "Open all" beside the count; a search or
  a tag opens what it finds; category headers stick while scrolling. Tiles are
  two across, three from 400 px. A tile is the element's picture and name; a
  sized element carries its stepper; a part already in the deck carries
  "Remove from N slides".
- **The preview card** opens after a third of a second of hover or focus,
  pinned over the top of the list: the element at full width, its name, one
  line saying where it lands, and a small slide with a ghost frame for the
  landing and grey boxes for what the slide already has. At 512 px and wider
  the card docks beside the list and hides no tiles.
- **Footer**: the last outcome with the measured slide count, then the actions
  (**Move to a new slide** when a whole-slide element landed on a slide that
  already had content, **Again**, **Undo (n)**), then a line with the current
  settings that opens the gear.
- **First open ever**: three coach marks (hover to preview, click to insert,
  Undo and the gear), dismissed once.
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
- **Undo** goes ten deep, in order, and answers Ctrl+Z while the pane has
  focus. It is positional and count-checked, never by id. **Again** repeats the
  last insert on the current slide or selection.
- **A whole-slide element onto a slide with content** gets the footer offer
  "Move to a new slide" beside Undo. **Empty content placeholders** on the
  slide are removed when a whole-slide element lands, so no "Click to add
  text" ghost sits behind it; the title placeholder stays; Undo puts the
  placeholder back.
- **Multi-shape elements land as one group** (gear option, default on), so the
  user moves them as one and ungroups when editing.
- **Right-click** (long-press on touch) on a tile offers the other insert
  target for that one insert, without touching the setting.
- **Deck-wide stamps.** A stamp already in the deck can be removed from every
  slide it is on with one click, found by the tag written at insert. The manual
  says that shape tags do not survive cut and paste on the web.
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
- **Report a problem**: opens the support page in the browser with the build
  stamp, host and platform prefilled, so a report is usable.
- **Browse the catalogue on the site**: opens the catalogue page.
- External links open a browser window from the pane
  (`Office.context.ui.openBrowserWindow`, probed at runtime; a plain new-tab
  link is unreliable on desktop) and never navigate the pane itself.

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

## 12. What AppSource certification needs

Checked against Microsoft's commercial-marketplace certification policies for
Office add-ins (section 1120) and the Partner Center submission form, assuming
the XML manifest is submitted and the listing is free.

Covered by the design or the release plan: a stable GUID and a version bump per
submission; an HTTPS-only origin, no localhost, no `AppDomains` because nothing
navigates the pane; `ReadWriteDocument` only; no sign-in, no network calls, no
data collection; the support and privacy pages on the site, the privacy page
naming the browser storage (favourites, pane state, the first-run flag) and
that nothing leaves the machine; Microsoft's standard EULA; the publisher
StruktureretSundFornuft as a company; first-run guidance; keyboard, focus rings,
live region, high contrast, 320 px; the runtime floor check with a plain message
instead of a `<Requirements>` element; Office.js from the official CDN; the
store logo 300×300, at least one screenshot 1366×768 (the docked 512 px view
serves), descriptions, testing notes and a validators' test deck; one measured
round on web, Windows and Mac, and iPad as section 9 says.

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
release and AppSource. Before the harvest the owner edits both decks: English
placeholder text, the Icons slide, the rules text into notes, the collection
marker line, then a fresh PDF print of each.

## 15. Measured and assumed

**Measured**: nothing against a real PowerPoint by this repo. Every host fact
here is borrowed from SSF-Charts and SSF-Merge and dated in `docs/SIBLING.md`.
Measured in the demo and the print: the 16:9 deck has 118 named elements, 21 of
them parts of four collection slides, twelve runs of sizes, 42 whole-slide
elements without a group; the stamps are rotated 29° and 35°; a table's frame is
narrower than the table PowerPoint draws.

**Assumed**: everything in sections 5, 6 and 10 that depends on the host, the
two-second budget in section 11, and the certification reading in section 12.
The six questions in section 13 are the first things to measure.

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
| The probe tells a second run from a first by a marker in the document settings, written before the slide it leaves so the user's Ctrl+Z still lands on the insert; the slide alone could not, because Ctrl+Z is what removes it (web round, 2026-09-10) | decided in the build, 2026-09-10 |
