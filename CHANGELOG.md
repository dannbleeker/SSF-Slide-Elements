# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added — the splice, the host handshake and the picker

- **The product does what it says on the tin: open the pane, click an element,
  and it lands on the slide you are on.** Three pieces arrived together, because
  none of them is worth anything alone.

- **The splice** (`src/core/splice/`) puts a library element's markup into a
  copy of the destination slide, in the file. It renumbers every shape id
  against the slide it is joining, repoints every relationship reference —
  anchored on the relationships NAMESPACE, so `r:embed`, `r:id` and `r:link` are
  all reached and a spelling nobody thought of is not silently skipped — copies
  every part the element carries into the package under a name the deck is not
  using, walks each of those parts' own relationships and rewrites them too, and
  declares a content type for every one. Then it lands the element where
  `docs/DESIGN.md` section 5 says: a stamp top-right, a marker around whatever
  shape you have selected, a whole-slide element below your own title and scaled
  into the space under it when it would not otherwise fit — including the
  table's grid, because PowerPoint draws a table from its rows and columns and
  ignores the frame around them. Every insert writes a tag naming the element
  and the catalogue it came from, in the package, before the insert, because a
  slide the run just added does not round-trip through the API.

- **A sweep over all 117 elements of the committed 16:9 library**, into a real
  deck, on every commit: every package is put through
  `scripts/package-integrity.mjs` and every shape id checked for collisions.
  `docs/BACKLOG.md` asked for exactly that. Both guards were proved by breaking
  the code and watching them go red.

- **The host handshake** (`src/office/powerpoint.ts`, `src/host/insert.ts`)
  reads the deck with `getFileAsync`, inserts with a `targetSlideId`, and takes
  the replaced slide away **by position**, with the deck counted before, after
  the insert and after the removal. The count is the evidence and the raise is
  not: a call can raise and still have done the work, and a call that raises
  nothing has not necessarily happened. Every sentence the footer shows is
  computed from those three numbers.

- **The picker** (`src/pane/`): search over the English name, the owner's Danish
  name, the category and the tags; a tag line; collapsible categories; a tile
  per element drawing where on the slide it lands; one tile with a stepper for
  an element that comes in several sizes; favourites and recent; a gear holding
  the insert target and whether shapes arrive grouped; a footer with the
  measured slide count, **Again** and **Undo**; the keyboard (`/`, `Esc`,
  arrows, Enter); a live region; Windows high contrast; and the pane reopening
  where you left it. Loading, failure and floor states all say what happened and
  what to do.

### Changed

- **Undo is one deep, not ten.** Taking back an insert that landed onto a slide
  means handing PowerPoint a package containing the slide it replaced, and ten
  of those is ten copies of your presentation inside a task pane. PowerPoint's
  own Ctrl+Z reverts an insert — measured on the web on 2026-09-10 — so the
  deeper history already exists. `docs/DESIGN.md` section 6 records the change
  and the reason.

- **The catalogue index now carries each carried part's content type**, read
  from the library deck rather than guessed from its extension at insert time. A
  part with no content type declared is a package PowerPoint refuses without
  saying which part.

- **The security page said the add-in reads nothing and writes nothing.** It
  reads your whole presentation and writes it back, which is what an insert is
  here, and that page is the one people read when deciding whether to trust the
  tool. It now says so, along with what is kept afterwards and for how long.
  The sentence had survived the commit that shipped the insert — the bullet
  above it, about network calls, was updated in that same commit and this one
  was not.

- **The sibling ledger described a deck read that does not exist.** Four rows
  of `docs/SIBLING.md` and its source table said the deck read pages
  `getItemAt` at twenty and never does a collection load. It does three
  collection loads and no paging. Every row has been re-triaged against the
  code that shipped, and the four now say what is actually there and what
  guards it. A row that reads as a description of this add-in and is not is
  worse than no row, because the next reader builds on it.

- **Numbers in the prose that no longer matched the repo**: the pane draws 84
  tiles, not 85 (the runs cover 45 of the 117 elements); six of the 117 16:9
  elements arrive already tagged, carrying 78 tag relationships between them,
  where the text read as though 78 elements did; the catalogue is about 190 KB,
  not 110; the built library is about 16 MB, which three files said and the
  design record put at 14; the 4:3 library has eleven categories to 16:9's
  twelve. Each was measured against the committed catalogue or the built
  bundle, and two of them against the pane's own rendering.

- **The dependency triage said an alert on `jszip` or `@xmldom/xmldom` was
  not reachable from shipped code.** Both are imported from `src/` now, both
  are in the bundle, and both are what parses a .pptx a user can be sent. The
  page carries a new row saying so, and the old reading is marked superseded
  rather than edited away.

- **Undo's two slide numbers are told apart by name.** One counts from one
  because it is read aloud, the other from zero because it is fed to the API,
  and they were both called `slide` in the same file — which is the confusion
  the undo defect was. A dead `UNDO_DEPTH` of ten went with them: nothing read
  it, and the design has said one deep since the build.

- **The README no longer calls the add-in a scaffold that inserts nothing.**
  The splice, the picker and the host handshake all shipped, and the feature
  table said "planned" for all three; the status line said the pane inserts
  nothing yet. Only the manual is held to the code by CI, so the README drifted
  exactly where the lockstep rule says it would.

- **`SECURITY.md`'s "it makes no network calls" is now "it sends nothing
  anywhere"**, which is both true and stronger. The pane fetches its own
  catalogue from its own origin, as `docs/DESIGN.md` sections 3 and 11 always
  said it would; the test that used to ban `fetch` outright now holds the
  property that matters — one named file may fetch, it may not name an absolute
  address, and it may not pass a method, a body or a header, so there is no way
  to send anything to anyone.

### Fixed

- **An undo could report failure and leave the deck wrong.** PowerPoint on the
  web can still give the OLD slide count after an insert that has already
  happened — 2.8 seconds of it, measured on 2026-09-11 by polling the count
  through a real insert. Undo read the count once, immediately after putting
  the user's slide back, got the old number, concluded the insert had not
  landed and stopped there: the deck kept six slides where five belonged, with
  the restored slide and the rebuilt one both in it. It said "Undo did not
  work" while it was the reading that had failed, not the insert. Every count
  that decides something is now asked again on a backoff until the deck agrees,
  and the ordinary case still costs one read.

- **A deck read that came back short would have named the wrong slide.**
  PowerPoint on the web answers a collection load of more than about fifty
  items with fewer than it has, and this add-in turns that list into a POSITION
  — which slide you are on, which slide an undo aims at — and then removes a
  slide by position. A list missing a slide in the middle makes every position
  after it name a different slide. `docs/SIBLING.md` had triaged this in
  September and promised the defence a sibling uses; the defence was never
  built, and the code shipped doing exactly the read the ledger said it never
  would. It now asks the deck for its slide COUNT in the same breath as the
  list, and hands out no position at all when the two disagree — the pane then
  says it does not know which slide you are on, which is a sentence a user can
  act on. No deck of ours is big enough to have shown this, which is the point.

- **The pane named the slide you were on when it opened, rather than the one
  you are on now.** It read the selection once, at startup, and never again, so
  clicking through the deck left the line under the header naming the first
  slide for the rest of the session. Found in PowerPoint for the web on
  2026-09-11: the API reported slide 2 selected and the pane still read
  "Slide 1.". The insert has always read the selection again for itself, so
  nothing ever landed in the wrong place — but the pane was telling you
  something untrue about where it was about to land.

- **Clicking a tile did nothing.** Most of a tile is the ghost drawing, an
  `<svg>`, and an SVG element is not an `HTMLElement` — which is all the click
  handler would walk up from. Found by opening the add-in in PowerPoint and
  pressing a tile. Every gate passed over it: jsdom was clicking the button,
  which is the part of a tile a user is least likely to hit, and the shot audit
  photographs states without pressing anything.
- **Undo reported success and changed nothing.** Taking back an insert that
  landed onto slide N means putting the user's original slide N back and
  removing the REBUILT one at index N. The code removed N+1, which is the copy
  it had just restored, so the deck came back to its old size, the count check
  passed, and the pane said "Undone" over a slide that had not moved. The
  arithmetic is a pure function now, with the off-by-one as a test, and the
  restoring insert aims at the rebuilt slide rather than at whatever the user
  happens to have selected by the time they press it. Confirmed in PowerPoint
  for the web on 2026-09-11 by reading the slide's shapes before and after: a
  triangle inserted onto a slide holding a title and a white box, then taken
  back, left the slide with the same id and the same two shapes it started
  with.

- **A new slide carried the previous slide's comments.** Found by running the
  real engine against PowerPoint for the web on 2026-09-10 and then reading the
  deck back: one comment came out on two slides. A modern comment is anchored
  from the slide's own extension list, so cloning a slide keeps it — right when
  the slide is being rebuilt, wrong when it is meant to be new. "As a new
  slide" now drops comments the way it already dropped speaker notes, and the
  same round afterwards showed the new slide with none.
- The harvest's `parts` list was documented as "every package part reachable
  from those relationships", and it is not: parts are collected once per deck,
  so the second element to use a picture lists nothing for it. Measured on the
  committed library, 64 relationship targets across the two sizes are absent
  from their own element's list. The splice resolves what to copy from the
  element's relationships instead, and the type says so.


### Added — the host probe

- A way to ask a real PowerPoint the questions the design rests on, before
  anything is built on a guess: paste `probe/probe-snippet.ts` into Script Lab,
  press Run, and copy the answer sheet back. It asks whether a package pruned
  to one slide is accepted, whether inserting a slide and then removing the one
  it replaced keeps the order, which slide the host says you are on, which of
  the two ways of reading the deck drops your comments, whether PowerPoint's
  own Ctrl+Z takes an insert back, and how long a read takes on a big deck.
  `docs/PROBE.md` has the steps; `scripts/read-answers.mjs` says what each
  answer means and files the sheet. Nothing in the pane changes.
- The first round, on PowerPoint for the web (2026-09-10): two pairs of
  sheets under `docs/host-answers/`, the second pair taken with the marker
  below. Every insert landed, both prunings land as one slide, the export
  drops comments and the authors part, and PowerPoint's own Ctrl+Z reverts an
  insert. `docs/PROBE.md` carries the summary.

### Fixed — the host probe

- The second run of a pair left a slide of its own. The only thing that told a
  second run from a first was the slide the first run leaves for Ctrl+Z, which
  is exactly what a successful Ctrl+Z removes, so on the web the second run
  took itself for a first. The first run now writes a marker into the
  document's settings, outside the undo stack, before it leaves the slide; the
  second run reads it, clears it, and leaves nothing. A lone second sheet also
  answers question 5 now, from the marker, and the reader says which source
  it read.
- The fixture-timestamp test compared local time components against an
  instant JSZip reads back in UTC, so the suite was red on any machine east
  of Greenwich and green on CI.

### Added — the library harvest

- The library exists as data: the two decks the owner authored (one per slide
  size, 117 elements each, in twelve categories for 16:9 and eleven for 4:3
  (the 4:3 deck has no white boxes with black headings), twenty-one of them
  stamps,
  markers, flowchart shapes and icons) are read into a catalogue with every
  element's English name, where it sits, where it lands and which pictures,
  charts and tags it carries. Nothing in the pane shows it yet; the picker is
  the next change but one. Editing the library is editing the deck: the rules
  are in the manual under "Adding an element to the library", and a change
  that breaks them is refused with every problem listed.

### Added — the design record

- `docs/DESIGN.md` says what the pane and the insert will do, decision by
  decision and dated: what counts as an element, how the library is authored,
  where each kind of element lands, what the footer reports, and what is still
  a question for a real PowerPoint. `template/names.en.json` carries the
  English name of every element in the library. Nothing in the pane changes
  yet.

### Added — the package layer

- The engine can now open a .pptx, read and change its parts, relationships,
  content types and slide list, and hand it back as bytes or base64. Nothing
  in the pane uses it yet; it is the ground the element library and the insert
  are built on, ported from SSF-Merge together with the tests that found its
  bugs there.

### Added — the sibling watch

- A weekly sweep of what SSF-Charts and SSF-Merge have learned about the
  PowerPoint host, filing one issue for any finding this repo has not answered.
  The answers so far live in `docs/SIBLING.md`: which of the siblings' findings
  shape the design, and which do not touch it at all.

### Added — the scaffold

The repository, before any of the product: everything that has to be true for
the first feature to land safely.

- Hosting on GitHub Pages at `ssf-slide-elements.struktureretsundfornuft.dk` (the
  domain's DNS record is the owner's step), with the deploy waiting for the same
  gate CI runs.
- CI: format, lint, types, the library build, the pane build, coverage with
  floors, and a floor under the number of tests that CI refuses to leave
  unrecorded.
- Manifests in both formats, dev and prod, generated from one source, checked
  offline by the suite and by Microsoft's validator in CI. The GUID is minted
  and pinned; the manifest version is `1.0.0.0` and moves only when the manifest
  does.
- Ribbon icons drawn by a script and byte-pinned by a test.
- A placeholder pane: one step, **Start here**, showing which build it is,
  checking that the host clears PowerPointApi 1.2, and saying plainly that there
  is nothing to insert yet. The **Insert an element** button is there and
  disabled.
- The documentation set — manual, backlog, this changelog, the security policy,
  the dependency-alert log — and the tests that keep the manual in step with the
  pane.
- A manual release workflow that validates what it ships and refuses to ship a
  manifest pointing at a host that is down; a weekly pane audit; Dependabot with
  a written log; a contributing guide, a PR template, code owners and the MIT
  licence.
