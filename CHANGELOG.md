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

- **`SECURITY.md`'s "it makes no network calls" is now "it sends nothing
  anywhere"**, which is both true and stronger. The pane fetches its own
  catalogue from its own origin, as `docs/DESIGN.md` sections 3 and 11 always
  said it would; the test that used to ban `fetch` outright now holds the
  property that matters — one named file may fetch, it may not name an absolute
  address, and it may not pass a method, a body or a header, so there is no way
  to send anything to anyone.

### Fixed

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
  happens to have selected by the time they press it.

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
  size, 117 elements each in twelve categories, twenty-one of them stamps,
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
