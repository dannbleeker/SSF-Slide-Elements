# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
