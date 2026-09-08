# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
