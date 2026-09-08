# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
