# Security

## Reporting

Report a vulnerability privately through GitHub's
[security advisories](https://github.com/dannbleeker/SSF-Slide-Elements/security/advisories/new)
rather than in a public issue.

## What this add-in touches

Worth stating plainly, because a presentation is often confidential.

- **The pane reads your whole presentation and writes it back**, because that
  is what an insert is here. Office.js has no call that puts arbitrary markup
  onto a slide, so the add-in reads the open deck through `getFileAsync`,
  changes the package, hands it back through one `insertSlidesFromBase64`, and
  removes the slide it replaced. It reads the deck once per insert and keeps
  one copy afterwards, so Undo has something to put back; that copy is dropped
  when you insert again or close the pane.
- **All of that happens on your own machine.** The element library is static
  files on the add-in's own site, built from the owner's own deck; your
  presentation is read and written through Microsoft's Office JavaScript
  interface, in the pane, and is never uploaded.
- **It sends nothing anywhere.** There is no `XMLHttpRequest`, no socket and no
  `sendBeacon` anywhere in `src/`, so the add-in has no way to transmit
  anything at all. It does make one kind of request: the element library is
  static files on this same site, and the pane fetches the catalogue index when
  it opens and an element's markup and pictures when you insert one. Those are
  plain GETs for the add-in's own files, from one file — `src/pane/catalogue.ts`
  — which may not name an absolute address and may not pass a method, a body or
  a header. Nothing from your presentation is ever put into a request.
  Everything else the pane loads is served from
  `ssf-slide-elements.struktureretsundfornuft.dk`, other than Microsoft's own
  `office.js`, which every Office add-in is required to load.
- **The add-in has no backend.** There is no server to breach and no log to
  leak.
- **The manifest asks for nothing beyond its own host.** No `AppDomains`, so the
  pane cannot navigate off it, and no `WebApplicationInfo`, so it requests no
  Microsoft identity or Graph scope.

Anything that changes that stands as a change to this section as much as to the
code. A security page is read by people deciding whether to trust the tool, and
it is the last page anyone thinks to update.

## Handling untrusted input

A .pptx arrives from outside, and a user can be sent one. The package layer
(`src/core/pptx/`) is ported from SSF-Merge together with the rules its two
security sweeps produced, and each rule is executed in `test/security.test.ts`:

- **A relationship target is never trusted to name a part for deletion.**
  Removing a slide collects that slide's notes page and comments, and a target
  resolving outside `ppt/` is left alone. The parts a slide's chart or diagram
  claims to own are held to an allowlist, one level down as well, so a crafted
  relationship naming `/ppt/presentation.xml` or `/[Content_Types].xml` cannot
  have the sweep delete it.
- **Traversal out of the package is not possible.** `..` past the root is
  clamped and the result is always a package-relative name; nothing here
  touches a real filesystem.
- **The XML parser does not fetch or expand what a deck tells it to.**
  `@xmldom/xmldom` resolves no external entity and expands no internal one,
  measured rather than assumed, so a version bump that changes its mind is a
  red test.
- **Text reaches XML as text and only as text.** Nothing here builds markup by
  string concatenation, and the one rule for characters XML cannot carry at all
  (`xmlSafe`) replaces them with a space. The splice writes into a deck through
  a parsed document; the one thing it builds by concatenation is a tag part,
  whose values run through `xmlSafe` and then through a single escape.

The claims on the front of this page are executable too: that file reads the
source and fails on a network call or a markup sink.

## All sample data is invented

The repository is public. Nothing in it is a real presentation, a real
customer, or a real person.
