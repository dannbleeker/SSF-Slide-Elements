# Security

## Reporting

Report a vulnerability privately through GitHub's
[security advisories](https://github.com/dannbleeker/SSF-Slide-Elements/security/advisories/new)
rather than in a public issue.

## What this add-in touches

Worth stating plainly, because a presentation is often confidential.

- **Today the pane reads nothing and writes nothing.** It shows which build it
  is and asks PowerPoint which API version it has. That is the whole of it.
- **By design, everything happens on the user's own machine.** The element
  library will ship inside the pane as static files built from the owner's own
  deck; the user's presentation will be read and written through Microsoft's
  Office JavaScript interface, in the pane, and never uploaded.
- **It makes no network calls.** There is no `fetch`, no `XMLHttpRequest` and
  no socket anywhere in `src/`. The pane and its assets are served from
  `ssf-slide-elements.struktureretsundfornuft.dk` and nothing else is requested
  after that, other than Microsoft's own `office.js`, which every Office add-in
  is required to load.
- **The add-in has no backend.** There is no server to breach and no log to
  leak.
- **The manifest asks for nothing beyond its own host.** No `AppDomains`, so the
  pane cannot navigate off it, and no `WebApplicationInfo`, so it requests no
  Microsoft identity or Graph scope.

Anything that changes that stands as a change to this section as much as to the
code. A security page is read by people deciding whether to trust the tool, and
it is the last page anyone thinks to update.

## Handling untrusted input

Nothing yet: the pane takes no input. When the package layer lands, a .pptx
arrives from outside and a user can be sent one, and the rules a sibling
project learned apply here — a value reaches XML as text and only as text, a
column or element name is data and never a property, and a relationship target
is never trusted to name a part for deletion. Each will be listed here when it
is executable, and executed in `test/security.test.ts`.

The claims on the front of this page already are: that file reads the source
and fails on a network call or a markup sink.

## All sample data is invented

The repository is public. Nothing in it is a real presentation, a real
customer, or a real person.
