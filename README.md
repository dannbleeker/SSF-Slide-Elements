# SSF Slide Elements

A library of ready-made slide elements for PowerPoint, as an Office.js task pane
add-in. Browse the collection, click one, and it lands on the slide you are on —
with every font, colour, table and chart exactly as it was drawn.

Part of the SSF add-in family from
[StruktureretSundFornuft](https://struktureretsundfornuft.dk), alongside
[SSF Merge](https://github.com/dannbleeker/SSF-Merge) (mail merge for
PowerPoint) and [SSF Charts](https://github.com/dannbleeker/SSF-Charts) (native,
editable charts).

## Status

A scaffold, wired for hosting and installable: sideload `manifest-prod.xml`, press
**Slide elements** on the Home tab, and the pane opens, shows which build it is
and checks that your PowerPoint clears the floor. **It inserts nothing yet.**
The library, the picker and the insert each arrive as their own change, in the
order [the backlog](docs/BACKLOG.md) gives.

| Piece | State |
| --- | --- |
| Hosting on GitHub Pages at `ssf-slide-elements.struktureretsundfornuft.dk`, deployed only after the same gate CI runs | done — live once the custom domain's DNS record exists |
| CI: format, lint, types, both builds, coverage floors, a floor under the number of tests | done |
| Manifests — XML and unified JSON, dev and prod, generated from one source and validated by Microsoft's tool in CI | done |
| Ribbon icons, drawn by a script rather than by hand — the PNGs are committed, and a test redraws them and fails on a byte of difference | done |
| Guard tests: the layering, the manifests, the release, the docs lockstep, the pane's boot | done |
| Release workflow — manual, validated, assets checked against the docs | done |
| Weekly pane audit: overflow, contrast, focus, hit areas, axe, at 320 and 512 in both themes | done |
| Weekly sibling watch over SSF-Charts and SSF-Merge — one issue for any finding with no row in this repo's ledger | done |
| Package layer — the .pptx as parts, relationships, content types and the slide list, ported from SSF-Merge with its tests | done |
| Library harvest — two library decks under `template/` read into a catalogue: categories, 117 elements per size with their English names, boxes, landing and sizes, and the markup and media each one carries. The index is committed and CI-checked against the decks; the markup and media are built on deploy | done |
| Splice — an element into the slide you are on, and the package self-check | planned |
| The picker — browse by section, search, insert | planned |
| Host handshake, and the first round against a real PowerPoint | planned |

## How it will work

A .pptx is a zip of XML parts, and the insert will happen **in the file**.

Office.js has no call that puts arbitrary markup onto a slide: the shape
collection can only add geometry it has a method for, so an element with a
gradient, a table or an icon cannot be reproduced through the API without losing
something. SSF Charts measured what drawing shape by shape costs on PowerPoint
for the web — re-authored text, ids the host will not resolve, runs that take
minutes — and SSF Merge is the answer this family arrived at: read the open
presentation, change the package, hand it back through one
`insertSlidesFromBase64`. This add-in takes the same route.

```
library deck ──harvest──► catalogue (committed) ─┐
                                                ├─► splice (pure TypeScript) ─► .pptx bytes ─► one host call
the open presentation ──────────────────────────┘
```

`src/core` imports nothing from Office.js and `src/office` decides nothing of
its own; `test/architecture.test.ts` holds both directions.

## Documentation

| Document | What it is for |
| --- | --- |
| [docs/MANUAL.md](docs/MANUAL.md) | How to install it, what the pane does today, and what is planned |
| [docs/DESIGN.md](docs/DESIGN.md) | The design record: what an element is, the pane top to bottom, where things land, the host questions still open, and every decision with its date |
| [docs/BACKLOG.md](docs/BACKLOG.md) | What is open, in order, and what has been rejected |
| [CHANGELOG.md](CHANGELOG.md) | What changed, newest first |
| [docs/DEPENDENCY-ALERTS.md](docs/DEPENDENCY-ALERTS.md) | Every Dependabot alert, and what was decided about it |
| [docs/SIBLING.md](docs/SIBLING.md) | What SSF-Charts and SSF-Merge learned about the host, what was done about each finding here, and the rule that keeps a borrowed number dated |
| [CLAUDE.md](CLAUDE.md) | Project memory: architecture, host rules, conventions |

These are kept in step with the code by `test/docs.test.ts`, which reads the
pane's steps and labels out of the source and fails when the manual has not
caught up. A feature and its documentation land in the same change or neither
does.

## Commands

```bash
npm install

npm test           # the suite
npm run typecheck  # types
npm run lint       # type-aware ESLint
npm run format     # Prettier, on code only
npm run coverage   # the suite with coverage floors
npm run test:count # a floor under the number of tests

npm run build:lib  # compile the engine to dist-lib/
npm run build      # the pane and the site, for GitHub Pages
npm run dev        # the pane at localhost:3002
npm run icons      # redraw public/assets/*.png
npm run manifests  # regenerate the four manifests from scripts/manifest-source.mjs
npm run pane-shots # render and measure the pane; needs `npx vite --port 5199 --strictPort &` first
```

Every npm script is FLAT. A script that nests `npm run` is blocked by AppLocker
on the owner's Windows box, which is where this gets run.
[CONTRIBUTING.md](CONTRIBUTING.md) explains the rules the checks hold.

## Installing it

The add-in is served from `ssf-slide-elements.struktureretsundfornuft.dk`.
Sideload `manifest-prod.xml` — from the
[latest release](https://github.com/dannbleeker/SSF-Slide-Elements/releases/latest)
once there is one, or
[the file on `main`](https://github.com/dannbleeker/SSF-Slide-Elements/raw/main/manifest-prod.xml),
which is the same pointer. `manifest.xml` points at `localhost:3002` for
development. [The manual](docs/MANUAL.md#installing-it) has the steps for each
PowerPoint.

The floor is PowerPointApi 1.2, checked when the pane opens rather than declared
in the manifest — a declared requirement the host lacks makes an add-in vanish
from the ribbon with no diagnostic at all.

## Licence

MIT.
