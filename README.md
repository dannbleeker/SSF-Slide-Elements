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

Built, and run end to end against a real PowerPoint. Sideload
`manifest-prod.xml`, press **Slide elements** on the Home tab, pick an element,
and it lands on the slide you are on. The first whole round on PowerPoint for
the web was **2026-09-11**: a tile clicked, the element on the selected slide,
and Undo putting that slide back to the shapes it held before. **Windows was
measured the same day** — a probe pair and a product round, both read in
`docs/DESIGN.md` section 15. Mac and iPad have had no round yet, and
[the backlog](docs/BACKLOG.md) is what is still open.

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
| Host probe — a Script Lab snippet that asks a real PowerPoint the seven questions the design rests on (a pruned package on the way in, insert-then-delete order, the selected slide, which read of the deck, Ctrl+Z, timing and the floor), and a reader that says what each answer means and files the sheet | done — four sheets from PowerPoint for the web under `docs/host-answers/`, 2026-09-10 |
| Splice — an element into the slide you are on, and the package self-check | done — proven over all 117 elements against the package integrity check |
| The picker — browse by section, search, insert, Recent, favourites, Undo one deep | done |
| Used in this deck — which library elements the open deck already holds, and where, read from the tags an insert writes | done |
| The jump — a slide number in Used in this deck goes to that slide, and the pane says so only once PowerPoint reports being there | done — on SSF-Charts' web measurement of the call, read back on every click; not yet measured on any host by this repo (probe question 7) |
| Right-click a tile for the other insert target, for one insert | done |
| Move to a new slide — a whole-slide element that landed on a slide with something already on it, put on a slide of its own instead | done — an undo and a second insert, so it can offer nothing the undo cannot deliver |
| The preview card's grey boxes — what the destination slide already holds, read from the deck | done |
| Remove from N slides — a stamp off every slide it is on, asked first, one confirmed cycle per slide | done — built, and the one feature no round has exercised |
| The colour switch — an element takes the destination deck's theme, or keeps the library's own colours, shapes and carried charts alike | done |
| Host handshake, and the first round against a real PowerPoint | done — PowerPoint for the web and PowerPoint on Windows, both 2026-09-11; Mac and iPad still unmeasured |
| Element pictures — each tile shows PowerPoint's own rendering, cut out of a committed PDF print of the library deck; the landing diagram stays as the fallback | done |
| The catalogue page on the site — every element in both sizes, generated at harvest and diffed by CI | done |
| The store listing — logo, descriptions, validators' notes and a PowerPoint-authored test deck, each held by a test | done — the screenshot and the listing NAME are the owner's, and `docs/LISTING.md` says why |

## How it works

A .pptx is a zip of XML parts, and the insert happens **in the file**.

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
| [docs/PROBE.md](docs/PROBE.md) | The host probe: how to run it in Script Lab on the web, Windows and Mac, what each question decides, and where the answer sheets are filed |
| [docs/BACKLOG.md](docs/BACKLOG.md) | What is open, in order, and what has been rejected |
| [docs/LISTING.md](docs/LISTING.md) | The AppSource listing: the fields, the descriptions, the validators' notes, and what is still the owner's |
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
npm run probe      # regenerate probe/probe-snippet.ts for Script Lab; needs build:lib first
npm run pane-shots # render and measure the pane; needs `npx vite --port 5199 --strictPort &` first
npm run dead-exports # list exports in src/ that nothing shipping calls; the suite holds it at zero
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
