# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build. The order below is the order the work arrives in; each entry is
one pull request, ported from SSF-Merge's code where the same code exists there.

## Open

### Library harvest

The owner's two library decks under `template/` (one per slide size, both
carrying the same names) read into a committed catalogue: one entry per element
with its category, English name from `template/names.en.json`, slide markup,
media, box, size run and landing. A collection slide yields one element per
shape and its title is their category. Previews are cut from the PDF prints the
owner commits beside the decks, and a CI job fails a change where the deck, the
print, the names file and the catalogue disagree. What an element is, how it is
named and how it is cut is `docs/DESIGN.md` sections 2 and 3.

### Host probe

Before anything is spliced or drawn: a thin probe pane that asks the six
questions in `docs/DESIGN.md` section 13 of a real PowerPoint (a pruned package
on the way in, insert-then-delete order, `getSelectedSlides`, which read of the
deck, Ctrl+Z after an insert, the time to read a large deck), one round each on
the web, Windows and Mac, recorded in `docs/host-answers/`. The splice and the
picker are built on the answers, not on the siblings' rounds.

### Splice, and the package self-check

An element's markup into the slide the user is on, in the file: part names,
nested relationship targets, relationship ids and shape ids all rewritten, each
with a failure that opens as "PowerPoint found a problem with this file"; the
landing rules; the tag written at insert; theme colours pinned or kept per the
colour switch. A sweep over every element in the library into fixture decks
(empty, dark theme, tall title, 4:3, A4), every package validated by the
integrity checker, on every commit.

### The picker

`steps.ts` grows from one step to browse → choose → insert; `render.ts` draws
the library as `docs/DESIGN.md` sections 4 to 10 describe it, all of it in v1:
search, tags, categories, sizes, the preview card, favourites, recent, used in
this deck, the footer with the measured delta, Undo, Again, the gear, the
first-run guide, keyboard, touch, high contrast, the loading and failure
states. `pane-shots` states grow with it; the manual's pane section stops being
planned.

### Host handshake, and the first real-host round

`src/office/powerpoint.ts` reads the deck, inserts with a `targetSlideId`, and
removes the replaced slide positionally, each step proven by the deck delta;
`test/fixtures/host.ts` fakes the host far enough to drive it; then one round
against PowerPoint on the web, Windows and Mac, recorded, and iPad as
`docs/DESIGN.md` section 9 says.

### Release and AppSource

The 300×300 store icon, a 1366×768 screenshot, the descriptions, the
validators' testing notes and test deck, the catalogue page on the site, the
privacy page naming the browser storage, then the Partner Center submission of
`manifest-prod.xml`, and v0.1.0 on the releases page.

## Rejected — do not re-propose

- **Inserting shape by shape through Office.js.** Rejected before trying, on
  SSF-Charts' rounds: setting text through the API re-authors it
  (office-js#5858), `addTextBox` deletes the selected shape on the web
  (#2775), `setSelectedShapes` wedges the host (#3083, #3698), a shape proxy
  does not survive a `context.sync()`, and a run of eight charts took minutes.
  An element with a gradient, a table or an icon cannot be reproduced through
  the shape collection without losing something. The package route has none of
  those failure surfaces, and it is what the family has measured.
- **Mining the discarded first implementation** (commit `fbc0870`, PR #1,
  closed unmerged 2026-09-08). Rejected by the owner: it shipped a manifest
  version Office rejects, named scripts that did not exist, sat behind both
  siblings on every toolchain major, pointed its CNAME at a different subdomain,
  and measured nothing against a host. The increments above are rebuilt from
  the siblings, not from it.
- **Twelve pane ideas the owner turned down** on 2026-09-08, listed in the
  decisions log of `docs/DESIGN.md`: among them an authored tag vocabulary, a
  per-element description, a "New" chip, swapping the element already on the
  slide, Enter inserting the top hit, several new slides in one pick, authored
  sets, version-aware elements, a density toggle and a go-to-category dropdown.
