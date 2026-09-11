# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build. The order below is the order the work arrives in; each entry is
one pull request, ported from SSF-Merge's code where the same code exists there.

## Open

### The library's own pictures

Every tile draws a small diagram of where the element lands, because that is
what the catalogue knows. What it does not have is a PICTURE of the element:
`docs/DESIGN.md` section 3 has previews cut from a PDF print of each library
deck, and neither print is committed. Until the owner prints both decks to PDF
and commits them beside the `.pptx`, the harvest has nothing to cut, and the
pane is a list of names and landings rather than a picture book. The tile has
the slot; nothing else is blocked by it.

**The print is blocked on the decks themselves.** Measured on Windows on
2026-09-11 (section 3): neither committed deck opens in desktop PowerPoint —
two `[Content_Types].xml` overrides per deck are missing the leading `/` on
their `PartName`, and PowerPoint offers Repair instead of opening. So this item
now has a step in front of it: correct those four part names, which changes no
slide, note, picture or embedded object and leaves the harvested catalogue index
byte-identical (measured, same version `1641fe687794`), and only then print. The
Repair is not the way out — on a copy it dropped ten of the deck's fifteen
embedded OLE objects.

### The preview card, and the rest of section 4

Built: search, tags, categories, tiles, the stepper, favourites, recent, the
gear, the footer with the measured delta, Undo, Again, the keyboard, the live
region, high contrast, and the pane reopening where you left it. Still to come,
all of it from `docs/DESIGN.md` sections 4, 6 and 8: the preview card that
opens after a third of a second of hover and docks beside the list at 512 px;
**Used in this deck**, read from the tags the insert writes (measured on 2026-09-11: they survive the insert, see `docs/DESIGN.md` section 15); **Remove from N
slides** for a stamp; the three first-run coach marks; right-click for the other
insert target on one insert; category chips with counts while searching, and
"Did you mean …" for a query with no hits.

### The colour switch

`docs/DESIGN.md` section 7 offers "This deck's theme" against "As in the
library". The first is what the splice does today and costs nothing — a
theme-referenced colour follows the destination deck by itself. The second means
pinning every `<a:schemeClr>` to the value the library's theme gives it, which
needs the harvest to carry that theme's colour map. The gear does not show the
option yet, because a switch with one working position is worse than no switch.

### The rest of the host round

PowerPoint for the web is measured, twice: `docs/host-answers/` carries two
pairs of sheets from 2026-09-10 and `docs/PROBE.md` reads them. **Windows is
measured too** — one pair of sheets from 2026-09-11 and a product round of four
inserts and undos from the pane the same day, both read in `docs/DESIGN.md`
section 15. Mac and iPad are not, and `docs/DESIGN.md` section 9 says how the
first iPad measurement will be taken. So what is left is a round on each of
those two, which is what turns the last of the design's "assumed" column into
"measured".

### Release and AppSource

The 300×300 store icon, a 1366×768 screenshot, the descriptions, the validators'
testing notes and test deck, the catalogue page on the site, the privacy page
naming the browser storage, then the Partner Center submission of
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
