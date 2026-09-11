# Backlog

The single curated list of what is open. Items graduate from here into a PR and
are **removed when they ship** — the README's feature table and the manual are
where shipped work is described, so anything still listed here is genuinely not
done.

Priority is what it costs the product to be without it, not how interesting it
is to build. The order below is the order the work arrives in; each entry is
one pull request, ported from SSF-Merge's code where the same code exists there.

## Open

### The rest of section 4

Built: search, tags, categories, tiles with the element's own picture on them,
the stepper, favourites, recent, the gear, the footer with the measured delta,
Undo, Again, the keyboard, the live region, high contrast, the pane reopening
where you left it, the preview card, "Did you mean …" for a query with no hits,
the category chips a search shows, the stepper greying the sizes a search did
not ask for, the first-run coach marks, the gear's two external links, and
**Used in this deck**. Still to come, all of it from `docs/DESIGN.md` sections
4, 6 and 8: **the jump** from a slide number in "Used in this deck", which needs
a host call no round has made. **Remove from N slides** is built, and is the one
feature whose mechanism — a sequence of insert-then-remove cycles — no round has
exercised; the next round on any platform should put a stamp on three slides and
take it off again.

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

Written and built: the 300×300 store logo (`npm run icons`), the descriptions
and the validators' testing notes (`docs/LISTING.md`, held to the manifests by
`test/listing.test.ts`), the catalogue page on the site
(`public/catalogue.html`, generated at harvest), and the privacy page naming
the browser storage.

Left, and all of it needs either a screen or the owner: **a 1366×768
screenshot** of the pane beside a real presentation; **a validators' test
deck**, authored in PowerPoint rather than built by this repo's own code;
**the listing name**, which waits on the same naming-policy answer SSF Merge
waits on; then the Partner Center submission of `manifest-prod.xml`, and
v0.1.0 on the releases page.

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
