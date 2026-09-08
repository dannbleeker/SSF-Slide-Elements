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

The owner's library deck under `template/`, read into a committed catalogue:
one JSON per element carrying its section, name, slide markup and media. The
deck is the authoring surface — the layout says what a slide is for, the title
placeholder is the name, everything that is not a placeholder is the element —
and a CI job fails a change where the deck and the catalogue disagree. Decide
then whether a catalogue of that size belongs in git or is built on deploy.

### Splice, and the package self-check

An element's markup into the slide the user is on, in the file: part names,
nested relationship targets, relationship ids and shape ids all rewritten, each
with a failure that opens as "PowerPoint found a problem with this file". A
sweep over every element in the library, on every commit.

### The picker

`steps.ts` grows from one step to browse → choose → insert; `render.ts` draws
the library by section with search and thumbnails; `pane-shots` states grow
with it; the manual's pane section stops being planned.

### Host handshake, and the first real-host round

`src/office/powerpoint.ts` reads the deck, inserts with a `targetSlideId`, and
removes the replaced slide positionally, each step proven by the deck delta;
`test/fixtures/host.ts` fakes the host far enough to drive it; then one round
against PowerPoint on the web, recorded, before the first release.

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
