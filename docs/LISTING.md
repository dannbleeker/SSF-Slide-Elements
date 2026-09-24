# The AppSource listing

Everything the Partner Center submission form asks for, written down here so
the submission is a copy-and-paste rather than a writing session, and so what
goes in the store can be reviewed in a pull request like everything else.

`docs/DESIGN.md` section 12 is the checklist this serves; `test/listing.test.ts`
holds the parts of it that can go stale — the name, the two descriptions and the
three URLs are asserted equal to what the manifests actually carry, and every
file named below is asserted to exist.

**Nothing here has been submitted.** The submission itself is the owner's, and
so is the decision under "Still the owner's" at the bottom.

## The fields

| Field | Value |
| --- | --- |
| Name | SSF Slide Elements |
| Provider | StruktureretSundFornuft ApS |
| Short description | Ready-made slide elements, dropped onto the slide you are on. |
| Support URL | <https://ssf-slide-elements.struktureretsundfornuft.dk/support.html> |
| Privacy URL | <https://ssf-slide-elements.struktureretsundfornuft.dk/privacy.html> |
| Terms | <https://ssf-slide-elements.struktureretsundfornuft.dk/terms.html> |
| Licence | <https://ssf-slide-elements.struktureretsundfornuft.dk/license.html> |
| Store logo | `public/assets/store-300.png`, 300×300, drawn by `npm run icons` |
| Test deck | `template/validators.pptx`, three slides, written by PowerPoint itself |
| Products | PowerPoint — web, Windows, Mac, iPad |
| Price | Free |

## Long description

> **A library of ready-made slide elements for PowerPoint.**
>
> Open the task pane, find the element you want, and click it. It lands on the
> slide you are on — with its fonts, its colours and its placement exactly as
> they were drawn — and the pane tells you what it did, with Undo one click
> away.
>
> **What is in it**
>
> - 106 elements in each of the two slide sizes: boxes and box layouts, process
>   flows, hierarchies, matrices, tables, triangles, document structure, stamps,
>   labels and markers.
> - Elements that come in sizes — a process flow of one to six boxes, a
>   hierarchy of two to five — are one tile with a stepper, so the list stays
>   short.
> - Every tile shows PowerPoint's own rendering of the element. Rest on one and
>   a card opens with it at full size and a line saying where it will land.
>
> **How it behaves**
>
> - An element lands on the slide you are on, or as a new slide after it —
>   whichever you set behind the gear. Stamps go top right, and a marker lands
>   on the shape you have selected.
> - Colours follow the deck you put the element in, so it looks like your deck
>   and follows your design if you change it later. One setting keeps the
>   library's own colours instead.
> - Undo puts the slide back exactly as it was. PowerPoint's own Ctrl+Z works
>   too.
> - Search matches names, sections and tags, in any word order, and offers what
>   you probably meant when it finds nothing.
>
> **What it does not do**
>
> There is no account, no sign-in and no licence check. It sends nothing
> anywhere: the library is static files served from the add-in's own site, and
> your presentation is read and written inside the task pane on your own
> machine. Nothing about you or your slides reaches us, because there is nowhere
> for it to go.

## Testing notes for the validators

> No account, sign-in, licence key or configuration of any kind is needed. The
> add-in works offline apart from fetching its own library files from its own
> origin.
>
> 1. Download the test deck —
>    <https://ssf-slide-elements.struktureretsundfornuft.dk/validators.pptx>,
>    three slides, nothing confidential on any of them — and open it in
>    PowerPoint, or use any presentation of your own with at least two slides.
> 2. On the **Home** tab, click **Slide elements**. The task pane opens on the
>    library that matches the deck's slide size, with the sections collapsed.
> 3. Open a section — **White boxes** is the first — and click a tile. The
>    element lands on the slide you are on. The footer says how many slides the
>    deck had before and after, and offers **Undo**.
> 4. Click **Undo**. The slide returns exactly as it was.
> 5. Type in the search box (try `flow`), pick a tag, or open the gear beside
>    the search to change where an element lands, whether its shapes arrive
>    grouped, and whether its colours follow your deck or the library.
>
> Accessibility: the pane is operable by keyboard throughout — `/` focuses the
> search, `Esc` steps back out of the card, the gear and the search, and every
> control shows a focus ring. It is tested at 320 px wide in light and dark and
> in high contrast.
>
> Platforms, and what the publisher has and has not measured: the add-in has
> been run end to end on **PowerPoint for the web** (10 September 2026) and on
> **PowerPoint on Windows** (11 September 2026) — the pane, the insert and Undo,
> with the host's own answers recorded. It has **not been measured on Mac or on
> iPad**, because the publisher has neither device. It is built to degrade
> honestly there rather than to assume: the requirement floor is checked at
> runtime and refused with a plain sentence rather than declared in the
> manifest, and touch is first class — a tap opens the preview, a second tap
> inserts, and nothing depends on hover. A finding on Mac or iPad is a first
> measurement rather than a surprise, and we would like to hear it.
>
> Privacy: <https://ssf-slide-elements.struktureretsundfornuft.dk/privacy.html>
> Support: <https://ssf-slide-elements.struktureretsundfornuft.dk/support.html>

## Done since, and how

**The test deck is written**: `template/validators.pptx`, a title slide, an
empty slide to insert onto, and a slide that already has a shape on it. It is
**PowerPoint's own file** — the application was driven over COM with no window
on 2026-09-11 and asked to save it — rather than a .pptx this repo assembled,
because a deck built by this project's own code is the wrong thing to hand the
people checking whether this project's code produces sound files.
`test/validators-deck.test.ts` holds the committed bytes to every claim made
about them here, including that an element actually goes into it.

**And it is SERVED rather than attached**, from 2026-09-24:
`public/validators.pptx`, which is copied verbatim into the site, so the notes
above link it on the add-in's own origin beside the privacy and support pages
they already cite. Partner Center's notes-for-certification field takes text and
nothing else — there is no attachment to be made — and the notes said "attached
to this submission" for a fortnight, which would have sent a validator looking
for a file that was never going to be there. The served copy is held byte for
byte to `template/validators.pptx` by `test/listing.test.ts`, because two copies
of a file is two files and it is the template one that
`test/validators-deck.test.ts` holds to the description above.

## Still the owner's

**The listing NAME is settled: `SSF Slide Elements`**, decided by the owner on
2026-09-16. It is what the manifests have always carried, so nothing in the
build changes; what changed is that it is now a decision rather than an open
question. Recorded as a DECISION and not as a policy clearance — nobody has read
it against certification policy 1100.7, and if a reviewer objects to it the
answer is a rename, not a claim that it was cleared. `docs/DESIGN.md` section 12
carries the same distinction.

This one should not be faked, and it is not a thing this repository can decide:

- **The screenshot** (at least one, 1366×768). It has to be an actual capture of
  the pane open beside a real presentation. A composite of the pane over a drawn
  window would be a picture of something that does not exist, and drawing
  Microsoft's own interface into it would be worse.

  **The CAPTURE is no longer the hard part, and this file used to say it was.**
  Since the display fix of 2026-09-12 the machine has a real 3840×2160 console,
  and on 2026-09-14 a genuine 1366×768 capture was taken unattended: PowerPoint
  sized to the pixel through `MoveWindow`, the pane opened from the ribbon by UI
  Automation and loaded from the live site, captured with `PrintWindow`. Nothing
  composited.

  **The deck was settled on 2026-09-14: an EMPTY presentation.**
  `template/validators.pptx` reads as a test fixture on a store page — its
  slides say things like "An empty slide to insert onto" — and the owner's own
  library is internal content. A blank slide shows the pane, which is the
  product, and nothing that has to be cleared for publication.

**The submission's pictures are committed — FIVE of them**, retaken on
2026-09-24 against build `49890b6`, which the live pane confirmed through its
own `data-build` before the shutter. `test/listing.test.ts` sweeps
`docs/listing-*.png` and holds every one to 1366×768, because a capture at the
wrong size looks right.

| file | what it shows |
| --- | --- |
| [`listing-screenshot.png`](listing-screenshot.png) | browsing the library: the tile grid with PowerPoint's own rendering of each element |
| [`listing-open-all.png`](listing-open-all.png) | the breadth of it, every category open, with the steppers that turn one tile into a sized run |
| [`listing-search.png`](listing-search.png) | search, with the matched words picked out in each name and the sections that have hits as chips |
| [`listing-inserted.png`](listing-inserted.png) | an element on the slide, with the footer reporting the deck delta and Undo offered |
| [`listing-used.png`](listing-used.png) | "Used in this deck", naming the element and the slide it is on |

Together they are the product's own order: browse, search, insert, check, and
see what the deck already holds.
### Alt text for each picture

The submission form asks for alt text per image and it is easy to skip. These
are committed so the next person does not have to invent them again, and so a
retake carries its description with it.

Written as a description of what is IN the picture, for somebody who cannot see
it — not as marketing copy, and not opening with "screenshot of", which a screen
reader already announces.

| file | alt text |
| --- | --- |
| `listing-screenshot.png` | The SSF Slide Elements task pane open in PowerPoint, showing the White boxes category as a grid of element thumbnails. |
| `listing-open-all.png` | The task pane with every category open, showing element thumbnails and the steppers that resize a run from 2 to 5 boxes. |
| `listing-search.png` | A search for "process" in the task pane, with the matched word highlighted in each element name and 5 of 73 results shown. |
| `listing-inserted.png` | A hierarchy diagram inserted onto the PowerPoint slide, with the pane reporting the slide count and offering Undo. |
| `listing-used.png` | The pane's "Used in this deck" list, naming the inserted element and the slide number it sits on. |

Each is under 125 characters, which is where screen readers and listing forms
start truncating.


**The previous single screenshot showed a defect.** It read `73 of 106` with an
empty search box, which says 33 elements are being withheld when nothing is
filtered — they were not withheld, they were behind their run's stepper. The
denominator was counting ELEMENTS against a numerator counting TILES. That was
fixed in `src/pane/render.ts`, and the old picture is exactly what the fix's own
comment cites as the evidence. It also predated the libraries being unified, so
its category counts were stale: `White boxes (22)` where the library now has 26.
A store picture that quotes a number has to be retaken when the number moves.

**Retake them all whenever the pane or the library changes enough to matter.**
The recipe is below and it is a script for that reason; it was retaken three
times on 2026-09-16 alone.

**The two sibling add-ins must be off the ribbon first.** This machine carries
SSF Merge and SSF Charts, and their ribbon groups would otherwise sit in a
picture on a public store page. `scripts/ribbon-cache.mjs` takes them off; back
the file up first and put them back afterwards, because it belongs to
PowerPoint:

```powershell
$cache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\AppCommands\18.0\PowerPoint.RibbonCache.en-GB"
Copy-Item $cache "$env:TEMP\RibbonCache.backup"
node scripts/ribbon-cache.mjs $cache 43ebbbac-44ad-42b2-a582-0ef079093e6c,b7f6d3a2-4c1e-4e8a-9f2b-7d5c0a1e6f43
# ... take the pictures, then:
Copy-Item "$env:TEMP\RibbonCache.backup" $cache -Force
```

PowerPoint reads that file at startup, so it has to be closed while it changes.

**The pane is scrolled** in the pictures that show tiles, so they carry actual
ELEMENTS rather than a column of headings. `listing-shot.ps1` does not scroll;
it is done through the pane's own devtools, which needs PowerPoint started with
`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9444`. At
1366×768 the pane's viewport is 391 px and the header, search box, tag row,
"Used in this deck" link and count come to about 389 of it, so the first tile
sits one pixel below the fold.

### Taking it

```powershell
powershell -ExecutionPolicy Bypass -File scripts\listing-shot.ps1 -Out shot.png
```

Start PowerPoint first. The script makes the empty deck, opens the pane from the
ribbon by name, sizes the window to **1366 × (768 + 48)**, captures it with
`PrintWindow`, and crops the title bar off, leaving exactly 1366×768.

**Cropping the title bar is what removes the account avatar and Microsoft's
"Upgrade your plan" button**, and it is not retouching: every pixel kept is a
real pixel of a real window. Painting them out would be, and is not done.

Two things the script deliberately leaves to a separate decision, because both
change something outside this repository:

1. **Other add-ins' ribbon groups.** The machine that takes the shot also
   carries SSF Merge and SSF Charts, and their groups would otherwise sit in a
   picture on a public page. `scripts/ribbon-cache.mjs` takes them off the
   ribbon; its header carries the format, and `test/ribbon-cache.test.ts` holds
   the rules. **Copy the cache file first and put it back afterwards** — they
   are the owner's tools, and deleting that file rather than editing it removes
   every add-in's ribbon entry including this one, measured, with no rebuild.
   The Office add-ins dialog would be the sanctioned route and does not open
   under automation: four attempts, two routes, 2026-09-14.

   **A STORE add-in puts itself back; the sideloaded ones do not.** Measured
   2026-09-16: the cache was trimmed to this add-in alone with PowerPoint shut,
   and the copy that started next listed two — this one and `wa104380862`, which
   is Script Lab, installed from AppSource. The two sideloaded siblings stayed
   off. Script Lab draws a TAB rather than a ribbon group, it is Microsoft's own
   developer tool, and the committed shot has always carried it; taking it off
   would mean uninstalling it rather than editing a cache, which is the owner's
   call and not worth making. So trim the two GUIDs and expect the store one
   back.
2. **The status bar's language indicator** ("English (Denmark)"). Right-click
   the status bar and untick **Language**. It is a display setting, it does not
   change the editing language, and it did not survive a PowerPoint restart when
   it was tried — so do it last, just before the capture.

Nothing else needs doing. The **build stamp** that used to sit in the pane's
header was taken off it on 2026-09-15 and now lives on the root element as
`data-build` — invisible to a capture, still readable in devtools and still
prefilled into "Report a problem". `docs/DESIGN.md` section 4 carries the
reasoning.
