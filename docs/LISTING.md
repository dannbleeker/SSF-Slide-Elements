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
| Provider | StruktureretSundFornuft |
| Short description | Ready-made slide elements, dropped onto the slide you are on. |
| Support URL | <https://ssf-slide-elements.struktureretsundfornuft.dk/support.html> |
| Privacy URL | <https://ssf-slide-elements.struktureretsundfornuft.dk/privacy.html> |
| Terms | Microsoft's standard EULA |
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
> - 117 elements in each of the two slide sizes: boxes and box layouts, process
>   flows, hierarchies, matrices, tables, timelines, stamps, labels, markers and
>   flowchart shapes.
> - Elements that come in sizes — a process flow of one to six boxes, a
>   hierarchy of two to five — are one tile with a stepper, so the list stays
>   short.
> - Every tile shows PowerPoint's own rendering of the element. Rest on one and
>   a card opens with it at full size and a line saying where it will land.
>
> **How it behaves**
>
> - An element lands on the slide you are on, or as a new slide after it —
>   whichever you set behind the gear. Stamps go top right, markers and
>   flowchart shapes go where your selection is.
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
> 1. Open the test deck attached to this submission — three slides, nothing
>    confidential on any of them — or any presentation of your own with at least
>    two slides.
> 2. On the **Home** tab, click **Slide elements**. The task pane opens on the
>    library that matches the deck's slide size, with the sections collapsed.
> 3. Open a section — **White boxes** is the first — and click a tile. The
>    element lands on the slide you are on. The footer says how many slides the
>    deck had before and after, and offers **Undo** and **Again**.
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

## Still the owner's

Two things cannot be produced without a screen and a real PowerPoint, and
neither should be faked:

- **The screenshot** (at least one, 1366×768). It has to be an actual capture of
  the pane open beside a real presentation. A composite of the pane over a drawn
  window would be a picture of something that does not exist, and drawing
  Microsoft's own interface into it would be worse.
- **The listing NAME.** `docs/DESIGN.md` section 12: nothing here has been read
  against the naming policy, and the sibling SSF Merge is held on the same
  question. That answer decides this one.
