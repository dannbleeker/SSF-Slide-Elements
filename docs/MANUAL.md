# SSF Slide Elements — manual

A library of ready-made slide elements for PowerPoint. Browse the collection in
the task pane, click an element, and it lands on the slide you are on with every
font, colour and placement intact.

> **Status.** The picker, the insert and taking it back are built, and were run
> end to end in PowerPoint for the web and in PowerPoint on Windows, both on
> 2026-09-11. Mac and iPad have had no round yet. Everything still marked
> *planned* below is designed and not built, and a line moves out of *planned*
> in the same change that makes it true — never before.
> [The backlog](BACKLOG.md) is the order the rest arrives in.

## Contents

- [Before you start](#before-you-start)
- [What it does](#what-it-does)
- [The pane](#the-pane)
- [Adding an element to the library](#adding-an-element-to-the-library)
- [Installing it](#installing-it) — [what needs re-installing](#what-needs-re-installing-and-what-does-not), [which PowerPoint](#why-it-does-not-say-which-powerpoint-it-needs)
- [When something goes wrong](#when-something-goes-wrong)
- [Limits](#limits)

## Before you start

You need PowerPoint — on the web, on Windows or on Mac — and the manifest file
from [Installing it](#installing-it). Nothing is installed onto your machine:
the add-in is a web page PowerPoint opens in a task pane, and the manifest is a
small file that says where that page is.

## What it does

| Piece | What it does | State |
| --- | --- | --- |
| The pane | Opens from **Slide elements** on the Home tab, shows its build, checks the host | built |
| The library | Two curated decks of elements — boxes, flows, tables, markers, stamps — harvested into the catalogue the pane reads | built |
| The picker | Browse the library by section, search it, see where each element lands | built |
| The insert | Drop the chosen element onto the slide you are on, formatting intact | built |
| Taking it back | Remove what an insert added, and nothing else | built |
| A picture of each element | A photograph of the element itself, cut from a PDF print of the library deck | built |
| Colours | An element takes the theme of the deck you put it in, or keeps the library's own — one setting behind the gear | built |
| Used in this deck | Which library elements are already in the open deck, and which slides they are on | built |
| What your slide already has | The preview card shows the shapes already on your slide, in grey, behind where the element would land | built |
| Remove from N slides | Taking a stamp off every slide it is on, in one click | planned |

## The pane

Open it from the ribbon: **Script Lab** is not it — look for **Slide elements**.
The pane shows which build you are running, seven characters in the header on
the right, and then one of three screens.

**Loading the library.** The catalogue is fetched from the add-in's own site.
It is about 190 KB and takes a moment; an element's own markup and pictures are
fetched only when you insert it.

**Slide elements**, once it has arrived. Top to bottom:

- the **getting-started note**, the first time you ever open the pane: three
  lines saying what to do. **Got it** dismisses it for good on that machine;

- a line saying which slide you are on, and — only when your deck is neither
  16:9 nor 4:3 — which library was borrowed for it and what it was scaled to;
- **Search**, which matches an element's English name, its Danish name in the
  owner's deck, its category and its tags. Every word has to match, so
  "white box" finds "White boxes, 2x1 vertical". Press `/` to jump to the box
  and `Esc` to clear it. If a search finds nothing, the pane offers what you
  might have meant — tap one and it searches for that instead. While you are
  searching, the sections that have hits appear as chips with counts: tap one to
  see only that section, tap it again to see them all;
- a line of **tags**, most used first. Picking one narrows the list, and a
  picked tag moves to the front so it stays visible;
- the **gear**, beside the search, holding three settings: whether an element
  lands **onto this slide** or **as a new slide** after it, whether its
  shapes arrive **as one group** or **loose**, and whether its colours follow
  **this deck's theme** or come out **as in the library**. The line at the
  bottom of the pane always says what those are set to, and opens the gear when
  clicked;
- **two links out of the gear.** **Report a problem** opens the support page in
  your browser with the build code, the app and the platform already filled in,
  so a report says which version was running without you looking anything up.
  **Browse the catalogue on the site** opens the element library as a web page,
  which is the easy way to show somebody an element without opening PowerPoint.
  Neither one navigates the pane itself; if PowerPoint refuses to open a
  browser window, the pane says so and gives you the address;
- **colours, in a little more detail.** Left alone — the setting you start on —
  an element takes the theme of the deck you put it in, so it comes out in your
  own colours and follows you if you change the deck's design later. Switch to
  **as in the library** and every colour is fixed to the value it has in the
  library deck, in that element and in any chart it brings with it, and it stays
  that colour wherever the slide ends up. Colours the library states outright
  are the same either way;
- **right-click a tile** — or long-press it on a touch screen — and it offers
  the insert target you are **not** set to, for that one insert: **Insert as a
  new slide** when you are set to land on the current slide, **Insert onto this
  slide** when you are set to a new one. It does not change the setting, and
  `Esc` closes it. Stamps and markers offer nothing, because they always land on
  the slide you are on whatever the setting says;
- **Used in this deck**, which starts as one line: *See what this deck already
  uses*. Click it and the pane reads the open presentation and lists the library
  elements it finds, with the slides each one is on. It reads on request rather
  than every time it opens, because reading means reading your whole deck.
  Elements the add-in inserted are recognised by a mark it leaves on the shapes;
  anything you drew yourself, or pasted in from elsewhere, is not from the
  library and does not appear;
- **Favourites** and **Recent**, when you have any. The star on a tile adds and
  removes a favourite, and the last six things you inserted are remembered;
- the **categories**, collapsed until you open one. Searching or picking a tag
  opens whatever it found. Each tile shows the element's name and its picture —
  PowerPoint's own rendering of it, cut from a print of the library deck. Where
  there is no picture yet the tile keeps the small drawing of where on the slide
  the element lands and how much of it it covers;
- **rest on a tile**, or reach it with the keyboard, and after a moment a card
  opens with the element at full size, its name, and a line saying where it will
  land. The little slide on it also shows, in grey, what your own slide already
  holds — so you can see whether the element would land on top of something.
  Those grey boxes are from the last time the pane read your deck: they appear
  for the slide it read, and not for another one you have since clicked onto,
  and **See what this deck already uses** brings them up to date. In a wide pane
  the card sits beside the list; in a narrow one it lies over the bottom of it.
  `Esc` shuts it;
- an element that comes in several sizes is **one tile with a stepper** naming
  what it counts — "boxes 1 2 3 4 5 6" — so the run does not fill the list. While
  you are searching, the sizes your search did not ask for are greyed, though you
  can still pick them;
- the **footer**, carrying what the last insert did, measured: "12 → 13 slides"
  for a new slide, "12 → 13 → 12 slides, slide 4 replaced" for one onto the
  slide you were on. Beside it, **Again** repeats the last insert and **Undo**
  takes it back;
- and the one primary button, **Insert an element**, which inserts whatever
  tile the keyboard is on. Clicking a tile inserts it directly.

**The library did not load** is the third screen, with **Try again** on it. It
says what happened. That is almost always the network rather than the add-in.

If your PowerPoint is below the floor (see
[which PowerPoint](#why-it-does-not-say-which-powerpoint-it-needs)) the pane
says so instead, and draws no button at all.

If the pane shows its header and nothing else, PowerPoint could not fetch
Microsoft's `office.js` library — see
[When something goes wrong](#when-something-goes-wrong).

### The keyboard

`/` focuses the search box. `Esc` shuts the preview card, then the gear, then clears the search and
the tags. `Tab` reaches the tiles; the arrow keys move between them and `Enter`
or `Space` inserts the one you are on. Every outcome is announced to a screen
reader as well as shown.

### What Undo does, and what Ctrl+Z does

The pane's **Undo** takes back the LAST insert and only that one. For an
element that landed as a new slide it removes that slide; for one that landed
onto a slide it puts the slide it replaced back, exactly as it was.

It is one deep rather than ten, and the reason is worth knowing: putting a
replaced slide back means handing PowerPoint a package containing it, and
holding ten of those means holding ten copies of your presentation inside the
task pane. PowerPoint's own **Ctrl+Z** reverts an insert — measured on
PowerPoint for the web on 2026-09-10 and on PowerPoint on Windows on
2026-09-11 — and that is the deeper history. Press it on the slide canvas
rather than in the pane.

### Where an element lands

Decided by the library rather than by you, per kind of element:

- a **stamp or a label** lands top-right, clear of the edge;
- a **marker, a flowchart shape or an icon** lands on the shape you have
  selected, or in the middle of the slide when nothing is; a marker sized to
  wrap the shape, unless the shape is more than about a third of the slide;
- a **wide part** lands where it sits in the library;
- a **whole-slide element** lands below your slide's own title and is scaled to
  fit the space under it when it would not otherwise. When it already fits, it
  is left exactly where the library put it.

Empty "Click to add text" placeholders are removed when a whole-slide element
lands over them. A placeholder you have typed into is content, and stays.

## Adding an element to the library

The two decks under `template/` are the library: `library-16x9.pptx` and
`library-4x3.pptx`, one per slide size, and both must carry the same elements.
The deck is the authoring surface, and these are its rules:

- **A heading slide starts a category**: a slide with a title and nothing else.
  Everything under it, until the next heading, belongs to that category.
- **Every other slide is one element**, named by its title. The title is the
  element's key and stays in Danish; the English name the pane shows comes from
  `template/names.en.json`, and a title with no entry there fails the harvest,
  all of them listed at once.
- **What is on the slide is the element**: everything that is not layout chrome
  (the title, the footer, the slide number, the date) or an empty placeholder.
  A table or a picture in a content placeholder counts. A shape pushed off the
  slide's right or bottom edge does not.
- **A collection slide yields one element per shape.** Write the line
  `SSF: ét element pr. figur` in the notes of the category's heading slide, and
  every top-level shape on the slides under it becomes an element of its own,
  the slide's title becomes their category, and each is named by its own text
  (brackets stripped, cut at a colon), else by the name PowerPoint's selection
  pane shows when it is not a made-up one like `Gruppe 12`, else by the slide's
  title numbered. `SSF: ét element pr. dias` in a slide's own notes takes that
  slide out again; the English spellings, `one element per shape` and `per
  slide`, work too.
- **Stamps and labels** are the parts of a collection slide whose title
  contains "Stempl": they land top-right. Other parts land at the cursor, and a
  part wider than half the slide lands where it sits.
- **Sizes** need no authoring: elements whose English names differ only by one
  count ("Process flow, horizontal, 3 boxes with table" for 1 to 6) become one
  tile with a stepper by themselves.

Then run `npm run harvest`. It reads both decks and the names file, writes the
catalogue's index to `public/catalogue/catalogue.json` and the elements'
markup and media beside it, and refuses a deck whose keys differ from the other
deck's. Commit the decks, the names file and the index; CI runs the harvest
again and fails when the committed index no longer matches the decks. The
markup and media are not committed: the site builds them from the decks on
every deploy.

The harvest also writes **`public/catalogue.html`**, the page that shows every
element with its picture — <https://ssf-slide-elements.struktureretsundfornuft.dk/catalogue.html>
— so commit that with the index. CI diffs it too.

**Every deck change needs a re-print.** The pane's pictures are cut from a PDF
print of each deck, committed beside it as `library-16x9.pdf` and
`library-4x3.pdf`: print the deck one slide per page with **Frame slides** off
and hidden slides included, commit the print, and `npm run previews` cuts the
elements out of it on deploy. Then stamp it, so the print and the deck are tied
together and CI can tell a stale print from a current one:

```
node scripts/stamp-print.mjs --powerpoint <build> --deck library-4x3
```

On a machine with no screen the print can also be taken by driving PowerPoint
itself — `Presentations.Open(deck, ReadOnly)` then `SaveAs(pdf, 32)` — in which
case stamp it with `--route com`, which records that it was taken that way.

## Installing it

The add-in is a **manifest** — a small file naming a web page — plus the page
itself, which is hosted at
<https://ssf-slide-elements.struktureretsundfornuft.dk>. Nothing is installed
onto your machine.

Download **`manifest-prod.xml`** from the
[latest release](https://github.com/dannbleeker/SSF-Slide-Elements/releases/latest)
once there is one, or take
[the file on `main`](https://github.com/dannbleeker/SSF-Slide-Elements/raw/main/manifest-prod.xml),
which is the same pointer. Both name the same hosted page, so the pane you get
is the same either way — the manifest is only a pointer, and it is the pointer
that is versioned, not the add-in.

| where | how |
| --- | --- |
| PowerPoint on the web | Home → Add-ins → More Add-ins → **My Add-ins** → Upload My Add-in, and pick the file |
| PowerPoint on Windows | Put the file in a folder, share the folder, then File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs and add the share. Restart PowerPoint; the add-in appears under **Shared Folder** |
| PowerPoint on Mac | Copy the file to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` |
| A whole tenant | An administrator deploys the unified manifest — **`manifest-prod.json`** zipped together with `assets/icon-outline-32.png` and `assets/icon-192.png` as an app package — from the Microsoft 365 admin centre |

Then look on the **Home** tab for the **Slide elements** button.

### What needs re-installing, and what does not

Almost nothing. The pane is served from the web, so a change to the pane, the
library or the wording reaches you the next time you open it — no re-install,
nothing to download.

The **manifest** is the exception. Re-sideload only when the manifest itself
changes: the ribbon button, the permissions, the display name, the icons, or the
page the button opens. Those changes are called out in the changelog.

### Why it does not say which PowerPoint it needs

The manifest declares no `<Requirements>` block, deliberately. SSF Slide
Elements needs **PowerPointApi 1.2** — reading the deck and inserting into it —
and that is checked when the pane opens, not declared in the manifest.

A declared requirement set that your PowerPoint does not meet makes the add-in
**vanish from the ribbon** with no message at all: nothing to see, nothing to
report, nothing to search for. The runtime check can tell you which version is
missing and what it costs you, which is worth more than a silent absence.

**The cost of that choice, stated plainly.** Because the manifest declares
nothing, Microsoft's own validator reports the add-in as installable on every
PowerPoint back to **2013 on Windows** — and PowerPoint 2013 does not have
PowerPointApi 1.2. So on an old enough PowerPoint the add-in installs, appears
on the Home tab, opens, and then tells you it cannot run and why. That is the
trade: a message you can read and act on, instead of a button that was never
there.

## When something goes wrong

- **The pane shows its header and nothing else, or a sentence naming
  `appsforoffice.microsoft.com`.** PowerPoint could not fetch Microsoft's
  `office.js` library, which every add-in needs. That is a network or proxy
  rule on your side rather than a fault in the add-in; check that address is
  reachable, then close and reopen the pane.
- **The pane says your PowerPoint does not have PowerPointApi 1.2.** There is
  no way around that one short of a newer PowerPoint.
- **Anything else:** note the build code in the pane's header and write to
  [support](https://ssf-slide-elements.struktureretsundfornuft.dk/support.html).

## Limits

- A tile falls back to a diagram of where the element lands when its picture is
  missing — a new element whose deck has been merged but not yet re-printed.
  See [Adding an element to the library](#adding-an-element-to-the-library).
- The pane follows PowerPoint's theme when it opens; switching PowerPoint's
  theme mid-session needs the pane reopened, because PowerPoint offers no
  theme-change event to a task pane.
