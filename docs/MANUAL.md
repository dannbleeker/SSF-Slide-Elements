# SSF Slide Elements — manual

A library of ready-made slide elements for PowerPoint. Browse the collection in
the task pane, click an element, and it lands on the slide you are on with every
font, colour and placement intact.

> **Status.** Very little of that is built yet. What you can install today is
> the pane itself: it opens, shows which build it is, checks that your
> PowerPoint clears the floor, and says plainly that there is nothing to insert
> yet. Everything marked *planned* below is designed and not built, and a line
> moves out of *planned* in the same change that makes it true — never before.
> [The backlog](BACKLOG.md) is the order it arrives in.

## Contents

- [Before you start](#before-you-start)
- [What it will do](#what-it-will-do)
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

## What it will do

| Piece | What it does | State |
| --- | --- | --- |
| The pane | Opens from **Slide elements** on the Home tab, shows its build, checks the host | built |
| The library | A curated deck of elements — boxes, flows, tables, markers — harvested into the pane | planned |
| The picker | Browse the library by section, search it, see a thumbnail of each element | planned |
| The insert | Drop the chosen element onto the slide you are on, formatting intact | planned |
| Taking it back | Remove what an insert added, and nothing else | planned |

## The pane

One step today, **Start here**. It shows which build you are running — seven
characters in the header, on the right — and one button, **Insert an element**,
which is disabled until there is a library to insert from. Above it the pane
says why: there is nothing to insert yet, and the element library arrives in a
later release.

If your PowerPoint is below the floor (see
[which PowerPoint](#why-it-does-not-say-which-powerpoint-it-needs)) the pane
says so instead, and draws no button at all.

If the pane shows its header and nothing else, PowerPoint could not fetch
Microsoft's `office.js` library — see
[When something goes wrong](#when-something-goes-wrong).

## Adding an element to the library

- *Planned.* The library deck is the authoring surface: open it in PowerPoint,
  draw the element on a slide, give the slide a title — that is the name the
  picker will show — and save. A script harvests the deck into the catalogue
  the pane ships, and CI fails a change where the deck and the catalogue
  disagree.
- *Planned.* The layout says what a slide is for, the title placeholder is the
  name, and everything that is not a placeholder is the element. No sidecar
  file and no naming convention.

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

- It inserts nothing yet. See [What it will do](#what-it-will-do).
- The pane follows PowerPoint's theme when it opens; switching PowerPoint's
  theme mid-session needs the pane reopened, because PowerPoint offers no
  theme-change event to a task pane.
