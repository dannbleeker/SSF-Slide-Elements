# The host probe

The six questions in [the design record](DESIGN.md#13-open-questions-for-the-host)
that only a real PowerPoint can answer, asked directly. Nothing in the splice or
the picker should be built on a guess about any of them.

Running it takes a few minutes and leaves your deck as it found it, apart from
one slide it leaves on purpose and takes back on the second run (see
[question 5](#5-does-powerpoints-own-ctrlz-revert-the-insert)), and one
master, one layout and one theme that come in with the probe's own fixture
deck and that no API call can take out again (measured on the web,
2026-09-10). That is why the deck should be a throwaway copy.

## Why a Script Lab snippet and not a probe pane

Every host fact in this repository is borrowed from SSF-Charts and SSF-Merge,
whose rounds against PowerPoint on the web are the reason `CLAUDE.md` says the
host lies about ids, accepts calls it does not perform, and answers differently
on two runs of the same build. A pane that asked these questions would need
hosting and a manifest re-install before the first answer; a Script Lab snippet
needs nothing installed but Script Lab, runs on the web, on Windows and on Mac
the same way, and prints its answer sheet where it can be copied out. It is the
same instrument SSF-Merge used for its six sheets.

## Running it

**Work on a throwaway copy of a deck, not a real one.** The probe adds slides
and removes them again, and on the web that is not undoable in the way you
would expect: PowerPoint for the web has AutoSave permanently on, so every step
is written to OneDrive as it happens. Any deck with at least two slides will do;
for question 4 it should carry a comment, and for question 6 it should be big.

Before you press Run, **click slide 2 in the strip** so one slide is selected.
Question 3 reads what is selected, never sets it, and the reader tells you which
slide number it found so you can compare.

### PowerPoint for the web

1. Open your throwaway deck in PowerPoint for the web, in the **editor**, not the
   read-only viewer. Script Lab cannot load in the viewer.
2. **Home → Add-ins** (or **Insert → Add-ins**), search for
   [Script Lab](https://appsource.microsoft.com/product/office/wa104380862),
   and **Add** it. A **Script Lab** tab appears on the ribbon.
3. **Script Lab → Code**. The editor opens in the task pane on the right.
4. Open the pane's menu (top left of the editor) and choose **New Snippet**, so
   nothing of a previous one is left behind.
5. Make sure the **Script** tab is selected, not HTML, CSS or Libraries. Select
   everything in it and replace it with the whole of
   [`probe/probe-snippet.ts`](../probe/probe-snippet.ts). It is about 70 KB; the
   [raw file](https://raw.githubusercontent.com/dannbleeker/SSF-Slide-Elements/main/probe/probe-snippet.ts)
   is the easiest thing to select-all and copy. Leave the other three tabs
   alone; the snippet needs no libraries beyond the Office.js a blank snippet
   already carries.
6. Click slide 2 in the strip. **Script Lab → Run**. The pane switches to the
   runner. You will see slides appear at the end of the deck and disappear
   again; on a large deck the first arm inserts a copy of every slide, so give
   it a minute.
7. Expand the **console** strip at the bottom of the runner pane.
8. Copy everything between `=== SSF SLIDE ELEMENTS ANSWER SHEET ===` and
   `=== END ===` into a file, `first.json`.
9. The deck now has **one extra slide at the end**, left on purpose. Click the
   slide canvas, press **Ctrl+Z once**, and look at whether the slide went away.
10. **Run** the snippet again, and copy the second sheet into `second.json`.
    The second run removes the extra slide if Ctrl+Z did not, and leaves none
    of its own: it knows it is the second of the pair from a marker the first
    run wrote into the document's settings before it left the slide. In Script
    Lab, "run again" is the runner's own **refresh** button beside the snippet
    name; the ribbon's **Run** does nothing while the runner is already open.

If the console shows `the probe itself failed:` instead of an answer sheet, send
that line too. It is an answer about the host as well.

### Windows and Mac

Same steps, with **Insert → Get Add-ins** (or **Add-ins**) to install Script Lab
and the runner opening in its own pane. Run the web first: it is where the ids
are refused and the collections come back short. iPad cannot be measured by the
owner; the design record says how the first iPad measurement will be taken.

Then, in this repository:

```bash
npm run build:lib
node scripts/read-answers.mjs first.json second.json --save
```

It prints what each answer means and files both sheets under
`docs/host-answers/`, stamped with when they were taken. One sheet alone reads
too; only question 5 needs the pair.

**A question the probe could not put is reported as `unknown`, never as `no`.**
The reader says NOT ASKED and names the re-run that would ask it, because a
sibling once recorded "the metadata scheme needs rethinking" on the strength of
a read that had fallen on a slide the probe never wrote.

## What it asks

| # | Question | What turns on it |
| --- | --- | --- |
| 1 | Does `insertSlidesFromBase64` accept a package **pruned to one slide** whose other parts are still present but unlisted? | Whether the engine can send one slide back by unlisting the rest (cheap) or must remove every other slide properly, orphans walked (`Pkg.removeSlide`, slow on a big deck) |
| 1b | Does inserting a slide on the deck's **own master** add a second master? | Whether every click grows the deck by a master and a theme, and which formatting option to pass |
| 2 | Does insert-after-current, then a **positional delete** of current, keep the order the engine expects? And can a slide the run just added be a `targetSlideId`? | The whole "onto this slide" mechanic, and "as a new slide" twice in a row |
| 3 | Does `getSelectedSlides()` name the slide the user is looking at, at the position `<p:sldIdLst>` gives it? | How the pane knows which slide to splice into |
| 4 | Which read of the deck, `getFileAsync` or `exportAsBase64Presentation`, and what does each **drop**? | Which read the engine builds on; a dropped comment part is a comment the user loses |
| 5 | Does PowerPoint's own **Ctrl+Z** revert an insert? | Whether the pane's Undo must stay out of the way of the host's |
| 6 | How long does a read take on a big deck, and is the **floor** met? | The two-second budget in the design, and the floor message on hosts below 1.2 |

### 1. The pruned package

Three fixture decks, built by the engine and embedded in the snippet, differ in
one thing each: **listed** has two slides both listed, the package as PowerPoint
writes it; **pruned** has the second slide's `<p:sldId>` and its relationship
removed, with the part and its content type still in the zip; **unlisted** has
only the `<p:sldId>` removed, the relationship still pointing at the part. Each
pruned arm lists one slide, so one landing is the answer, two means the host
walked the relationships instead of the list, and a refusal means the package
must be cleaned properly.

A **control arm** runs first, while the deck is still only the user's own: it
inserts the presentation's own bytes, read back through `getFileAsync`. That
deck is a package PowerPoint wrote seconds earlier, so it cannot be malformed,
and a host that refuses it is refusing insertion itself. `insertionBlame` is the
reading, and it says CANNOT TELL rather than guessing when the control did not
run. The control runs twice, under `KeepSourceFormatting` and
`UseDestinationTheme`, with the master count read before and after each, which
is question 1b.

### 2. Insert, then delete by position

The engine's "onto this slide" is: insert the rebuilt slide after the current
one, prove the deck grew, then delete the current one by position, so the copy
takes its place. The arm does exactly that with slides of its own: P1 stands in
for the current slide, S for the copy, P2 for the slide behind. It first aims S
at P1 **by id**, which is a slide the run just added, the case a second "as a
new slide" insert meets and one the web is documented to refuse for
`getItem(id)`. If that is refused the arm says so and aims at the user's last
slide instead, so the order half is still asked. Ids are read by position
throughout, never resolved.

### 3. The selected slide

Read-only. `getSelectedSlides()` is PowerPointApi 1.5; on a host without it the
reader says the pane will have to ask. The selected ids are looked up in a
positional read of the deck, and the positional ids' first halves (an API slide
id is `<sldId id>#…`) are compared with the file's own `<p:sldIdLst>`, inflated
from the `getFileAsync` bytes in the browser. That comparison is what lets the
engine turn a selected id into the `<p:sldId>` it splices.

### 4. Which read

`getFileAsync` is the floor's read and hands back the whole deck;
`exportAsBase64Presentation` (1.10) hands back the slides asked for, and
SSF-Merge's sixth sheet found it **drops comment parts and `ppt/authors.xml`**
on the web. This arm exports every slide in the deck and compares part NAMES
against the same deck read through `getFileAsync`. Names only, never content,
because the sheet is written to be pasted into an issue. **It needs a deck with
a comment on it**; without one there is nothing to drop and the reader says so.

### 5. Does PowerPoint's own Ctrl+Z revert the insert?

The probe cannot press Ctrl+Z, so this takes two runs. The first leaves one
slide at the end of the deck, tagged in the package, and says so on the console.
You press Ctrl+Z once on the slide canvas and run the snippet again; the second
run looks at the last five slides for the tag. Gone, with the deck exactly one
smaller than the first run left it: Ctrl+Z reverts an insert, and the pane's
Undo must not fight it. Still there: it does not, and the second run removes
it by position. A second run cannot tell a Ctrl+Z from a hand delete, which is
why the steps say to press it and not to delete. On a host below 1.3 the tag
cannot be read and the reader falls back to the counts.

**How the second run knows it is the second.** The slide is the one thing a
successful Ctrl+Z removes, so it cannot also be what tells the runs apart: on
the first web round the second run found no tag, took itself for a first run,
and left a slide of its own. So before the first run leaves the slide it writes
a marker into the document's settings (`Office.context.document.settings`),
which live outside the undo stack and survive the Ctrl+Z the slide does not.
The marker carries the deck size just before the slide was left, so a second
sheet on its own can answer the question; the pair is the cross-check. The
second run clears the marker, so the run after it is a first run again.

The marker goes in **before** the slide, never after. A settings save is
reported to disable undo on Excel
([office-js#3141](https://github.com/OfficeDev/office-js/issues/3141)), and a
save after the insert would take your one Ctrl+Z away from the very insert
this question is about. PowerPoint on the web measured otherwise on 2026-09-10
(a save after an add still left Ctrl+Z on the add), but the order costs
nothing and holds on hosts not yet measured.

### 6. Timing and the floor

Every read is timed and sized: `getFileAsync` at the start and at the end,
the export in the middle. The reader prints megabytes, milliseconds and the
rate; a number from a round carries the date of its sheet. The requirement
sets the host reports are printed against the floor in `src/host/capability.ts`
and against each call the probe needs, so a host that lacks `getSelectedSlides`
or the export is named rather than inferred.

## What it does to your deck

It appends slides after the last one (`targetSlideId`; without it the host
inserts at the FRONT), asks its questions, and removes what it added by
**position**, never by id. The sweep is clamped so it can never remove more
than the probe added, never more than the deck grew, and never reach an index
below the deck's size when it started. Those clamps are tested, and each is
proven load-bearing by removing it and watching the test go red.

The one slide it leaves is for question 5, and the reader's clean-up line
accounts for it. If the deck ends any other size than it started plus that
slide, the reader warns rather than pretending it is clean. The package is
read again at the end, so a master, a theme or a layout an insert left behind
is listed too. On the web that list is not empty: the fixture decks carry
their own master, `KeepSourceFormatting` brings it in once (the second fixture
insert finds it already there), and the sweep can remove slides but not
masters, layouts or themes. The first run also writes one entry into the
document's settings, the marker for question 5, which the second run clears.

## What it has answered so far

**PowerPoint for the web, 2026-09-10** — two sheets under `docs/host-answers/`
(`2026-09-10T17-34-14-659Z.json` and `2026-09-10T17-41-19-168Z.json`, host
`0.0.0.0`, platform `OfficeOnline`, PowerPointApi up to 1.10, a run of about
seven seconds on a three-slide deck): every insert landed; both prunings land
as one slide, so the host reads the slide list and ignores an unlisted part;
the deck's own master is not duplicated under either formatting option; a
slide the run had just added is accepted as `targetSlideId`, and
insert-then-positional-delete keeps the order; `getSelectedSlides` named slide
2 at its file position; the export drops the comment part and
`ppt/authors.xml` (and, once a foreign master is in the deck, that master, its
layout and its theme); a 34 KB `getFileAsync` took under a second; and
**Ctrl+Z reverts an insert** (four slides down to three, the tag gone). Those
sheets predate the marker: the second one took itself for a first run and left
a slide behind, which is the defect the marker fixes.

**Same host, same day, with the marker** — a second pair
(`2026-09-10T18-13-10-866Z.json` and `2026-09-10T18-29-13-985Z.json`): the
first run wrote the marker and left its slide, Ctrl+Z took the slide, and the
second run found the marker, found no tag, cleared the marker and left
nothing, the deck ending at the three slides it started with. The second sheet
reads "yes" on its own, from the marker. Two things about that pair are about
the minute, not the host: the document session had been through a
session-timeout reload, and every read was slow (`getFileAsync` 40 s for
40 KB where the first pair measured under a second), so one fixture insert
landed only after the probe's 120 s budget and is recorded as a timed-out
insert that landed, not a refusal.

**PowerPoint on Windows, 2026-09-11** — one pair
(`2026-09-11T09-54-40-421Z.json` and `2026-09-11T09-57-30-334Z.json`, platform
`PC`, host `16.0.20326.20132`, PowerPointApi up to 1.10, runs of 9.4 and 8.1
seconds against a 4-slide, 14.8 MB deck carrying one comment). **Every question
answered the way the web answered it**: both prunings land as one slide; the
deck's own master is not duplicated under either formatting option while the
fixture deck's is; a just-added slide is accepted as `targetSlideId` and
insert-then-positional-delete keeps the order; `getSelectedSlides` named slide
2, the slide that had been clicked, at its file position; the export drops the
comment part and `ppt/authors.xml` (and the three `ppt/webextensions/` parts);
and **Ctrl+Z reverts an insert** — one press took the deck from 5 slides to 4
with the tagged slide gone, and the second run cleared its marker and left
nothing. What differs is speed: `getFileAsync` read 14.13 MB in 2.8 to 3.0
seconds where the web's healthy reading was 34 KB in 874 ms, and no call came
near the 120 s budget.

`docs/host-answers/` is the count, not this line; Mac is still to come.

## One answer is not evidence about your host

It is evidence about your host **in that minute**. A sibling project has
question-by-question records of the same build answering differently minutes
apart. If an answer decides something expensive, run the probe twice.
