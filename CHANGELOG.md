# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **The tag line has its chevron.** It shows one row of tags and the chevron at
  its right opens the rest, which is what the pane was always meant to do.

- **A stamp on several slides at once.** Select more than one slide in the
  slide strip, then click a stamp or a marker: it lands on every slide you
  selected, and the footer says how many it reached. If a step cannot be
  confirmed the run stops there and says how far it got, so you always know
  what the deck holds. The pane cannot undo this one — PowerPoint's own
  `Ctrl+Z` can — and it says so when it finishes. Selecting one slide, or
  clicking a whole-slide element, works exactly as before.

### Fixed

- **The pane now names a PowerPoint you can move to when it cannot run.** On a
  PowerPoint too old for the add-in, the one screen you get said what was
  missing and what it cost you, and left you to work out what to do about it.
  It now names the versions that have it: a current Microsoft 365, PowerPoint
  2021, or PowerPoint on the web.

- **A question about removing an element no longer survives the insert that
  takes its tile away.** With six elements in Recent, opening "Remove from N
  slides" on the oldest one's Recent tile and then inserting anything else left
  the question drawn nowhere at all — and with a question notionally open, the
  Remove button was hidden on every tile, with nothing on screen saying why or
  how to get out of it. Inserting now closes the question, as searching,
  starring and picking a category already did.

- **The count above the list no longer reads as though something were
  filtered.** With an empty search box it said "73 of 106" — a count of tiles
  over a count of elements, two different things, since an element that comes in
  several sizes is one tile with a stepper. The fraction could never reach its
  own denominator. It now counts the same thing on both sides, and shows just
  the number when nothing is filtered.

- **The settings line at the bottom of the pane now announces itself as the
  control it is.** It opens and shuts the same options panel the ⚙ does, but a
  screen reader heard a plain button with no state and no link to the panel, so
  there was no way to tell whether pressing it had opened or closed anything.
  Both controls now say which panel they open and whether it is open.

- **The arrow keys no longer jump to the top of the list from wherever you
  are.** Pressing an arrow anywhere that was not a tile — the gear, a category
  heading, a tag chip, the size numbers, the star, the Insert button — threw
  the focus to the first tile and scrolled the list back to the top. It also
  meant the arrow keys could not scroll the pane at all. They now belong to the
  tiles, and to the search box on the way down into them; everywhere else they
  scroll, as they do on any other page.

- **Closing something puts you back where you were.** Escape out of a tile's
  right-click menu, out of the question before a removal, or out of the gear
  used to drop the focus onto nothing at all, so the next Tab started again at
  the top of the pane. It now goes back to the tile, the Remove button or the
  gear you opened it from. "Keep them" does the same.

- **A right-click while a removal question is open no longer does nothing, and
  then something.** The pane took the right-click, suppressed Windows' own menu
  and showed nothing — and the menu then appeared by itself later, on a
  different tile, when you pressed Escape to answer the question. The
  right-click is now left to the browser while a question is up.

- **An element inserted while the add-in was still measuring your deck no
  longer takes its pictures from the wrong library.** The tiles go live before
  the deck's shape has been read, and picking one straight away could give an
  element from the widescreen library a picture, chart or workbook from the 4:3
  one — silently, with the wrong thing on the slide — or stop with "the
  catalogue has no part", naming a part it does have. An insert now keeps the
  library it started in, whatever the deck turns out to be.

- **Taking an element off every slide no longer claims nothing else changed
  when it did.** The removal rebuilds each slide without the element and then
  takes the original away. If the second half failed, the pane said "The rest
  are as they were — try again" — but that slide now had both your original,
  element and all, and a copy without it. Pressing "try again" added another
  copy each time. It now says the deck has a slide too many, and tells you to
  look before trying again.

- **The pane no longer says it is inserting while it is undoing.** Pressing
  Undo, or taking an element off every slide it is on, showed "One insert at a
  time. This one is still going." directly above its own "Undoing…" — and put
  an "Inserting…" badge on the tile while the add-in was in fact taking a slide
  back out of your deck. On the web these take seconds, so both sentences were
  on screen long enough to read. It now names what it is actually doing.

- **The Insert button is no longer live with nothing behind it.** The pane
  remembers the last element you inserted into a deck. If a later version of
  the library no longer had that element, the button came back enabled with
  nothing selected and no explanation, and pressing it did nothing at all —
  no message, no sign anything had happened. It now treats a remembered choice
  the library has lost as no choice, and says so.

- **The keyboard no longer loses its place every time you insert.** Inserting
  greys out every tile while it runs, and the pane could not put the keyboard
  back on a tile that was greyed out — so it fell to the top of the page and
  stayed there, even after the insert had finished. The next arrow key took you
  back to the very first element in the library instead of leaving you near the
  one you had just used, and anyone using a screen reader lost their place in
  the list entirely. The pane now waits until the tile can take the keyboard
  again, and gives it back.

- **Half the tags were unreachable.** The tag line drew only the first twelve
  of the library's twenty-four, so the rest could not be picked as filters at
  all — and the only thing that unfolded the line was opening the settings,
  which has nothing to do with tags. Every tag is now there, and the chevron is
  what opens them.

- **"Remove from N slides" now puts the keyboard where the question is, and
  reads it out.** Asking the question hides the Remove button on every tile —
  including the one you just pressed — so the keyboard was left at the very top
  of the pane, and reaching the answer meant tabbing past the search box, the
  settings, every tag, and every element before it. Anyone using a screen
  reader was told nothing at all. The question now takes the focus when it
  opens and is announced, which matters most here because it is the one thing
  the pane does that it cannot undo for you.

- **Reordering slides while an insert is running no longer deletes the wrong
  one.** Inserting onto the slide you are on rebuilds it and then removes the
  original, and it found the original by its position — a position worked out
  before the insert started. Dragging a slide in the thumbnail strip while the
  insert was still going changed which slide sat at that position without
  changing how many there were, so the check that guarded the removal passed and
  your dragged slide was deleted instead. The pane reported success, because the
  count was exactly what success looks like. It now checks that the slide it is
  about to remove is still the one it meant, and if the deck has moved it leaves
  both slides in place and says so.

- **A new slide no longer arrives carrying the previous slide's animations, or
  its transition.** "As a new slide" copies the slide you are on and empties
  it, but a slide's animation timeline is stored beside its shapes rather than
  among them, so it was left behind — pointing at shapes that had just been
  removed. Because the numbering of the new slide's shapes started again from
  the top, those pointers did not go nowhere: they landed on the element you
  had just inserted. An entrance animation written for one of your old shapes
  would then be applied to the new element, which in PowerPoint means it is
  hidden until the animation runs — so you could ask for an element and get a
  slide that looks empty. The slide you are on is untouched either way: an
  insert onto your own slide keeps your animations, because they are yours.

- **An element whose chart names its workbook with a space in it now inserts.**
  A file name is written one way inside a PowerPoint file and stored another,
  and the add-in read the two spellings differently — so an element carrying a
  file with a space, a comma or a hash in its name was reported as missing from
  the library and refused. Nothing in today's library is named that way; this is
  about the libraries to come.

- **"Browse the catalogue on the site" no longer claims it failed.** On every
  PowerPoint that opens the page through a plain browser window, the add-in
  reported "PowerPoint would not open a window" over a tab it had just
  opened.

- **The line under the header keeps the slide number when PowerPoint answers a
  read only partly.** It used to blank the number as though nothing were
  selected, where the same failure arriving as a timeout correctly left the
  last number standing.

- **A failed insert that finds the deck shorter says so even when PowerPoint
  also reported an error.** The error used to be the only thing mentioned, in
  the pane's mildest wording, over a deck that had lost a slide.

- **Preview pictures update when the library decks change.** The address a
  preview is fetched from carries a version that only changed when an
  element's name, size or position did — so a deck edit that changed how
  something LOOKS, and nothing else, left you seeing the old picture.

- **The pane says what went wrong out loud.** A failed insert and a failed
  Undo were shown in the footer but never announced, so a screen reader user
  was told nothing on the one failure where the deck may be holding an extra
  slide.

- **Keyboard focus stays where you put it.** Pressing the settings line at the
  bottom of the pane moved the focus to the gear button at the top, past the
  search box and the whole list.

- **Un-starring an element no longer strands the pane** while its "Remove from
  N slides" question is open on the Favourites tile.

- **The arrow keys work in the search box again.** Pressing Left or Right to
  fix a typo moved the focus onto the first tile instead of the caret — and
  since a tile is a button, the next Enter inserted that element. Left and
  Right now belong to the search box; Down still steps out of it onto the
  tiles.

- **A long press on a touch screen no longer inserts the thing behind the menu
  it just opened.** Holding a tile opened the menu, and then the same gesture's
  own click closed it again and inserted onto the slide you were on — the very
  target the menu is there to change.

- **Undo is no longer left armed after "Remove from N slides".** Pressing it
  would have put the removed element back on one slide and reported "Undone."

- **An insert that leaves the deck shorter than it started no longer tells you
  to delete a slide.** The slide it named was the one your element had just
  landed on.

- **The add-in copes with a PowerPoint that names the selected slide
  differently from the deck's own list.** Where that happens, the pane could
  not tell which slide you were on and refused every insert for the rest of the
  session.

- **Elements from the 4:3 library land where they should, and no longer bring
  an invisible object into your deck.** The 4:3 library deck had been through
  think-cell, which leaves an invisible frame in the corner of every slide it
  touches. The add-in was treating that frame as part of the element: it
  stretched 42 elements' measurements to the corner of the slide — so they were
  placed and cropped as if they were nearly slide-sized — and it copied
  think-cell's invisible object, and the file behind it, into your presentation
  on every insert. Nothing you could see, and 0.7 MB of it across the library. Shapes
  PowerPoint does not draw are now left out of the library, and the 4:3
  elements carry 122 files where they carried 215.

- **A deck written by another tool is read the way the rest of the add-in
  reads it.** Where a deck referred to one of its own files with a doubled or
  trailing slash, the add-in worked out a name the file does not have and then
  quietly treated that file as missing — which could mean a new slide sharing
  the previous one's speaker notes, or an element placed against the wrong
  layout. PowerPoint does not write those, but other tools do.

- **An empty group is no longer left behind when you remove a part you had
  grouped with something of your own.** If the group was one PowerPoint had
  written its own bookkeeping onto — which it does to a group in a shared,
  co-authored deck — taking the element out left the group standing with
  nothing in it.

- **"Used in this deck" no longer forgets an element that is still there.**
  Inserting the same element twice onto one slide and then pressing Undo took
  the slide off the list entirely, so the pane reported nothing from the
  library in a deck holding it — and the "Remove from N slides" button went
  with the row. The earlier copy is still on the slide, and the list now says
  so.

- **The add-in uses far less memory while reading a deck.** Reading what a deck
  already uses kept part of every slide that carries a tag in memory for the
  rest of the session, so the cost grew with the length of your presentation.
  Measured on the library deck: 44 held down to 1, and now flat however long
  the deck is.

- **A failed insert that finds the deck SMALLER than it was now says so.** It
  used to report the size the deck had before and add that nothing had
  changed — neither of which was true of a deck that had lost a slide.

- **A tag can no longer be written over the wrong part of your deck.** If a
  shape's bookkeeping reference pointed at something that was not a tag file —
  a slide layout, say — the add-in would have overwritten it, and every slide
  on that layout would have lost its design. It now checks what the reference
  leads to first.

- **A new slide no longer arrives carrying your own table or picture.** "As a
  new slide" copies the slide you are on and empties it, so the new one keeps
  the same design. It emptied the text on it, but a table or a picture you had
  dropped into one of the layout's own boxes was left standing — so an element
  meant for a fresh slide arrived on top of your figures, and the figures were
  on a slide you had not put them on. The new slide now starts from the layout,
  as it says it does; your original slide is untouched either way.

- **"Did you mean" no longer suggests something your filters would hide.** With
  a tag or a category chip picked, a suggestion under "Nothing matches that."
  could be an element those filters exclude — so taking the way out of the dead
  end put you straight back in it, with the same message and the same
  suggestion underneath. Suggestions now come from what your filters actually
  leave.

- **The question before "Remove from N slides" no longer strands the pane.**
  The question is asked on the element's own tile, and while it is up the
  Remove button is hidden on every tile. Searching or closing a category while
  it was open took the tile — and the question with it — off the screen, and
  left no Remove button anywhere with nothing on screen to say why. Anything
  that takes the tile away now cancels the question with it.

- **The line under the header keeps naming the slide you are on, even if the
  add-in could not read your deck when it opened.** One failed read at startup
  — a slow network, a very large presentation — stopped the pane listening for
  you changing slides for the rest of the session, and nothing said so: the
  line just went on showing whatever it said at the time.

- **A marker dropped on a selected LINE is no longer invisible.** Markers wrap
  the shape you have selected, and a straight line has no thickness to wrap, so
  the marker came out with no height at all — impossible to see and hard to
  click on in order to delete. It now lands on the line at its normal size, the
  same as it does on a shape too big to wrap.

- **A part that could not be fetched no longer reports itself as missing from
  the library.** If the network failed while the add-in was collecting a
  picture or a chart an element needs, the message said the library does not
  have that file — which sends you looking for a broken add-in instead of
  trying again. It now says what actually happened, and only says a file is
  absent when the site says it is.

- **A category you picked while searching no longer keeps filtering the
  library after you delete the search.** The category chips appear only while
  there is something in the search box, so backspacing the box empty left the
  pick in force with no chip on screen to lift it — the library showed one
  category and nothing said why. Clearing the box now lifts the pick, which is
  what the Clear button and Escape already did.

- **The pane no longer refuses to open over a setting it cannot read.** If
  anything in the add-in's saved state for a deck was not the shape the pane
  expected, the pane failed while drawing its first screen — and because the
  value stayed saved, it failed again on every open. Whatever it cannot read is
  now simply treated as not set.

- **The preview card no longer draws a wildly oversized grey box for a table
  you have grouped with something else.** The card sketches what your slide
  already holds; a grouped table was measured by the table's own columns
  instead of the group's size on the slide, which on a scaled group could come
  out several times wider than the slide itself.

- **A comment on the slide you insert onto is no longer lost.** Inserting an
  element rebuilds the slide you are on, and a comment written by PowerPoint
  2016 or 2019 — or by any deck not yet upgraded to the newer kind — went with
  the slide it was on. Newer comments were never affected, which is why this
  went unnoticed: the two are stored differently, and only the older kind was
  dropped. Your reviewer's thread now survives an insert either way.

- **A stamp or a marker no longer lands on a blank slide of its own** when the
  gear is set to "As a new slide". That setting is meant for whole-slide
  elements only, and the right-click menu already treated it that way; the
  setting did not, so a stamp inserted with it switched on landed alone on an
  empty slide wedged after yours, and your own slide kept nothing. A stamp
  always lands on the slide you are on, as the preview card says it will.

- **The pane no longer tells you to delete a slide it already deleted.** When
  PowerPoint reported a problem tidying up after an insert but had in fact
  tidied up, the footer said "delete slide N by hand" — and slide N was by then
  the slide your element had just landed on. It now counts the deck to find out
  what happened rather than believing the error, so it only asks you to delete
  something when there really is something to delete.

- **A new slide no longer carries a broken reference to a comment.** Making a
  new slide from one that had a comment on it left the slide pointing at a
  comment that had been taken off it, and that pointer could end up aimed at
  the add-in's own bookkeeping instead. Nothing was visible to you; the file
  was untidy in a way PowerPoint is entitled to complain about.

- **"Move to a new slide" no longer loses the element instead of moving it.**
  The move takes the insert back and makes it again on a new slide. If
  PowerPoint reported a problem on the way but had in fact done the work, the
  pane stopped after the first half — so the element was gone from the deck
  altogether, under a message saying the undo had failed. It now checks the
  deck's own size, the way the other half of the same step already did.

- **"Remove from N slides" no longer leaves a copy behind and calls the deck
  untouched.** If PowerPoint reported a problem on a step it had actually
  completed, the run stopped and told you nothing had changed, while the deck
  had gained a slide with the element on it twice. Each step is now judged by
  what the deck is, not by whether an error was reported.

- **A failed insert no longer claims the deck is untouched when it may not
  be**, and no longer leaves an Undo armed that points at the previous insert.
  If the pane cannot find out what happened after handing the slide over, it
  now says so and tells you to check the end of the deck. A failure that
  happens before anything is asked of PowerPoint is unchanged: it still says
  the insert was refused, and an Undo you already had still works.

- **An element that failed to load once is no longer stuck for the rest of the
  session.** If the network dropped while the add-in was fetching an element,
  every later click on that tile replayed the same failure without trying
  again, and the only way back was closing the pane and reopening it. It now
  fetches again on the next click, which is what the carried-parts half of the
  same cache already did.

- **The preview picture of a rotated stamp is no longer cut into.** The two
  rotated stamps are masked to their own turned frame so the slide around them
  does not show in the corners, and that mask was being turned as though the
  page were square. On a 16:9 page it came out skewed by more than half the
  stamp's own height, so the mask clipped the stamp instead of framing it.
  Previews are rebuilt on release, so this arrives with the next one.

- **"Used in this deck" no longer sends you to the wrong slide.** Each slide
  number in that list is a button that takes you there. Inserting an element as
  a new slide pushes everything after it one slide along, and the list was not
  told — so a row still naming slide 5 took you to whatever slide 5 had become,
  and said it had worked. The numbers now move with the deck, and move back
  when the insert is undone.

- **The keyboard can reach the tiles again.** Tab, the arrow keys and Enter are
  meant to move between elements and insert the one you are on. They did not
  work at all: the moment a tile took focus the pane redrew itself, which
  removed the very button you had just reached, so Tab could not get past the
  first tile and the arrow keys never moved. Focus now survives the redraw.

- **Elements land on the slide on decks that are not the library's size.** The
  library is drawn at one slide size, and an element inserted into a deck of a
  different size kept the library's own coordinates instead of being scaled to
  fit — so it could sit well off the edge of the slide, while the pane's
  preview reported it as landing correctly. A stamp went about an inch past the
  right edge and a white box two and a half inches, on PowerPoint's ordinary
  ten-inch "On-screen Show (16:9)" deck. That size has the same 16:9 shape as
  the library, so the pane never even said it was borrowing one.

- **An element that would sit on top of your title is moved off it**, even when
  it was already small enough to fit below. The rule that keeps library
  elements clear of a taller title only ever resized them, so one that already
  fitted was left exactly where it was — on top of the title, which is the one
  thing that rule exists to prevent.

- **The preview card no longer promises an insert that cannot happen.** When
  PowerPoint will not say which slide you are on, the card said an element
  would land on the first slide. Nothing does that: the insert stops and asks
  you to click a slide. The card now says the same thing.

- **Escape now clears the whole search, and the clearing sticks.** Pressing
  Escape left the category you had picked still filtering the library — and the
  category buttons are only on screen while a search is, so there was nothing
  left to unpick and nothing saying why most of the library had gone. Clearing
  the search also was not remembered, so the search you had just got rid of came
  back the next time the pane opened.

- **Your footer, slide number and date are left alone.** The add-in treated a
  slide's running furniture as though you had put it there: an otherwise empty
  slide with a footer on it was described as already holding something, so the
  pane offered to move your element onto a slide of its own for no reason — and
  a blank footer or date placeholder was removed from the slide the insert
  rebuilt, which is part of your deck, not one of the "Click to add text"
  ghosts an insert is meant to tidy away.

- **A failed Undo no longer offers to run itself again.** Undo puts your
  original slide back first and takes the rebuilt one away second, so if
  PowerPoint stops between the two your deck holds both. The pane said only
  "Undo did not work" — which reads as though nothing happened — and left the
  Undo button armed. Pressing it again put your slide back a second time and
  reported success, leaving you with two copies of it and the element gone. It
  now says what it could not finish and where to look, and takes the button
  away rather than letting it act on a deck it can no longer describe.

- **The pane keeps answering after a reading fails.** If PowerPoint did not
  answer when the pane asked which slide you were on, and you clicked another
  slide while it was waiting, the pane could go on naming the slide you had
  left. It now asks again.

### Fixed — the host probe

- The answer to "which read of the deck, and what does each drop" no longer
  names only the comments and the authors part. On the sheet of 2026-09-10
  18:13 it read "the export DROPS ppt/authors.xml and 1 comment part(s)" over
  twelve parts gone — the other ten being a whole slide master, its layout and
  its theme, three task-pane registration parts, and two revision parts. The
  full list was printed below it, but the verdict is the line that gets quoted.

- And on a deck with no comments it no longer says "there was nothing for the
  export to drop" while parts were dropped. It still cannot answer the
  question, which is about comments; it just no longer claims more than that.

- The answer sheet no longer writes "-1 slide(s) landed anyway" when an insert
  both raised and left the deck SHORTER. It now says the deck lost a slide,
  which is the opposite fact and the one worth reading.

- Question 3 no longer reports a hard refusal when the selected slide simply
  sits past the 120 positions the probe reads. That is the probe's own cap, not
  the host declining to answer, and putting it on the sheet as "no" would have
  settled the question the wrong way. The sheet now says "unknown" and asks for
  a re-run with a slide inside the range.

- Question 1's reading no longer answers from an arm that did not run. A sheet
  missing one of its two pruned inserts was graded as though that insert had
  been refused, so the sheet carried a firm instruction about how the engine
  must build its package — drawn from nothing. It now says which arm is
  unanswered and asks for it to be re-run.

## [0.1.0] - 2026-09-16

The first release. Everything below is what the add-in does the day it ships:
a library of 106 ready-made slide elements in each of the two slide sizes,
inserted onto the slide you are on, with Undo one click away.

### Changed

- **The publisher is now named StruktureretSundFornuft ApS**, with the company
  form, on the add-in itself and on the privacy, terms and licence pages. It is
  the registered entity, and the sibling SSF Merge already said so. If you have
  sideloaded the add-in, **re-install it** — this is a manifest change, and the
  manifest version went to 1.0.0.1 with it.

### Added

- **Terms of use and a licence page**, both on the add-in's own site. The terms
  say what the add-in is allowed to do to your presentation — it rewrites one
  slide when you insert, and writes a small marker beside each element so
  "Used in this deck" can find its own work later — and the licence page says
  what you may do with the slides you make. Short version: they are yours, with
  nothing to attribute and no condition attached. Until now the terms were
  Microsoft's generic ones, which could not say either of those things.

### Removed

- **The Scales is gone** from Stamps and labels. The library now holds 106
  elements in each slide size. As with anything else removed, a copy already on
  one of your slides stays exactly where it is.

- **The Flowchart shapes category is gone**, and the ten small shapes in it:
  Document, Documents, Start / stop, Process / action, Decision, two labelled
  arrows, a dashed connector, a waste bin and a box with a footer strip. The
  library now holds 107 elements in each slide size rather than 117.

  Anything you already put on a slide stays exactly where it is — those are
  ordinary PowerPoint shapes now and the add-in has no hold on them. If a
  presentation of yours uses one, **Used in this deck** still lists it, as
  "an element from an older version of the library", so nothing disappears
  from that list either. You just cannot insert them again.

### Changed

- **A category now looks like something you can open.** Each heading has a small
  mark that turns when it opens. It was always a button — a screen reader has
  always said "collapsed" — but on screen it was a line of bold text with
  nothing to suggest you could press it.

- **The first time the add-in sees a presentation, it opens the top category.**
  Until now you were met with a search box, a row of tags and a column of closed
  headings, and nothing to look at until you guessed that a heading opens. It
  still remembers which ones you leave open, for that presentation — including
  when you leave them all closed.

- **The search box shows the `/` that jumps to it.** The shortcut was always
  there and written only in the manual.

- **"Again" has gone from the bar at the bottom.** It repeated the last thing
  you inserted — which is the first tile under **Recent**, a few lines further
  up, with a picture on it. Click that instead. The bottom of the pane is the
  narrowest part of it, and two buttons for one action was one too many.

- The short code in the corner of the pane's header is gone. It said which build
  of the add-in you were looking at, which matters when something has just been
  fixed and you want to know whether you have the fix yet. It is still there for
  that — **Report a problem** puts it into your report automatically — it simply
  is not printed on the pane any more.

### Added — nothing you can see

- The screenshot for the add-in's store page can now be taken by running one
  command, instead of being assembled by hand. It makes an empty presentation,
  opens the add-in beside it, and captures the window at exactly the size the
  store asks for. Nothing in the picture is retouched — the parts that should
  not appear on a public page are left out of the frame rather than painted
  over.

### Changed — nothing you can see

- The check that deliberately breaks the add-in's code learned a seventh way to
  break it: reading a comparison backwards, so that "does this element fit in
  the space" becomes "does the space fit in this element". That is the mistake
  behind a whole family of layout bugs in our sister project, and it is the kind
  a test can miss while still checking the edges. It made 81 such changes and
  **every one of them was caught** by a test that already existed.

- We asked PowerPoint the last of the seven questions we had never been able to
  put to it, and it answered. There are two ways for an add-in to read the deck
  you have open, and one of them quietly leaves your **comments** behind — and
  the people who wrote your comments with them. This add-in uses the other one.
  That was believed on the strength of a sister project's measurement; it is now
  measured here, on PowerPoint for Windows, against a presentation with a real
  comment in it.

### Added — nothing you can see

- A small presentation for testing the add-in against a real PowerPoint, kept
  with the project. One of the seven questions we put to PowerPoint asks what
  each way of reading your deck quietly leaves behind — and a comment is the
  thing most worth not losing. The last round could not answer it, because the
  presentation it was run on had no comment in it. This one does.

  PowerPoint wrote the presentation itself, and it is checked into the project
  with the name of whoever saved it taken out: PowerPoint stamps the signed-in
  account into every file it saves, and this project is public.

### Changed — nothing you can see

- The check that deliberately breaks the add-in's code now covers the last part
  of the engine: the splice, which puts an element into a copy of your slide and
  takes one back out. Fifty-three changes to it went unnoticed. Thirty-six are
  now caught, thirteen alter nothing and are written down with the measurement
  behind each, three are already refused by the type checker, and two were lines
  that could never run.

  Two of the thirty-six were defects waiting to happen rather than gaps in the
  abstract. A straight horizontal line has no height, and one guard was the only
  thing stopping an insert from giving the owner's line a height it never had.
  Another checks that a claim to be an empty placeholder is made by a SHAPE: a
  picture carrying that claim would otherwise have been read as leftover
  furniture and deleted — your image, removed by an insert. And a third keeps
  the add-in from mistaking PowerPoint's animation bookkeeping for a shape, which
  would have numbered every shape it adds from two billion upward.

### Changed — nothing you can see

- The check that deliberately breaks the add-in's code now covers the part that
  READS AND WRITES the PowerPoint file itself — the zip of XML parts every
  insert opens, edits and hands back. Seventy-seven changes to it went
  unnoticed. Thirty-eight are now caught, by 28 new cases; the other
  thirty-nine alter nothing at all and are written down with the measurement
  behind each.

  Nothing you can see changes, because nothing was broken. What was being held
  by nobody, and now is: the id the first slide of an empty deck gets, which the
  format fixes at 256; the counter that stops a new part from reusing a name a
  deleted part had, and which must not let one kind of part advance another
  kind's numbering; relationship ids starting at 1 and going up one at a time
  with no gaps; a part name written the absolute way staying absolute; a part
  with no file extension not being handed the content type of a rule that names
  none; and the deck's bytes going to the zip reader without being copied first,
  which on the file route means a copy of your entire presentation.

  Most of the thirty-nine that alter nothing are guards around a foreign XML
  library, checked against what it actually returns rather than what it is
  assumed to: 1,080 real parts from the two library decks holding 469,288
  elements and 571,356 attributes, plus 20,000 synthetic ones. They stay,
  because each is what makes the line under it honest.

### Changed — nothing you can see

- The check that deliberately breaks the add-in's code now double-checks the
  changes it believes were CAUGHT, not only the ones it believes slipped
  through. It always re-ran a change nothing noticed, slowly and against
  everything, before reporting it — and took a change something noticed at its
  word on the first try. That is the wrong way round: a change wrongly reported
  as slipping through wastes a reader's afternoon and says so out loud, while a
  change wrongly reported as caught is a gap in the tests that the report stays
  silent about. One was found exactly that way. A change is now put back to the
  one test file that objected, and only a second objection counts.

### Changed — nothing you can see

- The check that deliberately breaks the add-in's own code, one change at a
  time, to see whether the tests notice now covers the part that READS the
  library decks — turning a PowerPoint file into the elements the panel offers
  you. Seventy-three changes to it went unnoticed. Sixty-two are now caught, by
  47 new cases; ten alter nothing at all and are written down with the proof for
  each; and one was a line of code that could never run, which is gone.

  Nothing you can see changes. What was being held by nobody, and now is not:
  which shapes on a slide count as content rather than the layout's own
  furniture, so a placeholder holding a single character is not mistaken for an
  empty one; how a picture or a chart placeholder is told apart from an empty
  one; which parts of the file an element drags along with it, and what happens
  when the deck points at one that is not there; how an element's box, its
  rotation and its table columns are read; how a name spelled "one row" or "six
  rows" becomes the count behind it; and where an element's picture is cut from
  the printed deck.

- The tool doing that breaking was wrong about a quarter of its own work, and
  said nothing. Of the 373 comparisons it thought it was flipping, 272 were not
  comparisons at all — they were type annotations, which the test runner throws
  away before anything runs, and shifts, which are a different operator wearing
  the same characters. Every one was a run that could only ever report "not
  noticed", and every one of one file's ten findings was exactly that. The tool
  now requires a real comparison, which is checked by tests of its own.

- And it could record a change as NOTICED when nothing had noticed it. It gave
  every unnoticed change a second, slower check before believing it, and gave
  the noticed ones no check at all — so a test that failed for its own reasons,
  on a machine with too much else running, counted as the change being caught.
  One was found this way: a change to the engine reported as caught by the first
  run and, on a re-run against a larger set of tests, not caught at all. It
  alters nothing observable and is now written down as such. The asymmetry that
  hid it is recorded next to it; closing it is the next piece of work.

- The same tool could also stop dead without saying so. One of its test runs
  wedged, and because nothing set a time limit the whole sweep sat waiting on it
  for three hours and forty minutes while the report simply stopped growing — a
  stalled run is indistinguishable from a slow one to anyone reading the output.
  Runs now have a ten-minute limit, more than ten times the slowest honest one,
  and a run that hits it counts as no answer rather than as a pass.

### Changed — nothing you can see

- Three small pieces of code the tests could not tell apart from any other
  version of themselves: one that shortened a message, one that compared two
  words, and one that read a slide's name. All three do exactly what they did
  before; two are now simpler, and one is checked where it was not.

### Changed — nothing you can see

- The add-in's own tests are now checked by deliberately breaking the code, one
  change at a time, and seeing whether they notice. Out of 349 such changes, 56
  went unnoticed. Forty-five of those are now caught — one of them by deleting a
  line of code that turned out to do nothing — and the remaining eleven are
  changes that alter nothing at all, no matter what you do.

  Nothing you can see changes, because nothing was broken — but several things
  that keep the add-in honest were being held by nobody. Among them: the limit on
  how much of a PowerPoint error message is shown to you, which exists so that a
  failure cannot dump an entire presentation into the panel as text; what the
  panel says when PowerPoint hands back more slides than it was asked for; how
  long the add-in keeps re-checking a slide count that PowerPoint is slow to
  update; how it decides your theme is dark rather than light; and how it reads
  the answer sheets that every measured claim about PowerPoint's behaviour is
  built on.

  The eleven that alter nothing are written down with the proof for each, so the
  next person to run this does not spend an afternoon working out again what this
  one worked out — and the tool now says when one of those notes has gone out of
  date rather than staying quiet about it.

### Changed — nothing you can see

- One of the build's own checks re-runs the whole test suite to count it, and it
  did that in a mode that writes every result to a file and nothing to the
  screen. So when it failed, the log showed a stack trace from the tool that
  started it and not one word about which test had gone wrong — which is exactly
  what happened twice this afternoon. It now reads the results back and names
  them, and says plainly when a run failed with nothing in it failing, which is
  what a build machine running out of memory looks like.

### Fixed — the manual now warns you about the one thing that loses an element's mark

- Cutting a shape out of one slide and pasting it onto another drops the mark
  the add-in uses to recognise its own elements, on PowerPoint for the web. The
  element is untouched and looks the same, but it stops appearing under **Used
  in this deck** and *Remove from N slides* can no longer reach it. That was
  known and written down for the people building the add-in, and nowhere a user
  would look. It is in the manual now, with what to do about it.

### Changed — nothing you can see

- The hard-won rules about how PowerPoint behaves — the ones that each cost
  somebody a day to find — were prose in a file for maintainers, and most of
  them had nothing stopping a future change from quietly breaking one. Six of
  them are now checks that fail the build, each one proven to fail by breaking
  the code it protects. The three that no check could ever hold are written down
  as exactly that, rather than left looking gated.

### Changed — nothing you can see, again

- The numbers the add-in works out while inserting — how many shapes it added,
  how many empty boxes it cleared away, where exactly the element landed — are
  now checked against the slide it actually produced. They drive what the pane
  tells you and what it offers, and nothing had been holding them to the truth.
  They were all correct.

### Fixed — an element could vanish from "Used in this deck" straight after you inserted it

- **Reading the deck takes a while, and the pane deliberately stays usable while
  it does — so you can insert something in the middle of a read.** When you did,
  the read finished afterwards with an answer from before your insert and quietly
  put it back: the list said "Nothing from the library is in this deck yet" about
  an element you had just watched land, and the little slide on the preview card
  went back to how it looked before. The read now notices the deck changed under
  it, throws its own answer away, and tells you to ask again.

### Changed — nothing you can see, again

- The code that actually puts an element on your slide, takes one back off, and
  measures what is already there now gets tested against slides that are not
  shaped the way it expects — including, for the first time, a slide whose XML
  has been through a formatter, which is how any file that has been opened in an
  editor looks. It coped with all of them; nothing had been checking.

### Changed — nothing you can see, again

- The largest file behind the pane was one file doing six jobs, and finding
  anything in it meant knowing where it happened to be. It is now four, each
  named after the thing it decides: the picker's search, the preview card, the
  "Used in this deck" list, and everything else. Nothing about the pane changed.

### Changed — nothing you can see, again

- The part of the engine that reads a .pptx now gets tested against files that
  are damaged in the ways real files are: a picture declared twice, a link with
  nothing on the end of it, a slide the presentation lists but cannot find. It
  already coped with all of them; nothing was checking that it did.

### Fixed — a marker would have stopped being a marker

- **A future category whose name merely contained the word "mark" would have had
  every element in it resize itself around whatever shape you had selected.**
  Only markers are supposed to do that. Nothing in today's library triggered it,
  so there was nothing to see; the test for it was written the wrong way round
  and lived in a part of the code the build does not check.

### Changed — nothing you can see, again

- More rules moved out of the one file the build does not measure and gained
  tests: which elements wrap the selection, which slide an insert landed on,
  when "Move to a new slide" is offered, and how the pane decides whether
  PowerPoint's theme is light or dark. No behaviour changed.
- Two explanations in the code described the wrong thing entirely, having come
  loose from what they were written about.

### Changed — nothing you can see, again

- Rules about what the pane remembers, what Escape shuts and where an arrow key
  moves were written inside the one file the build does not measure. They moved
  somewhere it does, and gained tests. No behaviour changed.

### Changed — what the store listing tells Microsoft's reviewers

- **The testing notes now say which platforms have not been tested.** The
  add-in has been run end to end on PowerPoint for the web and on Windows; it
  has not been run on Mac or iPad, and the notes say so rather than leaving a
  reviewer to find out. A build check keeps that list honest: it is read off
  the recorded test results, so it cannot claim a platform was tested when it
  was not, or keep calling one untested after it has been.

### Changed — nothing you can see, again

- A build check that the pane actually has the controls its design record names.
  Four things the record described were never built, and nothing noticed until
  someone read the two side by side.

### Added — the list comes back where you left it

- **How far down the list you had scrolled is remembered**, per presentation,
  like the search and the tags. Reopen the pane and it scrolls back as soon as
  the tiles are drawn — and if you have already started scrolling by then, it
  leaves you where you are.

### Changed — nothing you can see

- Tests for parts of the engine that had none: leaving one slide listed in a
  package, and reading a deck's theme colours when the deck is not the tidy
  case. No behaviour changed. The build's own coverage floors went up to match.

### Changed — the pane now remembers each presentation separately

- **What you were doing is remembered per presentation**, not per machine. The
  search you typed, the tags you picked, the categories you left open, the size
  you picked on an element that comes in several, the last six you inserted and
  the three settings behind the gear all come back when you reopen that deck —
  and none of them follow you into the next one. Your starred elements still do,
  because a star is about the library rather than about a deck.
- **The search box and the tags are remembered at all**, which they were not
  before. That was the half of this the pane had never had.
- One presentation is told from another by a number worked out from its address.
  The address itself is not written down. A presentation you have not saved yet
  has no address, so unsaved decks share one memory between them.

### Added — put an element on a slide of its own after all

- **"Move to a new slide"** in the footer, when you have just dropped a
  whole-slide element onto a slide that already had something on it. It takes
  the insert back and makes it again as a slide of its own, so the slide you
  started with is yours again exactly as it was. Your own title does not count
  as "something on it", and neither does an empty "Click to add text" box,
  because the insert removes those anyway — so the offer turns up when the
  element really is likely to be sitting on top of your work.

### Added — open every section at once

- **"Open all" beside the count** opens every category in one click. The
  sections still start closed, because a list of six categories opened is a lot
  of scrolling before you have said what you are looking for. The link goes
  once they are all open, since there would be nothing left for it to do.

### Fixed — the pane in Windows high contrast

- **The little orange mark above the heading was invisible if you use Windows
  high contrast**, because its whole appearance is its colour and high contrast
  replaces colours with your own. It is drawn in your text colour there now, so
  it is there like everything else. Nothing changes for anybody not using high
  contrast.
- The weekly check that measures the pane now renders every screen in high
  contrast too, which is how this was found. It had been promised and never
  looked at.

### Added — click a slide number to go there

- **The slide numbers under "Used in this deck" are links now.** Click one and
  PowerPoint moves to that slide; the pane says "Slide 4" once PowerPoint
  confirms it is there, and if PowerPoint does not confirm it the pane says so
  and asks you to click the slide in the strip instead. Nothing in your deck
  changes either way.
- On Windows and Mac the jump can be refused while the notes pane has the
  focus; click the slide canvas first. On a PowerPoint too old for the call the
  numbers stay plain text, as before.
- The probe asks a seventh question, so the next round on any platform measures
  the jump directly.

### Changed — inserting into a big presentation is much faster

- **An insert into a 45 MB deck used to cost about four seconds of the add-in's
  own work. It now costs about a quarter of a second.** Nearly all of that time
  was converting the presentation to and from the text form PowerPoint hands it
  over in; the add-in now lets the browser do that conversion instead of doing
  it itself.
- Nothing about what lands on your slide changed.

### Added — a deck to try it on

- **A small test presentation ships with the project**, for anyone reviewing the
  add-in: a title slide, an empty slide to insert onto, and a slide that already
  has a shape on it. Nothing on it is confidential.

### Fixed — one right-click, one menu

- **An element can sit in Favourites, in Recent and in its own section at the
  same time.** Right-clicking one of them opened the little menu on all three —
  and the question before removing a stamp appeared three times too. Each now
  belongs to the tile you clicked.

### Fixed — two found by hunting rather than by using

- **The grey boxes on a deck that is not 16:9 or 4:3.** The rectangle showing
  where an element just landed was measured against the library's slide size
  instead of your deck's. On a deck that borrows the nearest library — A4,
  16:10, anything custom — it was drawn in the wrong place. The card's little
  slide is now your deck's shape too.
- **Removing a stamp after the deck moved on.** If slides were added, deleted or
  reordered between asking what the deck uses and pressing Remove, the pane
  worked from the older list. It now asks the deck it is about to change.

### Changed — inserting a marker over and over stays cheap

- **Every insert of an element that carries a picture used to put another copy
  of that picture into the presentation the add-in is working on.** It now uses
  the one that is already there. On the test deck the second and later inserts
  of a marker cost **1.9 KB instead of 11.6 KB** of the package that has to be
  handed to PowerPoint and back on every single insert — which is the slowest
  part of an insert, so a long session of stamping stays as quick as the first
  one.
- **Your saved file was never carrying those copies:** PowerPoint merges
  identical pictures when it saves, and we measured it doing so — four copies in,
  one out. This is about the work per insert, not about the size of your file.
- A chart or an embedded workbook is still copied per insert, deliberately —
  two charts sharing one workbook would mean editing one edits both.

### Fixed — an element you have grouped is still an element

- **Group something the add-in inserted with a shape of your own, and the pane
  stopped seeing it.** "See what this deck already uses" answered that the
  element was not in the deck while it sat on the slide in front of you, and
  "Remove from N slides" was never offered for it. Grouping is one gesture and
  the mark the add-in leaves is still there; the pane now finds it, inside your
  groups and inside groups within them.
- **Removing one takes the element out of your group and leaves your own shape
  where it was.** If the element was the only thing left in that group, the
  group goes too — an empty group is not something PowerPoint makes.
- **Ungrouping one no longer loses it either.** An element that arrives as a
  group carried its mark on the group alone, so ungrouping it to change one box
  took the mark with it: five shapes still on the slide, and the pane no longer
  able to see any of them. The mark now goes on the shapes inside as well, which
  is what the loose setting has always written. It adds about a quarter of a
  kilobyte per shape and nothing you can see.
- That fix covered elements the add-in groups for you, and **23 elements in the
  two libraries are drawn as a group already** — the stamps among them. Those
  behaved the old way until a sweep over every element in both libraries found
  them: ungrouping one left its shapes on the slide with nothing to find them
  by, and **Remove** then said the slide carried no such element. Every element
  is now covered, and the sweep runs on every commit.

### Fixed — a removal with nothing left to remove says so

- **If you take an element's shapes off by hand and then press Remove, the pane
  used to answer "Removed from 0 slides." and call it a success.** It now says
  the element is not on any slide any more and that nothing changed. Nothing
  went wrong in that case and nothing is left for you to finish, so it says
  neither.

### Fixed — the privacy page can now count

- **The page that tells you what the add-in keeps on your machine said "four
  things" and "two settings". It keeps five and the gear holds three.** The page
  had been right when it was written; two of the things it now keeps were added
  after it. The wording is corrected, and the suite refuses a page whose numbers
  disagree with what the code actually stores, so it cannot drift again.

### Changed — the toolchain

- Seven development dependencies moved forward, including vitest to 5. Nothing
  a user installs changed: none of them reach the add-in itself.
- **Three things the add-in does had never actually been run by the suite**, and
  now are: taking the "Click to add text" ghosts off a slide a whole-slide
  element lands on, emptying a placeholder when an element goes in as a new
  slide, and the fast route the pane itself takes when it converts a
  presentation. The first two were covered by a test that passed for the wrong
  reason. Nothing about them changed — they are simply checked now.
- **The suite now refuses code written for its own test.** A sweep lists every
  export the shipped add-in never calls, and the build fails on one that is not
  recorded as deliberate. It found an unused copy of a function the splice was
  documented to use, and a list of the pane's screens that had stopped being
  the list the documentation check reads — so a fourth screen could have been
  added with nothing noticing its heading was undocumented. Nothing a user sees
  changed.

### Added — taking a stamp off every slide it is on

- **A stamp or marker already in your deck now carries "Remove from N slides"**,
  under its tile, once you have asked what the deck uses. It asks before it does
  anything, names the slides it would touch, and says that the pane cannot undo
  it.
- It only ever takes off what the add-in put there. A shape you drew yourself
  carries none of its marks, and neither does another add-in's.
- If a step cannot be confirmed the run stops there, leaves the rest of the deck
  alone, and the footer says how far it got.
- **Run against a real PowerPoint on 2026-09-13**, on the web: a stamp onto
  three slides, listed as being on all three, and taken off all three, with the
  deck read back afterwards to check rather than taking the pane's word for it.

### Added — see what is already on your slide

- **The preview card now shows your own slide's shapes in grey**, behind where
  the element would land, so you can see whether it would cover something.
- They are what the pane saw the last time it read your deck: they show for that
  slide, and **See what this deck already uses** brings them up to date.

### Added — the other insert target, without changing the setting

- **Right-click a tile** — or long-press it on a touch screen — and it offers to
  insert the other way for that one insert: as a new slide when you are set to
  land on the slide you are on, and the other way round. Your setting stays as
  it was.
- Stamps and markers offer nothing, because they always land on the slide you
  are on whatever the setting says.

### Added — what this deck already uses

- **The pane can tell you which library elements are already in the open deck**,
  and which slides they are on. It is one line until you click it, because
  finding out means reading your whole presentation — so it reads when you ask
  rather than every time the pane opens.
- Inserting adds to the list and Undo takes it back out, without re-reading.
- An element from an older version of the library is still listed, and says that
  is what it is, rather than quietly going missing.

### Added — two ways out of the gear

- **Report a problem** opens the support page in your browser with the build
  code, the app and the platform already filled in. Only those three: nothing
  about your presentation goes with it, and the page shows nothing it does not
  recognise.
- **Browse the catalogue on the site** opens the element library as a web page.
- If PowerPoint will not open a browser window, the pane says so and gives you
  the address rather than doing nothing.

### Fixed — the same element, the same colours in both slide sizes

- **The two library decks were on different colour palettes**, and 79 of the 117
  elements came out in different colours depending on your deck's shape: the
  same box Office orange on a widescreen deck and light blue on a 4:3 one. The
  4:3 deck now uses the same palette as the 16:9 one, so an element looks like
  itself whichever size you are working in.

### Added — the library, on the web

- **Every element is now on the site**, in both slide sizes, with PowerPoint's
  own rendering of each one:
  <https://ssf-slide-elements.struktureretsundfornuft.dk/catalogue.html>. It is
  generated from the same catalogue the pane reads, so it cannot fall behind the
  library, and it is linked from every page of the site.

### Added — an element in the library's own colours

- **The gear now holds a colour setting.** Left where it starts, an element takes
  the theme of the deck you put it in, which is what makes it look like your
  deck. Switch it to **as in the library** and the element keeps the colours it
  has in the library instead, wherever the slide goes afterwards.
- The setting covers a chart an element brings with it, not only its shapes, so
  the two cannot come out in different colours.
- Colours the library states outright were never theme colours and are unchanged
  by either setting.

### Added — a note the first time you open the pane

- **The first time you open it, three lines say what to do**: rest on a tile to
  see the element, click to place it, and where Undo and the gear are. Dismiss it
  once and it stays dismissed, on that machine.

### Added — narrowing a search

- **While you search, the sections that have hits appear as chips with counts.**
  Tap one to see only that section, tap it again to see them all. The counts stay
  put when you pick one, so you can always see what is in the others.
- **A sized element greys the sizes your search did not ask for.** Search for
  "3 boxes" and the stepper says which of its six numbers you meant. They are
  still pickable — the size exists, you may still want it.

### Added — a search that finds nothing offers a way out

- **Mistype a name and the pane suggests what you probably meant**, as chips you
  can tap. Every suggestion is a name the library really has, so picking one
  always finds something. It matches against each WORD of a name, not just the
  whole of it, so "triangel" reaches "Triangle, simple, with text at the
  corners".

### Added — a closer look before you pick

- **Rest on a tile for a moment and a card opens with the element at full size**,
  its name, and a line saying where it will land — which follows your insert
  setting, except for stamps and markers, which always land on the slide you are
  on whatever the setting says. Keyboard focus opens it too. Escape shuts it.
  In a wide pane it sits beside the list without covering anything.

### Added — every tile shows the element itself

- **A tile is a picture of the element now, not a diagram of where it lands.**
  PowerPoint's own rendering, cut out of a print of the library deck, so what
  you pick from looks like what you will get. The little landing diagram stays
  as what you see while a picture is still on its way, or if one is missing.

### Added — a print cannot go stale unnoticed

- **Each committed PDF print now records which deck it came from**, and the
  build refuses when the two drift apart. Edit a deck and forget to re-print it,
  or replace a print without recording it, and you are told which of the two
  moved and what to do — rather than finding out later from element pictures cut
  from a file that no longer exists.

### Changed — the library speaks English

- **Every placeholder an element puts on your slide is now English.** `[ Tekst ]`
  reads `[ Text ]`, `[ Overskriften på kassen]` reads `[ Box heading ]`, and so
  on through 76 phrases and about 3,400 pieces of text across both decks. What
  lands on your slide now matches the language of the pane you picked it from.
- Deliberately unchanged: the dummy filler (`Lorem ipsum`, `xxxx`) which is meant
  to look obviously fake, and text that was already English.
- **A handful of elements changed their internal id**, because an element in a
  collection — the stamps, the labels, the flowchart shapes — is identified by
  its own text. If you starred one of those as a favourite before today, the star
  will not have followed it, and "Used in this deck" will not recognise a copy
  you inserted earlier. Re-star it and it sticks.
- Searching for **"meeting"** now finds four elements it had been missing.
- Both PDF prints were retaken from the translated decks.

### Added — the printed library

- **Both library decks are now committed as PDFs beside the decks themselves**,
  110 pages and 108. This is the raw material the tiles' pictures get cut from;
  the cutting itself is still to come, so the pane looks the same today. The
  prints are of the decks as they stand, which still carry their Danish
  placeholder text — they will be retaken once the decks are edited.

### Fixed — the library decks open in PowerPoint again

- **The two library decks under `template/` would not open in PowerPoint on
  Windows.** It offered to repair them instead, and repairing one threw away ten
  of the embedded objects it carried. Four characters were missing from a list
  inside each file that tells PowerPoint what its parts are; they are back. No
  slide, note, picture or embedded object changed, and the catalogue the pane
  reads is byte-for-byte the one it was already reading — so nothing you see in
  the pane moves. This only ever affected opening the decks by hand; the add-in
  itself was never reading them at run time.

### Measured — PowerPoint on Windows

- **The add-in has now been run on PowerPoint for Windows, not only in the
  browser.** Every behaviour the pane depends on works there the same way, and
  nothing had to change to make it so. Four rounds of picking an element and
  pressing Undo put the slide back exactly as it was, shape for shape, whether
  the element landed on the slide you were on or as a new slide after it.
  PowerPoint's own Ctrl+Z still takes an insert back, and the line under the
  header still keeps up when you click quickly through the slide strip.
- **Windows is faster than the browser, by a lot.** Reading a 14 MB
  presentation takes about three seconds where the browser needed as long for a
  file a thousandth of the size, and the pane no longer has to wait for the
  slide count to catch up after an insert — on Windows it is right immediately.
  The waiting the pane does for the browser costs nothing on Windows and stays
  put, because the browser still needs it.
- Mac and iPad have still had no round, and the manual says so.

### Added — the splice, the host handshake and the picker

- **The product does what it says on the tin: open the pane, click an element,
  and it lands on the slide you are on.** Three pieces arrived together, because
  none of them is worth anything alone.

- **The splice** (`src/core/splice/`) puts a library element's markup into a
  copy of the destination slide, in the file. It renumbers every shape id
  against the slide it is joining, repoints every relationship reference —
  anchored on the relationships NAMESPACE, so `r:embed`, `r:id` and `r:link` are
  all reached and a spelling nobody thought of is not silently skipped — copies
  every part the element carries into the package under a name the deck is not
  using, walks each of those parts' own relationships and rewrites them too, and
  declares a content type for every one. Then it lands the element where
  `docs/DESIGN.md` section 5 says: a stamp top-right, a marker around whatever
  shape you have selected, a whole-slide element below your own title and scaled
  into the space under it when it would not otherwise fit — including the
  table's grid, because PowerPoint draws a table from its rows and columns and
  ignores the frame around them. Every insert writes a tag naming the element
  and the catalogue it came from, in the package, before the insert, because a
  slide the run just added does not round-trip through the API.

- **A sweep over all 117 elements of the committed 16:9 library**, into a real
  deck, on every commit: every package is put through
  `scripts/package-integrity.mjs` and every shape id checked for collisions.
  `docs/BACKLOG.md` asked for exactly that. Both guards were proved by breaking
  the code and watching them go red.

- **The host handshake** (`src/office/powerpoint.ts`, `src/host/insert.ts`)
  reads the deck with `getFileAsync`, inserts with a `targetSlideId`, and takes
  the replaced slide away **by position**, with the deck counted before, after
  the insert and after the removal. The count is the evidence and the raise is
  not: a call can raise and still have done the work, and a call that raises
  nothing has not necessarily happened. Every sentence the footer shows is
  computed from those three numbers.

- **The picker** (`src/pane/`): search over the English name, the owner's Danish
  name, the category and the tags; a tag line; collapsible categories; a tile
  per element drawing where on the slide it lands; one tile with a stepper for
  an element that comes in several sizes; favourites and recent; a gear holding
  the insert target and whether shapes arrive grouped; a footer with the
  measured slide count, **Again** and **Undo**; the keyboard (`/`, `Esc`,
  arrows, Enter); a live region; Windows high contrast; and the pane reopening
  where you left it. Loading, failure and floor states all say what happened and
  what to do.

### Changed

- **Undo is one deep, not ten.** Taking back an insert that landed onto a slide
  means handing PowerPoint a package containing the slide it replaced, and ten
  of those is ten copies of your presentation inside a task pane. PowerPoint's
  own Ctrl+Z reverts an insert — measured on the web on 2026-09-10 — so the
  deeper history already exists. `docs/DESIGN.md` section 6 records the change
  and the reason.

- **The catalogue index now carries each carried part's content type**, read
  from the library deck rather than guessed from its extension at insert time. A
  part with no content type declared is a package PowerPoint refuses without
  saying which part.

- **The privacy page said nothing is stored on your device. Four things are.**
  The pane keeps your gear settings, which elements you have starred, the last
  six you inserted and which categories you left open, in the browser's local
  storage. The page had said otherwise since before the picker existed, and it
  was live on the site saying it. It now lists all four, names the key they sit
  under, says they never leave the device and that clearing site data removes
  them, and says plainly that the section was wrong until today. A test reads
  the page against the code, field by field, so a fifth thing stored without a
  word on the page turns the build red.

  The page had carried its own discipline in the same paragraph — "If that ever
  changes, this section changes with it" — which is the sentence nobody comes
  back to. That is the second false claim about user data found in this repo in
  one day, after the security page's, and both were written the same way: a
  promise with nothing executing it.

- **The site's front page said the add-in inserts nothing yet.** It inserts.
  The lockstep guard now reads the two public pages as well as the manual, the
  README and the security page, which is where it should have been looking.

- **The security page said the add-in reads nothing and writes nothing.** It
  reads your whole presentation and writes it back, which is what an insert is
  here, and that page is the one people read when deciding whether to trust the
  tool. It now says so, along with what is kept afterwards and for how long.
  The sentence had survived the commit that shipped the insert — the bullet
  above it, about network calls, was updated in that same commit and this one
  was not.

- **The sibling ledger described a deck read that does not exist.** Four rows
  of `docs/SIBLING.md` and its source table said the deck read pages
  `getItemAt` at twenty and never does a collection load. It does three
  collection loads and no paging. Every row has been re-triaged against the
  code that shipped, and the four now say what is actually there and what
  guards it. A row that reads as a description of this add-in and is not is
  worse than no row, because the next reader builds on it.

- **Numbers in the prose that no longer matched the repo**: the pane draws 84
  tiles, not 85 (the runs cover 45 of the 117 elements); six of the 117 16:9
  elements arrive already tagged, carrying 78 tag relationships between them,
  where the text read as though 78 elements did; the catalogue is about 190 KB,
  not 110; the built library is about 16 MB, which three files said and the
  design record put at 14; the 4:3 library has eleven categories to 16:9's
  twelve. Each was measured against the committed catalogue or the built
  bundle, and two of them against the pane's own rendering.

- **The dependency triage said an alert on `jszip` or `@xmldom/xmldom` was
  not reachable from shipped code.** Both are imported from `src/` now, both
  are in the bundle, and both are what parses a .pptx a user can be sent. The
  page carries a new row saying so, and the old reading is marked superseded
  rather than edited away.

- **Undo's two slide numbers are told apart by name.** One counts from one
  because it is read aloud, the other from zero because it is fed to the API,
  and they were both called `slide` in the same file — which is the confusion
  the undo defect was. A dead `UNDO_DEPTH` of ten went with them: nothing read
  it, and the design has said one deep since the build.

- **The README no longer calls the add-in a scaffold that inserts nothing.**
  The splice, the picker and the host handshake all shipped, and the feature
  table said "planned" for all three; the status line said the pane inserts
  nothing yet. Only the manual is held to the code by CI, so the README drifted
  exactly where the lockstep rule says it would.

- **`SECURITY.md`'s "it makes no network calls" is now "it sends nothing
  anywhere"**, which is both true and stronger. The pane fetches its own
  catalogue from its own origin, as `docs/DESIGN.md` sections 3 and 11 always
  said it would; the test that used to ban `fetch` outright now holds the
  property that matters — one named file may fetch, it may not name an absolute
  address, and it may not pass a method, a body or a header, so there is no way
  to send anything to anyone.

### Fixed

- **And it could freeze on the right one.** The read that keeps that line fresh
  had the same twenty-second budget as a read somebody is waiting on, and a
  selection read fired straight after an insert can sit unanswered for most of
  it while PowerPoint finishes writing the deck. The pane answers one selection
  change at a time, so that one read froze the line: three clicks on three
  different slides went by with the pane still naming the first, for as long as
  anyone cared to watch. A glance gives up after four seconds now, and a glance
  that gave up leaves the line alone — the last thing the host actually said is
  a better answer to "I could not find out" than "PowerPoint did not say".

- **The pane could settle on the wrong slide number and stay there.** The line
  under the header follows your selection now, but the first version of that
  dropped any change that arrived while it was still reading the last one — and
  on the web that read is slow enough for an ordinary second click to fall
  inside it. Clicking slides 2, 3 and 4 a quarter of a second apart left the
  pane saying "Slide 3." for as long as anyone cared to watch, with no later
  event coming to put it right. It now remembers that something changed and
  asks once more when the read in flight finishes, so a burst of ten costs two
  reads rather than ten or one.

- **An undo could report failure and leave the deck wrong.** PowerPoint on the
  web can still give the OLD slide count after an insert that has already
  happened — 2.8 seconds of it, measured on 2026-09-11 by polling the count
  through a real insert. Undo read the count once, immediately after putting
  the user's slide back, got the old number, concluded the insert had not
  landed and stopped there: the deck kept six slides where five belonged, with
  the restored slide and the rebuilt one both in it. It said "Undo did not
  work" while it was the reading that had failed, not the insert. Every count
  that decides something is now asked again on a backoff until the deck agrees,
  and the ordinary case still costs one read.

- **A deck read that came back short would have named the wrong slide.**
  PowerPoint on the web answers a collection load of more than about fifty
  items with fewer than it has, and this add-in turns that list into a POSITION
  — which slide you are on, which slide an undo aims at — and then removes a
  slide by position. A list missing a slide in the middle makes every position
  after it name a different slide. `docs/SIBLING.md` had triaged this in
  September and promised the defence a sibling uses; the defence was never
  built, and the code shipped doing exactly the read the ledger said it never
  would. It now asks the deck for its slide COUNT in the same breath as the
  list, and hands out no position at all when the two disagree — the pane then
  says it does not know which slide you are on, which is a sentence a user can
  act on. No deck of ours is big enough to have shown this, which is the point.

- **The pane named the slide you were on when it opened, rather than the one
  you are on now.** It read the selection once, at startup, and never again, so
  clicking through the deck left the line under the header naming the first
  slide for the rest of the session. Found in PowerPoint for the web on
  2026-09-11: the API reported slide 2 selected and the pane still read
  "Slide 1.". The insert has always read the selection again for itself, so
  nothing ever landed in the wrong place — but the pane was telling you
  something untrue about where it was about to land.

- **Clicking a tile did nothing.** Most of a tile is the ghost drawing, an
  `<svg>`, and an SVG element is not an `HTMLElement` — which is all the click
  handler would walk up from. Found by opening the add-in in PowerPoint and
  pressing a tile. Every gate passed over it: jsdom was clicking the button,
  which is the part of a tile a user is least likely to hit, and the shot audit
  photographs states without pressing anything.
- **Undo reported success and changed nothing.** Taking back an insert that
  landed onto slide N means putting the user's original slide N back and
  removing the REBUILT one at index N. The code removed N+1, which is the copy
  it had just restored, so the deck came back to its old size, the count check
  passed, and the pane said "Undone" over a slide that had not moved. The
  arithmetic is a pure function now, with the off-by-one as a test, and the
  restoring insert aims at the rebuilt slide rather than at whatever the user
  happens to have selected by the time they press it. Confirmed in PowerPoint
  for the web on 2026-09-11 by reading the slide's shapes before and after: a
  triangle inserted onto a slide holding a title and a white box, then taken
  back, left the slide with the same id and the same two shapes it started
  with.

- **A new slide carried the previous slide's comments.** Found by running the
  real engine against PowerPoint for the web on 2026-09-10 and then reading the
  deck back: one comment came out on two slides. A modern comment is anchored
  from the slide's own extension list, so cloning a slide keeps it — right when
  the slide is being rebuilt, wrong when it is meant to be new. "As a new
  slide" now drops comments the way it already dropped speaker notes, and the
  same round afterwards showed the new slide with none.
- The harvest's `parts` list was documented as "every package part reachable
  from those relationships", and it is not: parts are collected once per deck,
  so the second element to use a picture lists nothing for it. Measured on the
  committed library, 64 relationship targets across the two sizes are absent
  from their own element's list. The splice resolves what to copy from the
  element's relationships instead, and the type says so.


### Added — the host probe

- A way to ask a real PowerPoint the questions the design rests on, before
  anything is built on a guess: paste `probe/probe-snippet.ts` into Script Lab,
  press Run, and copy the answer sheet back. It asks whether a package pruned
  to one slide is accepted, whether inserting a slide and then removing the one
  it replaced keeps the order, which slide the host says you are on, which of
  the two ways of reading the deck drops your comments, whether PowerPoint's
  own Ctrl+Z takes an insert back, and how long a read takes on a big deck.
  `docs/PROBE.md` has the steps; `scripts/read-answers.mjs` says what each
  answer means and files the sheet. Nothing in the pane changes.
- The first round, on PowerPoint for the web (2026-09-10): two pairs of
  sheets under `docs/host-answers/`, the second pair taken with the marker
  below. Every insert landed, both prunings land as one slide, the export
  drops comments and the authors part, and PowerPoint's own Ctrl+Z reverts an
  insert. `docs/PROBE.md` carries the summary.

### Fixed — the host probe

- The second run of a pair left a slide of its own. The only thing that told a
  second run from a first was the slide the first run leaves for Ctrl+Z, which
  is exactly what a successful Ctrl+Z removes, so on the web the second run
  took itself for a first. The first run now writes a marker into the
  document's settings, outside the undo stack, before it leaves the slide; the
  second run reads it, clears it, and leaves nothing. A lone second sheet also
  answers question 5 now, from the marker, and the reader says which source
  it read.
- The fixture-timestamp test compared local time components against an
  instant JSZip reads back in UTC, so the suite was red on any machine east
  of Greenwich and green on CI.

### Added — the library harvest

- The library exists as data: the two decks the owner authored (one per slide
  size, 117 elements each, in twelve categories for 16:9 and eleven for 4:3
  (the 4:3 deck has no white boxes with black headings), twenty-one of them
  stamps,
  markers, flowchart shapes and icons) are read into a catalogue with every
  element's English name, where it sits, where it lands and which pictures,
  charts and tags it carries. Nothing in the pane shows it yet; the picker is
  the next change but one. Editing the library is editing the deck: the rules
  are in the manual under "Adding an element to the library", and a change
  that breaks them is refused with every problem listed.

### Added — the design record

- `docs/DESIGN.md` says what the pane and the insert will do, decision by
  decision and dated: what counts as an element, how the library is authored,
  where each kind of element lands, what the footer reports, and what is still
  a question for a real PowerPoint. `template/names.en.json` carries the
  English name of every element in the library. Nothing in the pane changes
  yet.

### Added — the package layer

- The engine can now open a .pptx, read and change its parts, relationships,
  content types and slide list, and hand it back as bytes or base64. Nothing
  in the pane uses it yet; it is the ground the element library and the insert
  are built on, ported from SSF-Merge together with the tests that found its
  bugs there.

### Added — the sibling watch

- A weekly sweep of what SSF-Charts and SSF-Merge have learned about the
  PowerPoint host, filing one issue for any finding this repo has not answered.
  The answers so far live in `docs/SIBLING.md`: which of the siblings' findings
  shape the design, and which do not touch it at all.

### Added — the scaffold

The repository, before any of the product: everything that has to be true for
the first feature to land safely.

- Hosting on GitHub Pages at `ssf-slide-elements.struktureretsundfornuft.dk` (the
  domain's DNS record is the owner's step), with the deploy waiting for the same
  gate CI runs.
- CI: format, lint, types, the library build, the pane build, coverage with
  floors, and a floor under the number of tests that CI refuses to leave
  unrecorded.
- Manifests in both formats, dev and prod, generated from one source, checked
  offline by the suite and by Microsoft's validator in CI. The GUID is minted
  and pinned; the manifest version is `1.0.0.0` and moves only when the manifest
  does.
- Ribbon icons drawn by a script and byte-pinned by a test.
- A placeholder pane: one step, **Start here**, showing which build it is,
  checking that the host clears PowerPointApi 1.2, and saying plainly that there
  is nothing to insert yet. The **Insert an element** button is there and
  disabled.
- The documentation set — manual, backlog, this changelog, the security policy,
  the dependency-alert log — and the tests that keep the manual in step with the
  pane.
- A manual release workflow that validates what it ships and refuses to ship a
  manifest pointing at a host that is down; a weekly pane audit; Dependabot with
  a written log; a contributing guide, a PR template, code owners and the MIT
  licence.
