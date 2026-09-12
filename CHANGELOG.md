# Changelog

Notable changes to SSF Slide Elements. Newest first.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses [semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
- **This one has not been run against a real PowerPoint yet.** The mechanism is
  the insert's, which has; a sequence of them has not.

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
