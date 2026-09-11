# Contributing

## Getting set up

```bash
npm install
npm test
```

Node 22 (`.nvmrc`).

`npm run dev` serves the pane over plain HTTP, which is what you want for
looking at it in a browser. Sideloading it into PowerPoint needs HTTPS on
`localhost:3002`, because that is the URL the dev manifest names — so install a
development certificate once and point Vite at it:

```bash
npx office-addin-dev-certs install
SSF_DEV_KEY=<key> SSF_DEV_CERT=<cert> npm run dev
```

The first command generates and trusts the pair and prints where it put them;
put those two paths into the two variables (in PowerShell,
`$env:SSF_DEV_KEY = "<key>"; $env:SSF_DEV_CERT = "<cert>"; npm run dev`).
Vite 8 has no `--https` flag, so `vite.config.ts` reads them from the
environment.

## The checks

| Command | What it holds |
| --- | --- |
| `npm test` | The suite |
| `npm run typecheck` | Types, with `noUncheckedIndexedAccess` on |
| `npm run lint` | Type-aware ESLint, including `no-floating-promises` |
| `npm run format:check` | Prettier, on code only. Prose is not reformatted |
| `npm run coverage` | Coverage floors on `src/core`, `src/host` and `src/pane` |
| `npm run test:count` | A floor under the number of tests, and a cap on how many may be skipped |
| `npm run build:lib` | The library build. `tsc --noEmit` cannot see it fail |
| `npm run build` | The pane bundle — what a user actually loads. Nothing else compiles it |
| `npm run probe` | Regenerates `probe/probe-snippet.ts` from the engine; CI diffs the committed file. Needs `build:lib` first |

CI runs all of them, `npm run coverage` standing in for `npm test`, and the
Pages deploy waits for the same list. `npm run lint -- --fix` and
`npm run format` fix most of what the check commands find.

Run the WHOLE gate as the last thing before a commit, on the code that is
committed — not before the last edit.

## Rules worth knowing before your first change

**`src/core` imports nothing from Office.js, and `src/office` decides nothing.**
A test holds both directions. That seam is what lets the engine be tested
without a PowerPoint, and PowerPoint on the web is documented, in `CLAUDE.md`,
to lie about object ids.

**A regression test must be proven to fail without its fix.** Break the source,
re-run, confirm the new test goes red *and that the right assertion went red*,
then restore. A guard that passes against the unfixed code is decoration. Say in
the PR which assertion you saw fail.

**Documentation lands in the same change as the feature.** `test/docs.test.ts`
reads the pane's steps and labels out of the source and fails when the manual
has not caught up. `CHANGELOG.md`, `docs/BACKLOG.md` and the README's feature
table are on you.

**Backlog items are removed when they ship, not ticked.** Anything still listed
is genuinely open. The manual may not mark anything *planned* while the backlog
has nothing open — a guard reads both.

**A manifest change costs the owner a re-sideload.** Say so in the PR and the
changelog, and bump `VERSION` in `scripts/manifest-source.mjs` — never the npm
version.

**Where the prose points has to be there.** `test/references.test.ts` checks
every relative markdown link, every `#anchor`, and every "`docs/DESIGN.md`
section N" in the repository — 130 of those, most of them in code comments.
Renumbering a heading is the change that breaks dozens of sentences at once.

**Nothing in `src/` is written for its test alone.** `test/dead-exports.test.ts`
sweeps for exports that no other file in `src/` or `scripts/` reaches, because
an export the product never calls is a comment that compiles — it type-checks,
its own test is green, and the thing it was written to do is still not done. If
one is deliberate, record it in `ALLOWED` in `scripts/dead-exports.mjs` with the
reason; `npm run dead-exports` prints the list.

**Every npm script stays flat** — no script that starts with `npm run`.

**Test files are named by topic, never by increment.** No `batch-3.test.ts`.

**Sample data is invented.** The repository is public.

## Looking at the pane

`npm run pane-shots` renders the pane at 320 and 512 in both themes and
measures it — overflow, contrast, focus rings, hit areas, axe — because jsdom
has no layout and no colour, so a rule about how the pane LOOKS is invisible to
the suite. Looking at the output is part of done for any change to the pane.

```bash
npx vite --port 5199 --strictPort &
npm run pane-shots          # PNGs in /tmp/pane-shots, plus the audit
```

## Host behaviour

Before writing anything that talks to PowerPoint, read the host rules in
`CLAUDE.md` and the ledger in `docs/SIBLING.md`. The rules are recordings from
the sibling projects' rounds against a real PowerPoint, not opinions, and
several of them are the reason the architecture is what it is; the ledger says
which finding each one came from, what was done about it here, and why a
borrowed counter carries its date. Reading either by hand needs both sibling
repositories checked out in the same session.
