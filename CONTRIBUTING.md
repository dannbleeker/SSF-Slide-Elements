# Contributing

## Adding or changing an element

The library deck in `template/` is the authoring surface. Open it in PowerPoint,
draw the element on a slide using the `Kun titel` layout, give the slide a title
— that is the name the picker shows — and save.

Then regenerate the catalogue and commit what it wrote:

```bash
npm run build:lib
node scripts/harvest.mjs
```

Read its output. It reports every slide it skipped and why, and CI fails a pull
request where the deck and the committed catalogue disagree.

## Before opening a pull request

```bash
npm run format:check
npm run lint
npm run typecheck
npm test
```

If you changed the pane, look at it:

```bash
npm run build
npx vite preview --port 4178 --strictPort &
npm run pane-shots
```

## Developing the pane

```bash
npm run dev            # localhost:3000
```

The pane renders its picker without a host, so it can be inspected in an
ordinary browser. Only the INSERT needs PowerPoint — clicking a card without one
fails at the point it reads the deck and says so.

Sideloading needs HTTPS: `npx office-addin-dev-certs install`, then run Vite with
`--https.key` and `--https.cert`. `manifest.xml` points at `localhost:3000`.

## House rules

- Every npm script stays FLAT — no script that starts with `npm run`.
- `src/core` imports nothing from Office.js; `src/office` decides nothing.
- A new test must be proven to fail without its fix.
- Say what is measured and what is assumed. `docs/UNPROVEN.md` is the ledger.
