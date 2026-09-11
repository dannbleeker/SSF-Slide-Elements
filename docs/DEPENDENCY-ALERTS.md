# Dependency alerts

Every alert gets read and the reading gets written down here, **including "no
exposure"** — that entry is what records that somebody looked.

A standing red count with nobody looking is worse than no scanner at all: it
trains everyone to scroll past the one that matters. If this file is empty and
the alert count is not zero, the process has stopped working.

## How to triage one

1. Is the vulnerable code path reachable from this package? Most alerts on a
   build-time dependency are not.
2. Is there a fix that does not move a major version?
3. **Never take `npm audit`'s advice unread.** On a sibling project its
   proposed remedy for one alert was to move a runtime dependency back three
   major versions.
4. Record the decision below with the date, whatever it was.

## Log

| Date | Alert | Reading | Action |
| --- | --- | --- | --- |
| 2026-09-11 | **Dependabot SECURITY ALERTS are switched off for this repository** | Found while reading the version-update PR below: `GET /repos/:owner/:repo/vulnerability-alerts` answers **404**, and the repository's own `security_and_analysis` reports `dependabot_security_updates: disabled`. So the count this page is built around has been zero because **nobody is scanning**, not because nothing was found — and the page's own warning ("if this file is empty and the alert count is not zero, the process has stopped working") cannot catch that, because the count is not merely zero, it does not exist. The version updates that produced PR #3 are unaffected: those come from `.github/dependabot.yml` and do not need alerts to be on. Nothing here says there IS an unread vulnerability; it says this page has been reporting on a scanner that was never running. | **The owner's to switch on** — Settings → Advanced Security → Dependabot alerts, and Dependabot security updates beside it. A token with `admin:repo_hook` is needed to do it through the API, which this one does not have. Re-read the two runtime dependencies the moment it is on |
| 2026-09-11 | version updates, PR [#3](https://github.com/dannbleeker/SSF-Slide-Elements/pull/3): seven dev-tooling bumps, one of them a MAJOR | All seven are development dependencies — nothing here reaches the shipped bundle, and `dist/assets/` names only `jszip` and `@xmldom/xmldom`, which are untouched. The one to read rather than wave through is **vitest 4 → 5** with `@vitest/coverage-v8` beside it, because the coverage floors are a GATE and a major that quietly stopped enforcing them would look exactly like a green build. It still enforces: with `statements` raised to 99.9 the run exits 1, and at 95 it exits 0. Everything else agrees too — 932 tests, the same count; every generator (harvest, probe, manifests, icons, previews, the catalogue page) byte-identical; the pane audit's 88 shots green; the site builds. One number moved: v8 now reports **functions at 100%** where it reported 99.7%, which is a counting change in the reporter rather than a test. Playwright 1.62 → 1.63 needs its browser re-downloaded on a developer machine (`npx playwright install chromium`); CI installs it fresh every run, so no workflow changes. | Taken, on this branch rather than by merging #3 — merging somebody else's PR is outside what the owner authorised. #3 closes itself once these versions are on `main` |
| 2026-09-11 | none open; re-reading the 2026-09-08 baseline | **The 2026-09-08 reading no longer holds, and it was the only one on this page.** The package layer landed and the splice with it, so both runtime dependencies are now imported from `src/` — `jszip` by `src/core/pptx/pkg.ts`, `@xmldom/xmldom` by `src/core/pptx/xml.ts` and `src/core/catalogue/harvest.ts` — and both are in the shipped bundle, which `dist/assets/` confirms by name. They are also what parses a .pptx a user can be SENT, which is the untrusted input `SECURITY.md` is about: an alert on either is now reachable from shipped code and from hostile bytes, and is to be read that way rather than as a build-time concern. What has not changed is that neither is reached over a network and neither writes a file. The toolchain majors stay held in `dependabot.yml` for the reasons written there. | Alerts on `jszip` and `@xmldom/xmldom` are reachable; triage each on arrival rather than deferring |
| 2026-09-08 | none open at the scaffold | Baseline entry. Two runtime dependencies, `jszip` and `@xmldom/xmldom`, declared ahead of the engine that will use them to read package bytes a user supplies. Nothing under `src/` imports either yet — `@xmldom/xmldom` is used only by `test/manifest.test.ts` — and the built pane does not bundle them, so an alert on either is not reachable from shipped code until the package layer lands; re-read this entry when it does. **Superseded on 2026-09-11 by the row below; the package layer landed and this reading no longer holds.** The toolchain is pinned to SSF-Merge's current majors; `@types/node` and `typescript` majors are held in `dependabot.yml` for the reasons written there. | Watching |
