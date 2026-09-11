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
| 2026-09-11 | none open; re-reading the 2026-09-08 baseline | **The 2026-09-08 reading no longer holds, and it was the only one on this page.** The package layer landed and the splice with it, so both runtime dependencies are now imported from `src/` — `jszip` by `src/core/pptx/pkg.ts`, `@xmldom/xmldom` by `src/core/pptx/xml.ts` and `src/core/catalogue/harvest.ts` — and both are in the shipped bundle, which `dist/assets/` confirms by name. They are also what parses a .pptx a user can be SENT, which is the untrusted input `SECURITY.md` is about: an alert on either is now reachable from shipped code and from hostile bytes, and is to be read that way rather than as a build-time concern. What has not changed is that neither is reached over a network and neither writes a file. The toolchain majors stay held in `dependabot.yml` for the reasons written there. | Alerts on `jszip` and `@xmldom/xmldom` are reachable; triage each on arrival rather than deferring |
| 2026-09-08 | none open at the scaffold | Baseline entry. Two runtime dependencies, `jszip` and `@xmldom/xmldom`, declared ahead of the engine that will use them to read package bytes a user supplies. Nothing under `src/` imports either yet — `@xmldom/xmldom` is used only by `test/manifest.test.ts` — and the built pane does not bundle them, so an alert on either is not reachable from shipped code until the package layer lands; re-read this entry when it does. **Superseded on 2026-09-11 by the row below; the package layer landed and this reading no longer holds.** The toolchain is pinned to SSF-Merge's current majors; `@types/node` and `typescript` majors are held in `dependabot.yml` for the reasons written there. | Watching |
