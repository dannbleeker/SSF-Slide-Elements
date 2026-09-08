/**
 * The engine's public surface.
 *
 * Nothing here imports Office.js, and `test/architecture.test.ts` holds that:
 * the engine takes bytes and answers bytes, which is what lets it run in the
 * pane, in a script and in the suite with no PowerPoint anywhere.
 *
 * Empty on purpose. The package layer, the harvest and the splice each arrive
 * in their own change (`docs/BACKLOG.md`), and a placeholder that pretended to
 * be one of them would be the thing the next reader builds on.
 */
export {};
