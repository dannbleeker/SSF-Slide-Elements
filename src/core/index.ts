/**
 * The engine's public surface.
 *
 * Nothing here imports Office.js, and `test/architecture.test.ts` holds that:
 * the engine takes bytes and answers bytes, which is what lets it run in the
 * pane, in a script and in the suite with no PowerPoint anywhere.
 *
 * Today this is the package layer: a .pptx as parts, relationships, content
 * types and a slide list, with base64 in and out. The harvest and the splice
 * arrive in their own changes (`docs/BACKLOG.md`).
 */
export { Pkg, extensionOf, resolveTarget, resolveTargetSpellings } from "./pptx/pkg.js";
export { COMMENT_REL_TYPES, OWNABLE_BY_GRAPHIC, OWNED_BY_SLIDE, REL_TYPE } from "./pptx/parts.js";
export {
  A_NS,
  CT_NS,
  C_NS,
  CX_NS,
  MC_NS,
  PKG_REL_NS,
  P_NS,
  R_NS,
  SSML_NS,
  child,
  children,
  element,
  elements,
  parseXml,
  relationshipIdsIn,
  serializeXml,
  xmlSafe,
} from "./pptx/xml.js";
export { harvest, HarvestError } from "./catalogue/harvest.js";
export type { HarvestOptions } from "./catalogue/harvest.js";
export { boxOf, offSlide, rounded, topLevelShapes, union } from "./catalogue/boxes.js";
export { countKeys, countedNoun, sizeRuns } from "./catalogue/runs.js";
export { tagsFor } from "./catalogue/tags.js";
export { paragraphsOf, partName, placeholderType, shapeName, slug, textOf, titleOf } from "./catalogue/text.js";
export type {
  Box,
  Catalogue,
  Element,
  Harvest,
  Landing,
  Markup,
  MarkupRel,
  Names,
  SizeRun,
  SlideSize,
} from "./catalogue/types.js";
