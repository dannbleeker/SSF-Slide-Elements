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
