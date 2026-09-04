/**
 * XML for the package layer.
 *
 * One implementation everywhere, deliberately. `@xmldom/xmldom` is pure
 * JavaScript, so the task pane, the harvest script and the test suite all parse
 * and serialise with the same code. Reaching for the browser's native
 * `DOMParser` where it happens to exist would buy a little speed and a class of
 * bug this project cannot afford: an insert that works in the suite and
 * produces a file PowerPoint refuses to open, because two parsers disagreed
 * about a namespace declaration nobody was looking at.
 *
 * Ported from SSF-Merge, which learned most of what is written below the
 * expensive way. `docs/SIBLING.md` is the ledger of what came across.
 */
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

/** DrawingML: shapes, runs, paragraphs, fills, the transform on every element. */
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
/** PresentationML: slides, the shape tree, the slide id list. */
export const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
/** Relationship references *inside* a part (`r:embed="rId3"`). */
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/** The `.rels` parts themselves, which use a different namespace from `r:id`. */
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
/**
 * Markup Compatibility: `<mc:AlternateContent>` and its `Choice`/`Fallback`
 * branches, which is how a slide carries a feature older hosts cannot read. A
 * modern chart in a harvested element arrives wrapped in one.
 */
export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
/** `[Content_Types].xml`. */
export const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/**
 * A UTF-8 byte order mark at the very start of an XML part.
 *
 * Legal in an OPC package and emitted by default: .NET's `UTF8Encoding` writes
 * one unless explicitly told not to, so a deck from any third-party generator
 * built on it carries one on every part it wrote. JSZip's `async("string")`
 * hands the character straight through — it decodes UTF-8 and has no opinion
 * about what the first code point means — so the mark reaches the parser as
 * content.
 */
const BOM = "﻿";

/**
 * Parse a part, tolerating a leading byte order mark.
 *
 * `@xmldom/xmldom` refuses one outright: with a `U+FEFF` in front of the XML
 * declaration it reports that the declaration is not at the start of the
 * document and THROWS out of whatever was reading the part — on a deck
 * PowerPoint opens without a murmur. Stripped here rather than at each reader,
 * because this is the one door every part in the package comes through and a
 * second reader would be free to forget.
 */
export function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(
    xml.startsWith(BOM) ? xml.slice(BOM.length) : xml,
    "text/xml",
  ) as unknown as Document;
}

export function serializeXml(doc: Document | Element): string {
  return new XMLSerializer().serializeToString(doc as never);
}

/** Every descendant with this local name in the given namespace, in document order. */
export function elements(root: Document | Element, ns: string, local: string): Element[] {
  return Array.from(root.getElementsByTagNameNS(ns, local));
}

/** The first such descendant, or undefined. Never null, so callers can `??`. */
export function element(root: Document | Element, ns: string, local: string): Element | undefined {
  return elements(root, ns, local)[0];
}

/**
 * Direct children with this local name, never deeper.
 *
 * `elements` walks DESCENDANTS, which is right for "find every tag in this
 * part" and catastrophically wrong for "what does this element own". A slide's
 * `<p:spTree>` contains every nested shape, so asking it for descendant
 * `<p:sp>` returns each group's children alongside the groups themselves — and
 * an element harvested that way carries its own contents twice. Ask for
 * children when the parent is the point, which for element extraction it
 * always is.
 */
export function children(parent: Element, ns: string, local: string): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    const el = n as Element;
    if (el.nodeType === 1 && el.localName === local && el.namespaceURI === ns) out.push(el);
  }
  return out;
}

/** The first direct child with this local name, or undefined. */
export function child(parent: Element, ns: string, local: string): Element | undefined {
  return children(parent, ns, local)[0];
}

/** Every direct child element, whatever it is called. */
export function childElements(parent: Element): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) out.push(n as Element);
  }
  return out;
}

/**
 * Every relationship id this markup names, whatever names it.
 *
 * The question is "which relationships must travel with this element", and the
 * only safe way to ask it is of the WHOLE subtree: any attribute in the
 * relationship namespace is a reference, and a reference means the
 * relationship has to come along.
 *
 * This began life in SSF-Merge as a scan for `a:blip/@r:embed`, which is where
 * a picture's image sits and is not the only place an image id appears. Two
 * others are entirely ordinary:
 *
 * - `<asvg:svgBlip r:embed>`, inside the blip's own extension list. That is how
 *   PowerPoint stores an ICON — a raster fallback in the blip and the real SVG
 *   beside it, under a SECOND image relationship. Icons are everywhere in a
 *   modern deck, and this library has a whole section of them.
 * - `<a:blip r:link>`, a picture linked rather than embedded.
 *
 * Neither is an `r:embed` on an `a:blip`, so both were missed — producing an
 * element that names a relationship which is not there, which is precisely what
 * PowerPoint calls a damaged file. Reading every `r:`-namespaced attribute
 * rather than a list of the ones known today is the conservative direction: an
 * unknown reference carries a relationship that could have stayed behind, where
 * a missed one breaks the file.
 *
 * Matched by namespace OR by prefix: a document that came out of a host may not
 * carry the namespace where a reader expects it.
 */
export function relationshipIdsIn(root: Document | Element): Set<string> {
  const used = new Set<string>();
  const walk = (node: Element): void => {
    const attrs = node.attributes;
    for (let i = 0; i < (attrs?.length ?? 0); i++) {
      const attr = attrs?.item(i);
      if (!attr) continue;
      if (attr.namespaceURI === R_NS || attr.name.startsWith("r:")) used.add(attr.value);
    }
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) walk(c as Element);
    }
  };
  const start = (root as Document).documentElement ?? (root as Element);
  if (start) walk(start);
  return used;
}

/**
 * Characters no XML document may contain, in any spelling.
 *
 * XML 1.0's `Char` production excludes most of the C0 controls, the lone
 * surrogates and `FFFE`/`FFFF`, and there is no escape for them either —
 * `&#11;` is exactly as ill-formed as the raw character. `@xmldom/xmldom`
 * writes such a character straight through and reads it back again, so the part
 * serialises, the suite sees a perfectly good document, and every gate in this
 * repo is green. PowerPoint parses conformingly, refuses the part, and calls
 * the whole file damaged.
 *
 * Replaced with a SPACE rather than dropped, because the likeliest one is a
 * line break: dropping it joins two words that were separate, and a space
 * cannot make that mistake in the other direction.
 *
 * Astral characters are NOT touched. A well-formed surrogate pair is one code
 * point above `FFFF`, which `\p{Surrogate}` under the `u` flag does not match;
 * only an unpaired half does, and an unpaired half is already broken text.
 */
// Both rules are aimed at a control character reaching a pattern by accident.
// Here they ARE the subject: this is the set XML refuses, and it cannot be
// written without naming them.
// eslint-disable-next-line no-control-regex, no-irregular-whitespace
const XML_FORBIDDEN = /[ --￾￿]|\p{Surrogate}/gu;

/** Text a conforming XML parser will accept. See `XML_FORBIDDEN`. */
export function xmlSafe(text: string): string {
  return text.replace(XML_FORBIDDEN, " ");
}

/** Every unit of measure in this file is an EMU. A point is 12700 of them. */
export const EMU_PER_POINT = 12700;
