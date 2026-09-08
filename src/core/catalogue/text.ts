/**
 * Text read out of DrawingML, and the names built from it.
 */
import { A_NS, P_NS, elements, element } from "../pptx/xml.js";

/** Every run of text under an element, joined, whitespace collapsed. */
export function textOf(el: Element | Document): string {
  return elements(el, A_NS, "t")
    .map((t) => t.textContent ?? "")
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

/** Each paragraph's text under an element, empty ones dropped. */
export function paragraphsOf(el: Element): string[] {
  return elements(el, A_NS, "p")
    .map((p) => textOf(p))
    .filter((t) => t.length > 0);
}

/** The placeholder type of a shape, or undefined when it is not a placeholder. `body` is the type an untyped `<p:ph>` has. */
export function placeholderType(shape: Element): string | undefined {
  const nvPr = element(shape, P_NS, "nvPr");
  const ph = nvPr ? element(nvPr, P_NS, "ph") : undefined;
  if (!ph) return undefined;
  return ph.getAttribute("type") ?? "body";
}

/** The slide's title: the text of its title placeholder, or "" when it has none. */
export function titleOf(slide: Document): string {
  for (const sp of elements(slide, P_NS, "sp")) {
    const type = placeholderType(sp);
    if (type === "title" || type === "ctrTitle") return textOf(sp);
  }
  return "";
}

/** The `name` PowerPoint gave a shape in its selection pane. */
export function shapeName(shape: Element): string {
  return element(shape, P_NS, "cNvPr")?.getAttribute("name")?.trim() ?? "";
}

/** A name PowerPoint made up (`Gruppe 12`, `Rectangle 3`), which says nothing about the shape. */
const GENERIC_NAME =
  /^(Gruppe|Group|Rektangel|Rectangle|Billede|Picture|Ink|Tekstfelt|TextBox|Text Box|Buet forbindelse|Forbindelse|Connector|Oval|Ellipse|Frihåndstegning|Freeform|Pil|Arrow|Grafik|Graphic)\s*\d*$/i;

/** Text that is a placeholder for text rather than a name. */
const FILLER = /^(tekst|text|lorem\b.*|xx+|overskrift|note:?)$/i;

/**
 * A part's name: its own text with brackets stripped and cut at a colon, else
 * its shape name when PowerPoint did not make that up, else the slide's title
 * numbered. Duplicates on one slide get ` (2)`, ` (3)`.
 */
export function partName(shape: Element, slideTitle: string, used: Set<string>, nameless: { count: number }): string {
  let name: string | undefined;
  for (const para of elements(shape, A_NS, "p")) {
    let t = textOf(para).replace(/[[\]]/g, " ");
    if (t.includes(":")) t = t.slice(0, t.indexOf(":"));
    t = t
      .replace(/\s+/g, " ")
      .replace(/^[\s–-]+|[\s–-]+$/g, "")
      .trim();
    if (t && !FILLER.test(t)) {
      name =
        t.length <= 40
          ? t
          : t
              .slice(0, 40)
              .replace(/\s\S*$/, "")
              .replace(/[\s,;–-]+$/, "") + "…";
      break;
    }
  }
  if (!name) {
    const own = shapeName(shape);
    if (own && !GENERIC_NAME.test(own)) name = own;
  }
  if (!name) {
    nameless.count += 1;
    name = `${slideTitle} ${nameless.count}`;
  }
  const base = name;
  for (let n = 2; used.has(name); n++) name = `${base} (${n})`;
  used.add(name);
  return name;
}

/** A stable id from a key: lower case, Danish letters spelled out, everything else a hyphen. */
export function slug(key: string): string {
  return key
    .toLowerCase()
    .replace(/æ/g, "ae")
    .replace(/ø/g, "oe")
    .replace(/å/g, "aa")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}
