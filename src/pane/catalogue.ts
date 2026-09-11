/**
 * The library, fetched from the add-in's own site.
 *
 * **This is the only file in the add-in that touches the network, and it can
 * only ever GET a static file from the origin the pane itself was served
 * from.** `test/security.test.ts` holds both halves of that: no other file may
 * name `fetch`, and nothing here may carry an absolute URL, a request method or
 * a body. `SECURITY.md` says the same thing in prose. The add-in has no
 * backend, sends nothing anywhere, and never puts anything from the user's deck
 * into a request.
 *
 * Why it fetches at all: `docs/DESIGN.md` section 3 makes the catalogue static
 * files on Pages rather than a bundle, because the element markup and the parts
 * it carries come to about 16 MB across both sizes and the pane has to be
 * usable in about two seconds (section 11). The index is small and comes first;
 * an element's markup and the pictures it carries are fetched when somebody
 * actually inserts it.
 *
 * Everything fetched is cached for the session. A user who inserts the same
 * element twice pays for it once, and a part shared by two elements is fetched
 * once however many elements name it.
 */
import type { Catalogue, Element, Markup } from "../core/catalogue/types.js";
import type { Library } from "./steps.js";

/**
 * Where the catalogue sits, relative to the pane.
 *
 * Relative, and it must stay relative: an absolute URL here would be a request
 * to somewhere other than the add-in's own origin, which is the one thing the
 * security page promises never happens.
 */
const ROOT = "./catalogue";

/** The whole committed index: both sizes, and the version that stamps them. */
export interface Index {
  version: string;
  sizes: Record<string, Catalogue>;
}

/** A part of the catalogue, as the splice wants it: text for XML, bytes for the rest. */
export type Part = string | Uint8Array;

/** A part is XML when the package holds it as text. The harvest uses the same rule. */
function isText(path: string): boolean {
  return /\.(xml|rels)$/i.test(path);
}

/** The directory a size's files live under: `16:9` is not a legal path segment. */
export function dirOf(size: string): string {
  return size.replace(":", "x");
}

async function getText(path: string): Promise<string> {
  const response = await fetch(path);
  if (!response.ok) throw new Error(`ssf-slide-elements: ${path} answered ${response.status}`);
  return response.text();
}

async function getBytes(path: string): Promise<Uint8Array> {
  const response = await fetch(path);
  if (!response.ok) throw new Error(`ssf-slide-elements: ${path} answered ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

/** The index, fetched once. */
export async function loadIndex(): Promise<Index> {
  return JSON.parse(await getText(`${ROOT}/catalogue.json`)) as Index;
}

/**
 * Which library a deck of this shape should use, and what to say about it.
 *
 * `docs/DESIGN.md` section 3: two decks, one per slide size, and "a deck that
 * is neither 16:9 nor 4:3 (A4, 16:10, a custom size) borrows the nearest
 * library, scaled to fit, and the line under the pane's header says so". The
 * nearest is decided by ASPECT RATIO, because that is what decides whether an
 * element's proportions survive being scaled.
 *
 * Never refuses. A deck of an unusual shape still gets a library, because an
 * element that lands slightly scaled is worth more than a pane that will not
 * open on somebody's A4 report.
 */
export function libraryFor(index: Index, width: number, height: number): Library {
  const ratio = height > 0 ? width / height : 16 / 9;
  const sizes = Object.entries(index.sizes);
  if (sizes.length === 0) throw new Error("ssf-slide-elements: the catalogue index carries no libraries");
  let best = sizes[0] as [string, Catalogue];
  for (const entry of sizes) {
    const [, catalogue] = entry;
    const theirs = catalogue.height > 0 ? catalogue.width / catalogue.height : 0;
    const bestRatio = best[1].height > 0 ? best[1].width / best[1].height : 0;
    if (Math.abs(theirs - ratio) < Math.abs(bestRatio - ratio)) best = entry;
  }
  const [size, catalogue] = best;
  const exact = Math.abs((catalogue.width / catalogue.height - ratio) / ratio) < 0.01;
  return {
    size: catalogue.size,
    width: catalogue.width,
    height: catalogue.height,
    version: index.version,
    categories: catalogue.categories,
    elements: catalogue.elements,
    ...(exact ? {} : { borrowed: `${size} library, scaled to ${nameOfRatio(ratio)} slides.` }),
  };
}

/**
 * What to call a slide shape in a sentence.
 *
 * The two the design names by hand, then the ratio itself. A number is not
 * pretty, and it is honest: a user with a custom size recognises their own
 * deck in "1.50:1" where "custom" tells them nothing.
 */
export function nameOfRatio(ratio: number): string {
  if (Math.abs(ratio - 1.414) < 0.02) return "A4";
  if (Math.abs(ratio - 1.6) < 0.02) return "16:10";
  if (Math.abs(ratio - 16 / 9) < 0.02) return "16:9";
  if (Math.abs(ratio - 4 / 3) < 0.02) return "4:3";
  return `${ratio.toFixed(2)}:1`;
}

/**
 * The catalogue's own record of what each carried part IS.
 *
 * The splice declares a content type for every part it copies into the user's
 * package, and a part with none is a package PowerPoint refuses outright.
 */
export function carriedTypes(index: Index, size: string): Record<string, string> {
  return index.sizes[size]?.carried ?? {};
}

/** One element's markup and everything it needs, fetched and cached. */
export class Store {
  private readonly elements = new Map<string, Promise<Markup>>();
  private readonly parts = new Map<string, Promise<Part>>();

  constructor(private readonly size: string) {}

  /** The markup of one element, by its catalogue id. */
  markup(element: Element): Promise<Markup> {
    const key = element.id;
    const held = this.elements.get(key);
    if (held) return held;
    const path = `${ROOT}/${dirOf(this.size)}/elements/${encodeURIComponent(key)}.json`;
    const work = getText(path).then((text) => JSON.parse(text) as Markup);
    this.elements.set(key, work);
    return work;
  }

  /**
   * One carried part, by the path the library knew it by.
   *
   * Answers undefined rather than throwing when the part is not there, because
   * that is what `PartStore` in the splice is documented to do — and the splice
   * turns it into a named refusal with the part's path in it, which is a better
   * sentence than a bare 404.
   */
  async part(path: string): Promise<Part | undefined> {
    const held = this.parts.get(path);
    if (held) return held;
    const url = `${ROOT}/${dirOf(this.size)}/parts/${path.split("/").map(encodeURIComponent).join("/")}`;
    const work = isText(path) ? getText(url) : getBytes(url);
    this.parts.set(path, work);
    try {
      return await work;
    } catch {
      // Forgotten, so a part missed because the network dropped is tried again
      // on the next insert rather than being remembered as absent forever.
      this.parts.delete(path);
      return undefined;
    }
  }
}
