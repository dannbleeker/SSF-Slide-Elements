/**
 * What a deck's `<a:schemeClr>` names actually resolve to.
 *
 * A colour in a slide is written one of two ways. `<a:srgbClr val="C83A3A"/>`
 * is that red and stays that red wherever the shape goes. `<a:schemeClr
 * val="accent1"/>` is a NAME, resolved at draw time against whatever theme the
 * deck the shape is sitting in happens to use — which is why an element pasted
 * into a customer's deck comes out in the customer's colours, and why
 * `docs/DESIGN.md` section 7 can offer "This deck's theme" for nothing at all.
 *
 * The other position on that switch, "As in the library", needs those names
 * resolved ONCE, against the library's own theme, and written into the markup
 * as explicit values. This is the reader that resolves them.
 *
 * Two hops, and the second is the one a reader forgets. The theme's
 * `<a:clrScheme>` is keyed by SLOT — `dk1`, `lt1`, `dk2`, `lt2`, `accent1`…6,
 * `hlink`, `folHlink` — while a slide writes `bg1`, `tx1`, `bg2` and `tx2`,
 * which are not slots at all. The master's `<p:clrMap>` says which slot each of
 * those four means, and it is routinely a SWAP: `bg1="lt1" tx1="dk1"` on a
 * light design, the other way round on a dark one. Resolving `tx1` straight to
 * the theme's `tx1` finds nothing; assuming `tx1` is `dk1` is right until the
 * first dark master, and then every text colour in the library comes out white
 * on white.
 *
 * Pure, like everything in `src/core`: a `Pkg` in, a map of name to hex out.
 */
import { Pkg } from "./pkg.js";
import { layoutOf, masterOf } from "./layout.js";
import { REL_TYPE } from "./parts.js";
import { A_NS, P_NS, PKG_REL_NS, child, elements } from "./xml.js";

/**
 * A resolved colour map: every `<a:schemeClr val="…">` this deck can write,
 * against the six-digit RGB it resolves to. Upper case, no `#`, which is how
 * the format itself spells a colour.
 *
 * Both spellings are in it — the theme's own slots and the four names the
 * colour map redirects — because both are legal in a slide and the library
 * uses five of them.
 */
export type ThemeColours = Record<string, string>;

/** The theme slots, in the order `<a:clrScheme>` declares them. */
const SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

/**
 * The colour a slot's element states.
 *
 * `<a:srgbClr val="44546A"/>` is the simple case. `<a:sysClr
 * val="windowText" lastClr="000000"/>` is what PowerPoint writes for `dk1` and
 * `lt1` on the stock Office theme, and the measured decks use it: the 16:9
 * library's `dk1` is `windowText`. `lastClr` is the value the authoring
 * application last resolved that system colour to, and it is the only concrete
 * answer in the file — a pin that ignored it would drop black and white, the
 * two commonest colours in the library, out of the map entirely.
 */
function colourOf(slot: Element): string | undefined {
  const srgb = child(slot, A_NS, "srgbClr");
  const value = srgb?.getAttribute("val");
  if (value && /^[0-9A-Fa-f]{6}$/.test(value)) return value.toUpperCase();
  const sys = child(slot, A_NS, "sysClr");
  const last = sys?.getAttribute("lastClr");
  if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return last.toUpperCase();
  return undefined;
}

/** The one part of a relationship type that a part points at, if any. */
async function relatedOfType(pkg: Pkg, owner: string, type: string): Promise<string | undefined> {
  const relsPath = Pkg.relsPathFor(owner);
  if (!pkg.has(relsPath)) return undefined;
  const rels = await pkg.doc(relsPath);
  for (const rel of elements(rels, PKG_REL_NS, "Relationship")) {
    if (rel.getAttribute("Type") !== type) continue;
    if ((rel.getAttribute("TargetMode") ?? "") === "External") continue;
    const target = rel.getAttribute("Target");
    if (!target) continue;
    const path = pkg.resolved(owner, target);
    if (pkg.has(path)) return path;
  }
  return undefined;
}

/** The theme a master points at. */
export function themeOf(pkg: Pkg, masterPath: string): Promise<string | undefined> {
  return relatedOfType(pkg, masterPath, REL_TYPE.theme);
}

/** Slide, layout, master, theme: the chain a scheme colour is resolved down. */
export async function themeChain(
  pkg: Pkg,
  slidePath: string,
): Promise<{ layout?: string; master?: string; theme?: string }> {
  const layout = await layoutOf(pkg, slidePath);
  const master = layout ? await masterOf(pkg, layout) : undefined;
  const theme = master ? await themeOf(pkg, master) : undefined;
  return { ...(layout ? { layout } : {}), ...(master ? { master } : {}), ...(theme ? { theme } : {}) };
}

/**
 * Read one theme part and its master's colour map into resolved colours.
 *
 * A master is optional: without one the theme's own slots still resolve, and
 * the four mapped names are left out rather than guessed. An identity map is a
 * guess that is right most of the time, and the times it is wrong are the ones
 * that matter.
 */
export async function coloursOf(pkg: Pkg, themePath: string, masterPath?: string): Promise<ThemeColours> {
  const doc = await pkg.doc(themePath);
  const scheme = elements(doc, A_NS, "clrScheme")[0];
  const out: ThemeColours = {};
  if (!scheme) return out;
  for (const slot of SLOTS) {
    const el = child(scheme, A_NS, slot);
    const colour = el ? colourOf(el) : undefined;
    if (colour) out[slot] = colour;
  }
  if (!masterPath || !pkg.has(masterPath)) return out;
  const master = await pkg.doc(masterPath);
  const map = child(master.documentElement, P_NS, "clrMap");
  if (!map) return out;
  for (let i = 0; i < map.attributes.length; i++) {
    const attr = map.attributes.item(i);
    if (!attr) continue;
    const slot = out[attr.value];
    if (slot !== undefined) out[attr.name] = slot;
  }
  return out;
}

/** The colours a slide's own scheme names resolve to, walking the chain for it. */
export async function themeColoursFor(pkg: Pkg, slidePath: string): Promise<ThemeColours> {
  const chain = await themeChain(pkg, slidePath);
  if (!chain.theme) return {};
  return coloursOf(pkg, chain.theme, chain.master);
}
