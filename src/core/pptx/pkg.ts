/**
 * A .pptx as what it is: a zip of XML parts that reference each other.
 *
 * Both halves of this add-in live here rather than in Office.js. The HARVEST
 * reads the library deck and lifts an element's markup and parts out of it; the
 * INSERT opens the user's own deck, splices that markup into one slide, and
 * hands the result back to PowerPoint in a single call. Nothing is drawn shape
 * by shape, so none of the per-shape failure surfaces a sibling project
 * documents at length exist here.
 *
 * Ported from SSF-Merge and trimmed to what element insertion needs.
 */
import JSZip from "jszip";
import { CT_NS, PKG_REL_NS, P_NS, R_NS, element, elements, parseXml, serializeXml } from "./xml.js";
import { EXTERNAL } from "./parts.js";

const CONTENT_TYPES = "[Content_Types].xml";
const PRESENTATION = "ppt/presentation.xml";

/**
 * The next number for a part family, given the ones already in use.
 *
 * Highest plus one, never filling a gap. The whole of the safety is
 * `countable`: above 2^53, `max + 1 === max`, so a package holding
 * `image9007199254740992.png` would answer a number already in use and the copy
 * would overwrite it silently. Ignoring a digit run too large to count exactly
 * leaves the maximum exact, so `max + 1` is a real step past every member.
 */
function nextFree(used: Set<number>): number {
  let max = 0;
  for (const n of used) if (n > max) max = n;
  const next = max + 1;
  if (!Number.isSafeInteger(next))
    throw new Error("ssf-slide-elements: this package's part numbers are too large to extend");
  return next;
}

/** Every whole number a path matched, ignoring any too large to count exactly. */
function countable(n: number): boolean {
  return Number.isSafeInteger(n) && n > 0;
}

export class Pkg {
  private readonly docs = new Map<string, Document>();

  private constructor(private readonly zip: JSZip) {}

  static async open(source: string | Uint8Array | ArrayBuffer): Promise<Pkg> {
    const zip = await JSZip.loadAsync(source, typeof source === "string" ? { base64: true } : {});
    return new Pkg(zip);
  }

  /** The part as text, or a throw naming it. A missing part is a broken deck. */
  async text(path: string): Promise<string> {
    const file = this.zip.file(path);
    if (!file) throw new Error(`ssf-slide-elements: this presentation has no ${path}`);
    return file.async("string");
  }

  /** The part as text, or undefined. For parts that are legitimately optional. */
  async maybeText(path: string): Promise<string | undefined> {
    const file = this.zip.file(path);
    return file ? file.async("string") : undefined;
  }

  has(path: string): boolean {
    return this.zip.file(path) !== null;
  }

  /**
   * Replace a part's text.
   *
   * The parsed copy is dropped, because a caller that writes text and then asks
   * for the document would otherwise get the version from before the write —
   * silently, and with no way to tell from the result.
   */
  setText(path: string, xml: string): void {
    this.zip.file(path, xml);
    this.docs.delete(path);
  }

  setBytes(path: string, bytes: Uint8Array): void {
    this.zip.file(path, bytes);
    this.docs.delete(path);
  }

  async bytes(path: string): Promise<Uint8Array> {
    const file = this.zip.file(path);
    if (!file) throw new Error(`ssf-slide-elements: this presentation has no ${path}`);
    return file.async("uint8array");
  }

  /**
   * The part as a parsed document, cached.
   *
   * Callers MUTATE what they get back and then `write` it. Handing out a fresh
   * parse each time would silently discard every edit made through an earlier
   * handle to the same part.
   */
  async doc(path: string): Promise<Document> {
    const cached = this.docs.get(path);
    if (cached) return cached;
    const doc = parseXml(await this.text(path));
    this.docs.set(path, doc);
    return doc;
  }

  /** Serialise a document obtained from `doc` back into the package. */
  write(path: string, doc: Document): void {
    this.zip.file(path, serializeXml(doc));
    this.docs.set(path, doc);
  }

  /** Where a part's relationships live: `ppt/slides/_rels/slide1.xml.rels`. */
  static relsPathFor(part: string): string {
    const cut = part.lastIndexOf("/");
    return `${part.slice(0, cut)}/_rels/${part.slice(cut + 1)}.rels`;
  }

  /** Every `<Relationship>` on a part, as read. Empty when the part has none. */
  async rels(ownerPart: string): Promise<Element[]> {
    const path = Pkg.relsPathFor(ownerPart);
    const xml = await this.maybeText(path);
    if (xml === undefined) return [];
    return elements(parseXml(xml), PKG_REL_NS, "Relationship");
  }

  /**
   * Add a relationship to a part, creating its `.rels` if it has none, and
   * answer the id it was given.
   *
   * The id is free within the part, found the same way part numbers are.
   */
  async addRel(ownerPart: string, type: string, target: string, mode?: string): Promise<string> {
    const path = Pkg.relsPathFor(ownerPart);
    const existing = await this.maybeText(path);
    const doc = parseXml(
      existing ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"/>`,
    );
    const root = doc.documentElement;
    if (!root) throw new Error(`ssf-slide-elements: ${path} has no root element`);
    const used = new Set<number>();
    for (const rel of elements(doc, PKG_REL_NS, "Relationship")) {
      const m = /^rId(\d+)$/.exec(rel.getAttribute("Id") ?? "");
      if (m?.[1] !== undefined) {
        const n = Number(m[1]);
        if (countable(n)) used.add(n);
      }
    }
    const id = `rId${nextFree(used)}`;
    const rel = doc.createElementNS(PKG_REL_NS, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", type);
    rel.setAttribute("Target", target);
    if (mode) rel.setAttribute("TargetMode", mode);
    root.appendChild(rel);
    this.setText(path, serializeXml(doc));
    return id;
  }

  /**
   * What a relationship id on a part points at, as a package path — or, for an
   * external target, the raw value with a marker.
   */
  async relTarget(
    ownerPart: string,
    rId: string,
  ): Promise<{ path: string; external: boolean; type: string } | undefined> {
    for (const rel of await this.rels(ownerPart)) {
      if (rel.getAttribute("Id") !== rId) continue;
      const target = rel.getAttribute("Target") ?? "";
      const type = rel.getAttribute("Type") ?? "";
      if (rel.getAttribute("TargetMode") === EXTERNAL) return { path: target, external: true, type };
      return { path: resolveTarget(ownerPart, target), external: false, type };
    }
    return undefined;
  }

  /**
   * The next free number in a part family — `ppt/media/image` + `.png`, and so
   * on. Counts every extension for media, because `image7.png` and `image7.emf`
   * would otherwise both be answered 7.
   */
  nextNumber(prefix: string, suffix?: string): number {
    const used = new Set<number>();
    const pattern = new RegExp(
      `^${prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(\\d+)${suffix ? suffix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") : "\\.[A-Za-z0-9]+"}$`,
    );
    for (const name of Object.keys(this.zip.files)) {
      const m = pattern.exec(name);
      if (m?.[1] !== undefined) {
        const n = Number(m[1]);
        if (countable(n)) used.add(n);
      }
    }
    return nextFree(used);
  }

  /** Register a default content type for an extension, if it has none. */
  async addContentTypeDefault(extension: string, contentType: string): Promise<void> {
    const doc = await this.doc(CONTENT_TYPES);
    const ext = extension.toLowerCase();
    for (const d of elements(doc, CT_NS, "Default")) {
      if ((d.getAttribute("Extension") ?? "").toLowerCase() === ext) return;
    }
    const root = doc.documentElement;
    if (!root) throw new Error("ssf-slide-elements: [Content_Types].xml has no root element");
    const node = doc.createElementNS(CT_NS, "Default");
    node.setAttribute("Extension", ext);
    node.setAttribute("ContentType", contentType);
    // Defaults precede overrides in every package PowerPoint writes. The schema
    // does not require it and PowerPoint does not care, but a file that reads
    // like the ones around it is one less difference to rule out.
    root.insertBefore(node, root.firstChild);
    this.write(CONTENT_TYPES, doc);
  }

  /** Register an explicit content type for one part, replacing any it had. */
  async addContentTypeOverride(partName: string, contentType: string): Promise<void> {
    const doc = await this.doc(CONTENT_TYPES);
    const name = partName.startsWith("/") ? partName : `/${partName}`;
    for (const o of elements(doc, CT_NS, "Override")) {
      if (o.getAttribute("PartName") === name) {
        o.setAttribute("ContentType", contentType);
        this.write(CONTENT_TYPES, doc);
        return;
      }
    }
    const root = doc.documentElement;
    if (!root) throw new Error("ssf-slide-elements: [Content_Types].xml has no root element");
    const node = doc.createElementNS(CT_NS, "Override");
    node.setAttribute("PartName", name);
    node.setAttribute("ContentType", contentType);
    root.appendChild(node);
    this.write(CONTENT_TYPES, doc);
  }

  /** What content type the package declares for a part, override or default. */
  async contentTypeOf(part: string): Promise<string | undefined> {
    const doc = await this.doc(CONTENT_TYPES);
    const name = part.startsWith("/") ? part : `/${part}`;
    for (const o of elements(doc, CT_NS, "Override")) {
      if (o.getAttribute("PartName") === name) return o.getAttribute("ContentType") ?? undefined;
    }
    const ext = extensionOf(part).toLowerCase();
    for (const d of elements(doc, CT_NS, "Default")) {
      if ((d.getAttribute("Extension") ?? "").toLowerCase() === ext) return d.getAttribute("ContentType") ?? undefined;
    }
    return undefined;
  }

  /**
   * The deck's slide parts, in PRESENTATION ORDER.
   *
   * Never the zip's own order and never a sort of the file names: `slide10.xml`
   * sorts before `slide2.xml`, and a deck whose slides were reordered in
   * PowerPoint keeps its original part numbers. The `<p:sldIdLst>` is the only
   * statement of what order the deck is in.
   */
  async slidePaths(): Promise<string[]> {
    const doc = await this.doc(PRESENTATION);
    const list = element(doc, P_NS, "sldIdLst");
    if (!list) return [];
    const out: string[] = [];
    for (const sld of elements(list, P_NS, "sldId")) {
      const rId = sld.getAttributeNS(R_NS, "id") ?? sld.getAttribute("r:id");
      if (!rId) continue;
      const target = await this.relTarget(PRESENTATION, rId);
      if (target && !target.external) out.push(target.path);
    }
    return out;
  }

  /** The deck's slide size in EMU, from `<p:sldSz>`. */
  async slideSize(): Promise<{ cx: number; cy: number }> {
    const doc = await this.doc(PRESENTATION);
    const sz = element(doc, P_NS, "sldSz");
    const cx = Number(sz?.getAttribute("cx"));
    const cy = Number(sz?.getAttribute("cy"));
    if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
      throw new Error("ssf-slide-elements: this presentation does not state its slide size");
    }
    return { cx, cy };
  }

  /**
   * Reduce the deck to one slide, by presentation order.
   *
   * Used to hand PowerPoint a package containing only the slide that was
   * spliced. The alternative — inserting the whole deck and naming the wanted
   * slide through `insertSlidesFromBase64`'s `sourceSlideIds` — needs a host
   * behaviour this project has not measured, where the sibling add-in HAS
   * measured a plain `insertSlidesFromBase64` with a `targetSlideId` against a
   * real PowerPoint. Reaching for the proven call and doing the extra work in
   * the file is the cheaper risk.
   *
   * The other slides' PARTS are left in the package. They are unreferenced,
   * which OPC permits, and removing them properly means the referrer sweep a
   * sibling project spends five hundred lines on — every one of them there to
   * avoid deleting a part something else still needs. Nothing here reads the
   * package again, so unreferenced bytes cost one insert's bandwidth and
   * nothing else.
   */
  async keepOnlySlide(slidePath: string): Promise<void> {
    const doc = await this.doc(PRESENTATION);
    const list = element(doc, P_NS, "sldIdLst");
    if (!list) throw new Error("ssf-slide-elements: this presentation has no slide list");
    let kept = 0;
    for (const sld of elements(list, P_NS, "sldId")) {
      const rId = sld.getAttributeNS(R_NS, "id") ?? sld.getAttribute("r:id");
      const target = rId ? await this.relTarget(PRESENTATION, rId) : undefined;
      if (target && !target.external && target.path === slidePath) {
        kept += 1;
        continue;
      }
      sld.parentNode?.removeChild(sld);
    }
    if (kept === 0) throw new Error(`ssf-slide-elements: ${slidePath} is not in this presentation's slide list`);
    this.write(PRESENTATION, doc);
  }

  /** Copy a part's bytes to a new name. Content types are the caller's job. */
  async copyPart(from: string, to: string): Promise<void> {
    this.setBytes(to, await this.bytes(from));
  }

  partNames(): string[] {
    return Object.keys(this.zip.files).filter((n) => !this.zip.files[n]?.dir);
  }

  async toBase64(): Promise<string> {
    return this.zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  }

  async toBytes(): Promise<Uint8Array> {
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}

/**
 * Where a relationship target points, as a package path.
 *
 * Targets are relative to the OWNER's folder and routinely reach upwards:
 * `../media/image1.png` from `ppt/slides/slide1.xml` is `ppt/media/image1.png`.
 *
 * **A target that climbs past the package root is refused rather than
 * clamped.** `../../../../etc/passwd` resolving to `etc/passwd` would be a path
 * inside the zip that no legitimate deck names, and both callers would then
 * treat it as an ordinary part. There is nothing sensible to return for a
 * target that leaves the package, so this raises.
 */
export function resolveTarget(ownerPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segments = ownerPart.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "" || part === ".") continue;
    if (part === "..") {
      if (segments.length === 0) {
        throw new Error(`ssf-slide-elements: relationship target "${target}" reaches outside the package`);
      }
      segments.pop();
      continue;
    }
    segments.push(part);
  }
  return segments.join("/");
}

/** A path's extension, without the dot. Empty when it has none. */
export function extensionOf(part: string): string {
  const name = part.slice(part.lastIndexOf("/") + 1);
  const dot = name.lastIndexOf(".");
  return dot < 0 ? "" : name.slice(dot + 1);
}
