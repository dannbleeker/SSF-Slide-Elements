/**
 * Is this slide still a slide PowerPoint will open?
 *
 * Every way a splice goes wrong produces the same message from PowerPoint —
 * "PowerPoint found a problem with content" — with no indication of which of
 * the four rewrites in `splice.ts` missed. That message costs a round trip to a
 * real host to see and tells you nothing once you get there, so the conditions
 * that cause it are checked here instead, in the suite, on every element.
 *
 * This is a STRUCTURAL check, not a rendering one. It cannot tell you the
 * element looks right; it tells you the file is not broken, which is the class
 * of failure that is both catastrophic and completely checkable.
 */
import { Pkg } from "../pptx/pkg.js";
import { MC_NS, P_NS, element, relationshipIdsIn } from "../pptx/xml.js";

export interface Finding {
  kind: "dangling-rel" | "missing-part" | "duplicate-shape-id" | "undeclared-type" | "unparsable";
  detail: string;
}

/**
 * Check one slide of a package.
 *
 * The four findings map one-to-one onto the four rewrites a splice performs,
 * deliberately: a red result here names which step failed rather than leaving
 * somebody to bisect it.
 */
export async function verifySlide(pkg: Pkg, slidePath: string): Promise<Finding[]> {
  const findings: Finding[] = [];

  let doc;
  try {
    doc = await pkg.doc(slidePath);
  } catch (e) {
    return [{ kind: "unparsable", detail: `${slidePath}: ${e instanceof Error ? e.message : String(e)}` }];
  }

  const tree = element(doc, P_NS, "spTree");
  if (!tree) return [{ kind: "unparsable", detail: `${slidePath} has no shape tree` }];

  // 3. Relationship ids. Every id the markup names must be defined by the
  // slide's own .rels, and its target must be a part that is actually there.
  const declared = new Map<string, string>();
  for (const rel of await pkg.rels(slidePath)) {
    const id = rel.getAttribute("Id");
    if (id) declared.set(id, rel.getAttribute("Target") ?? "");
  }
  for (const rId of relationshipIdsIn(tree)) {
    if (!declared.has(rId)) {
      findings.push({
        kind: "dangling-rel",
        detail: `${slidePath} names ${rId}, which its relationships do not define`,
      });
      continue;
    }
    const target = await pkg.relTarget(slidePath, rId);
    // 1 and 2. Part names and nested targets. A rename that missed a pointer
    // shows up here as a target nothing answers.
    if (target && !target.external && !pkg.has(target.path)) {
      findings.push({ kind: "missing-part", detail: `${rId} points at ${target.path}, which is not in the package` });
    }
    if (target && !target.external && (await pkg.contentTypeOf(target.path)) === undefined) {
      findings.push({
        kind: "undeclared-type",
        detail: `${target.path} has no content type; PowerPoint will refuse to load it`,
      });
    }
  }

  // 4. Shape ids, across the whole slide including inside groups.
  //
  // `<mc:Fallback>` is skipped, and that is not a loophole. An
  // `<mc:AlternateContent>` carries the same shape twice — a modern form in
  // `<mc:Choice>` and an older one in `<mc:Fallback>` — sharing one
  // `<p:cNvPr id>` on purpose, because only one branch is ever live. Counting
  // both reports every such shape as a collision; the shipped library's own
  // cover slide is one, so this check would have been crying wolf on the very
  // first slide anybody looked at.
  const seen = new Map<string, number>();
  const countIds = (node: Element): void => {
    if (node.localName === "Fallback" && node.namespaceURI === MC_NS) return;
    if (node.localName === "cNvPr" && node.namespaceURI === P_NS) {
      const id = node.getAttribute("id");
      if (id !== null) seen.set(id, (seen.get(id) ?? 0) + 1);
    }
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) countIds(c as Element);
    }
  };
  countIds(tree);
  for (const [id, count] of seen) {
    if (count > 1) {
      findings.push({ kind: "duplicate-shape-id", detail: `${count} shapes on ${slidePath} share id ${id}` });
    }
  }

  return findings;
}

/**
 * Every part a slide reaches, transitively, that is not in the package.
 *
 * Separate from `verifySlide` because it walks the whole graph — a chart's
 * workbook is two hops from the slide and a broken pointer there is invisible
 * until somebody double-clicks the chart.
 */
export async function verifyReachable(pkg: Pkg, slidePath: string): Promise<Finding[]> {
  const findings: Finding[] = [];
  const seen = new Set<string>();
  const queue = [slidePath];
  while (queue.length > 0) {
    const part = queue.shift();
    if (part === undefined || seen.has(part)) continue;
    seen.add(part);
    for (const rel of await pkg.rels(part)) {
      const target = await pkg.relTarget(part, rel.getAttribute("Id") ?? "");
      if (!target || target.external) continue;
      if (!pkg.has(target.path)) {
        findings.push({
          kind: "missing-part",
          detail: `${part} points at ${target.path}, which is not in the package`,
        });
        continue;
      }
      // Only XML parts have relationships worth following; media is a leaf.
      if (target.path.endsWith(".xml")) queue.push(target.path);
    }
  }
  return findings;
}

/** Both checks, for a caller that wants one answer. */
export async function verify(pkg: Pkg, slidePath: string): Promise<Finding[]> {
  return [...(await verifySlide(pkg, slidePath)), ...(await verifyReachable(pkg, slidePath))];
}
