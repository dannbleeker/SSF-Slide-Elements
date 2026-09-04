/**
 * A host error, as something a person can read.
 *
 * Office echoes the failing argument back through `debugInfo`, and the argument
 * here is AN ENTIRE PRESENTATION as base64. Uncapped, a failed insert puts
 * megabytes of the user's own deck on screen as the failure sentence, with the
 * explanation at the front and nothing after it readable — and into a bug
 * report, if they paste it. SSF-Merge learned this; the cap is the whole point
 * of the function.
 */
const CAP = 400;

export function readable(e: unknown): string {
  const raw =
    e instanceof Error
      ? e.message
      : typeof e === "string"
        ? e
        : typeof e === "object" && e !== null && "message" in e
          ? String(e.message)
          : "PowerPoint refused the request and said nothing about why";
  const collapsed = raw.replace(/\s+/g, " ").trim();
  // A cap is not a break: 400 characters with no space in them is one word, and
  // one word 400 characters long takes the pane's layout with it.
  const capped = collapsed.length > CAP ? `${collapsed.slice(0, CAP)}…` : collapsed;
  return capped || "PowerPoint refused the request and said nothing about why";
}
