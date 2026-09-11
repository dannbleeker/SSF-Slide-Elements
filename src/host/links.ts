/**
 * The two links that leave the pane, and what they are allowed to carry.
 *
 * `docs/DESIGN.md` section 7 puts both behind the gear: **Report a problem**,
 * which opens the support page with the build stamp, the host and the platform
 * already filled in, and **Browse the catalogue on the site**, which opens the
 * catalogue page. Neither navigates the pane itself.
 *
 * The decisions live here, away from Office.js, for the usual reason: what a
 * URL may carry is exactly the kind of rule that must be testable. And this one
 * is a privacy rule as much as a formatting one — the support page is on the
 * open web, a URL is the least private thing there is, and the pane holds the
 * user's whole presentation. So the rule is an ALLOWLIST of three facts about
 * the software, never about the user:
 *
 * - the build stamp, which is the seven characters in the pane's header;
 * - the host, which is PowerPoint;
 * - the platform, which is one of Office's own handful of names.
 *
 * Nothing from the deck — no file name, no slide text, no element id — has a
 * way in, because nothing else is a parameter. A value that does not look like
 * what it claims to be is dropped rather than passed on: the support page reads
 * these back out and shows them, so a value from anywhere else must not be able
 * to reach it.
 */

/** Where the site lives, as the pane sees it. The pane is served from it. */
export interface Site {
  /** The origin the pane was served from, with no trailing slash. */
  origin: string;
}

/** The three facts a report may carry, as the pane knows them. */
export interface Stamp {
  /** The commit the pane was built from: hexadecimal, seven characters or more. */
  build?: string | undefined;
  /** `Office.context.diagnostics.host`, or the older `Office.context.host`. */
  host?: string | undefined;
  /** `Office.context.diagnostics.platform`, or the older `Office.context.platform`. */
  platform?: string | undefined;
}

/**
 * Office's own platform names, and the only ones this passes on.
 *
 * From the `PlatformType` enum. An unknown value is dropped: the point of the
 * allowlist is that a string from the host cannot become arbitrary text in a
 * URL on a public page, and "some future platform is not named in the report"
 * is a far smaller cost than the alternative.
 */
const PLATFORMS = ["PC", "OfficeOnline", "Mac", "iOS", "Android", "Universal"];

/** The hosts this add-in runs in. It is a PowerPoint add-in; the rest are the enum's. */
const HOSTS = ["PowerPoint", "Word", "Excel", "Outlook", "OneNote", "Project", "Access"];

/** A build stamp is a commit: hexadecimal and nothing else. */
const BUILD = /^[0-9a-f]{7,40}$/;

/**
 * The support page, with what is known about this pane already in it.
 *
 * Every parameter is optional in both directions: the pane may not know the
 * platform (it is undefined outside a host), and the page works with none of
 * them. A parameter that is not known is left out rather than sent as
 * "unknown", so the page can tell "not reported" from "reported as unknown".
 */
export function reportUrl(site: Site, stamp: Stamp): string {
  const parameters = new URLSearchParams();
  if (stamp.build && BUILD.test(stamp.build)) parameters.set("build", stamp.build);
  if (stamp.host && HOSTS.includes(stamp.host)) parameters.set("host", stamp.host);
  if (stamp.platform && PLATFORMS.includes(stamp.platform)) parameters.set("platform", stamp.platform);
  const query = parameters.toString();
  return `${site.origin}/support.html${query ? `?${query}` : ""}`;
}

/** The catalogue page: every element in both sizes, to look at outside PowerPoint. */
export function catalogueUrl(site: Site): string {
  return `${site.origin}/catalogue.html`;
}

/**
 * The origin a pane served from `url` should use for its own pages.
 *
 * Its own, always. The links go to the support and catalogue pages of the site
 * this pane came from — the dev build's links point at the dev origin, and a
 * fork's at the fork's — and a hard-coded production URL in the pane would be
 * the one thing here that could send somebody somewhere the add-in did not come
 * from. Answers undefined for anything that is not an http(s) page, which is
 * what a pane opened from a file looks like.
 */
export function siteFrom(url: string): Site | undefined {
  try {
    const parsed = new URL(url);
    if (parsed.protocol !== "https:" && parsed.protocol !== "http:") return undefined;
    return { origin: parsed.origin };
  } catch {
    return undefined;
  }
}
