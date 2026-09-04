/**
 * What a slide's relationships are CALLED, in one place.
 *
 * These strings decide which parts travel with a harvested element and which
 * stay behind in the library deck. SSF-Merge kept the same list spread across
 * four files; the copies agreed and nothing had gone wrong, but PowerPoint has
 * already added a second spelling of a relationship type once, and adding a
 * third to one copy and not the other is how a copy quietly stops matching.
 */
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_REL = "http://schemas.microsoft.com/office/2007/relationships";
const MS_REL_2014 = "http://schemas.microsoft.com/office/2014/relationships";

export const REL_TYPE = {
  slide: `${REL}/slide`,
  slideLayout: `${REL}/slideLayout`,
  notesSlide: `${REL}/notesSlide`,
  tags: `${REL}/tags`,
  image: `${REL}/image`,
  chart: `${REL}/chart`,
  /**
   * A callout, arrow or text box drawn ON a chart, in its own drawing part.
   * The chart owns the relationship, not the slide, so an element carrying a
   * chart has to carry this too or the callout's text vanishes.
   */
  chartUserShapes: `${REL}/chartUserShapes`,
  /**
   * A MODERN chart — waterfall, funnel, treemap, sunburst, histogram, pareto,
   * box-and-whisker, region map. PowerPoint stores none of those as a
   * `<c:chartSpace>`; they are a separate part under a Microsoft namespace.
   *
   * The RELATIONSHIP is the stable thing to match on. On the slide these charts
   * sit inside `<mc:AlternateContent>`, whose `Requires` token is `cx1`, `cx2`
   * or `cx4` depending on which of three dated namespaces the layout came from
   * — a reader keying on the token misses the other two.
   */
  chartEx: `${MS_REL_2014}/chartEx`,
  diagramData: `${REL}/diagramData`,
  diagramLayout: `${REL}/diagramLayout`,
  diagramQuickStyle: `${REL}/diagramQuickStyle`,
  diagramColors: `${REL}/diagramColors`,
  diagramDrawing: `${MS_REL}/diagramDrawing`,
  /** A whole package inside the package: the workbook behind a chart. */
  package: `${REL}/package`,
  /** An OLE object — a linked or embedded workbook, drawn as a picture. */
  oleObject: `${REL}/oleObject`,
  hyperlink: `${REL}/hyperlink`,
} as const;

/**
 * A relationship whose target is somewhere else entirely, not a part.
 *
 * A hyperlink's target is a URL and carries `TargetMode="External"`. Copying
 * one as though it were a part looks for `https://example.com` inside the zip,
 * finds nothing, and either throws or writes an empty part over a name the
 * package already uses. The mode is the thing to check, not the type: an image
 * can be external too.
 */
export const EXTERNAL = "External";

/**
 * The parts an ELEMENT may drag along with it, as an allowlist.
 *
 * An allowlist rather than a folder prefix, and that is the security half. A
 * relationship target comes out of a deck, and a deck can be sent to somebody:
 * a crafted one naming `/ppt/presentation.xml` would otherwise put that part
 * into the travelling set and overwrite the target deck's own copy — producing
 * a file PowerPoint cannot open, from an insert that reported success.
 *
 * `resolveTarget` already refuses to escape the package, so this is the second
 * of two gates rather than the only one. Both are cheap and they fail
 * differently, which is the point.
 */
export const TRAVELS_WITH_ELEMENT =
  /^ppt\/(?:media\/[^/]+|(?:charts|diagrams|drawings|embeddings|tags)\/[^/]+|theme\/[^/]+)$/;

/**
 * Which relationship types are followed when gathering an element's parts.
 *
 * Everything an element can legally point at, minus the things that belong to
 * the SLIDE rather than to the shapes on it — the layout, the notes, the
 * slide's own tags. Those describe the library deck's slide, and an element
 * inserted into somebody else's deck must take none of them.
 */
export const FOLLOWED: ReadonlySet<string> = new Set([
  REL_TYPE.image,
  /**
   * Shape-level custom data, and the reason this list is not "everything except
   * the layout".
   *
   * A `<p:tags>` under a shape's `<p:nvPr>` is that SHAPE's data, not the
   * slide's — the shipped library has 41 of them on one slide alone, every one
   * named `THINKCELLSHAPEDONOTDELETE`. They were classified as slide furniture
   * here at first, which left 41 relationship ids in the markup pointing at
   * nothing: a file PowerPoint calls damaged, from an insert that reported
   * success. They are a few hundred bytes each and they are what keeps a
   * think-cell chart editable by think-cell after it lands.
   */
  REL_TYPE.tags,
  REL_TYPE.chart,
  REL_TYPE.chartEx,
  REL_TYPE.chartUserShapes,
  REL_TYPE.diagramData,
  REL_TYPE.diagramLayout,
  REL_TYPE.diagramQuickStyle,
  REL_TYPE.diagramColors,
  REL_TYPE.diagramDrawing,
  REL_TYPE.package,
  REL_TYPE.oleObject,
]);
