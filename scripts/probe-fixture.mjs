/**
 * The small decks the probe's fixtures are built from.
 *
 * Ported from SSF-Merge (`scripts/probe-fixture.mjs`) on 2026-09-08 and given
 * a slide list: this probe's questions are about a package with MORE than one
 * slide in it (one listed, one only present), so the deck takes any number of
 * slides, each with its own text, creation id and, when asked, a package tag.
 *
 * Minimal on purpose: the question is whether PowerPoint accepts an inserted
 * slide, not whether it renders a chart, and every byte here has to travel
 * inside a snippet somebody pastes into an editor.
 *
 * Deterministic on purpose too: every zip entry carries the same fixed date,
 * so `npm run probe` writes the same bytes twice and CI can diff the committed
 * snippet against a fresh build.
 */
import JSZip from "jszip";

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
export const T = {
  slide: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  master: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  layout: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  pres: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  core: "application/vnd.openxmlformats-package.core-properties+xml",
  app: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  tags: "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
};
export const RT = {
  doc: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  master: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  app: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  tags: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags",
};

/** The one date every zip entry carries, so two builds of the same deck are byte-identical. */
export const FIXED_DATE = new Date(Date.UTC(2026, 8, 8, 12, 0, 0));

/**
 * A theme with all three of its required children.
 *
 * `CT_BaseStyles` requires clrScheme, fontScheme and fmtScheme in that order,
 * every one of them mandatory. SSF-Merge's first real sheet came back
 * `InvalidArgument` from every insert because this part was
 * `<a:themeElements/>`, and `KeepSourceFormatting` is precisely the path that
 * has to import the source theme.
 */
const clr = (n, v) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
const font = (n) => `<a:${n}><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:${n}>`;
const SOLID = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
const LN = `<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">${SOLID}<a:prstDash val="solid"/></a:ln>`;
const THEME =
  `<a:clrScheme name="Probe">` +
  clr("dk1", "000000") +
  clr("lt1", "FFFFFF") +
  clr("dk2", "44546A") +
  clr("lt2", "E7E6E6") +
  clr("accent1", "4472C4") +
  clr("accent2", "ED7D31") +
  clr("accent3", "A5A5A5") +
  clr("accent4", "FFC000") +
  clr("accent5", "5B9BD5") +
  clr("accent6", "70AD47") +
  clr("hlink", "0563C1") +
  clr("folHlink", "954F72") +
  `</a:clrScheme>` +
  `<a:fontScheme name="Probe">${font("majorFont")}${font("minorFont")}</a:fontScheme>` +
  `<a:fmtScheme name="Probe">` +
  `<a:fillStyleLst>${SOLID}${SOLID}${SOLID}</a:fillStyleLst>` +
  `<a:lnStyleLst>${LN}${LN}${LN}</a:lnStyleLst>` +
  `<a:effectStyleLst>${`<a:effectStyle><a:effectLst/></a:effectStyle>`.repeat(3)}</a:effectStyleLst>` +
  `<a:bgFillStyleLst>${SOLID}${SOLID}${SOLID}</a:bgFillStyleLst>` +
  `</a:fmtScheme>`;

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * A 16:9 deck of the given slides, every one listed.
 *
 * @param {{ text: string; creationId: number; tags?: [string, string][] }[]} slides
 *   One entry per slide, in deck order. `tags` writes a `ppt/tags/tagN.xml`
 *   part related from the slide and referenced from its `<p:custDataLst>`,
 *   which is how the add-in will mark what it inserted: in the FILE, before the
 *   insert, because a tag write through the API is refused on the web.
 * @returns {Promise<Uint8Array>}
 */
export async function makeDeck(slides) {
  const zip = new JSZip();
  const file = (path, content) => zip.file(path, content, { date: FIXED_DATE });

  const overrides = slides
    .map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${T.slide}"/>`)
    .concat(
      slides.flatMap((s, i) =>
        s.tags ? [`<Override PartName="/ppt/tags/tag${i + 1}.xml" ContentType="${T.tags}"/>`] : [],
      ),
    )
    .join("");
  file(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="${T.pres}"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${T.master}"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${T.layout}"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="${T.theme}"/>` +
      overrides +
      `<Override PartName="/docProps/core.xml" ContentType="${T.core}"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="${T.app}"/></Types>`,
  );
  file(
    "_rels/.rels",
    `${HEAD}<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${RT.doc}" Target="ppt/presentation.xml"/>` +
      `<Relationship Id="rId2" Type="${RT.core}" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="${RT.app}" Target="docProps/app.xml"/></Relationships>`,
  );
  // Not required by the OPC spec, and present in every deck PowerPoint writes.
  // Cheap insurance on a fixture whose rejection costs a whole round to diagnose.
  file(
    "docProps/core.xml",
    `${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>SSF Slide Elements probe</dc:title><cp:revision>1</cp:revision></cp:coreProperties>`,
  );
  file(
    "docProps/app.xml",
    `${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
      `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<Application>SSF Slide Elements</Application><Slides>${slides.length}</Slides></Properties>`,
  );
  // Slide ids from 256 up and relationship ids after the master's rId1, both
  // in deck order, so a test can read the list back and know which is which.
  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`).join("");
  file(
    "ppt/presentation.xml",
    `${HEAD}<p:presentation ${P} ${A} ${R}>` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
      `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  );
  const slideRels = slides
    .map((_, i) => `<Relationship Id="rId${2 + i}" Type="${RT.slide}" Target="slides/slide${i + 1}.xml"/>`)
    .join("");
  file(
    "ppt/_rels/presentation.xml.rels",
    `${HEAD}<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${RT.master}" Target="slideMasters/slideMaster1.xml"/>` +
      slideRels +
      `</Relationships>`,
  );
  file(
    "ppt/slideMasters/slideMaster1.xml",
    `${HEAD}<p:sldMaster ${P} ${A} ${R}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  );
  file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `${HEAD}<Relationships ${REL}>` +
      `<Relationship Id="rId1" Type="${RT.layout}" Target="../slideLayouts/slideLayout1.xml"/>` +
      `<Relationship Id="rId2" Type="${RT.theme}" Target="../theme/theme1.xml"/></Relationships>`,
  );
  file(
    "ppt/slideLayouts/slideLayout1.xml",
    `${HEAD}<p:sldLayout ${P} ${A} ${R} type="blank"><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `${HEAD}<Relationships ${REL}><Relationship Id="rId1" Type="${RT.master}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  );
  file(
    "ppt/theme/theme1.xml",
    `${HEAD}<a:theme ${A} name="Probe"><a:themeElements>${THEME}</a:themeElements></a:theme>`,
  );

  slides.forEach((spec, i) => {
    const n = i + 1;
    const tagRef = spec.tags ? `<p:custDataLst><p:tags r:id="rId2"/></p:custDataLst>` : "";
    file(
      `ppt/slides/slide${n}.xml`,
      `${HEAD}<p:sld ${P} ${A} ${R}><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Probe"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="1325563"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="4000"/><a:t>${spec.text}</a:t></a:r></a:p></p:txBody>` +
        `</p:sp></p:spTree>${tagRef}` +
        `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">` +
        `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${spec.creationId}"/>` +
        `</p:ext></p:extLst></p:cSld>` +
        `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );
    const tagRel = spec.tags ? `<Relationship Id="rId2" Type="${RT.tags}" Target="../tags/tag${n}.xml"/>` : "";
    file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `${HEAD}<Relationships ${REL}><Relationship Id="rId1" Type="${RT.layout}" Target="../slideLayouts/slideLayout1.xml"/>${tagRel}</Relationships>`,
    );
    if (spec.tags) {
      const tags = spec.tags.map(([name, val]) => `<p:tag name="${name}" val="${val}"/>`).join("");
      file(`ppt/tags/tag${n}.xml`, `${HEAD}<p:tagLst ${P}>${tags}</p:tagLst>`);
    }
  });
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
