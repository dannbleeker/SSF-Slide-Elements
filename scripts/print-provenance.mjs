/**
 * Whether a committed PDF print still belongs to the deck beside it.
 *
 * `docs/DESIGN.md` section 3 cuts every element's preview out of a print of the
 * deck, and the only thing it checks is that the print's page count matches the
 * deck's slide count. That is not a gate. 110 pages match 110 slides whatever
 * the pages are OF: a print of yesterday's deck, of a corrected copy, or of a
 * different deck with the same number of slides all pass it, and the previews
 * come out cut from the wrong file with nothing to say so. It nearly happened
 * here — the first pair of prints was taken from slash-corrected COPIES while
 * the committed decks would not open, and was thrown away rather than committed
 * for exactly this reason.
 *
 * So each print carries a sidecar naming the bytes it was taken from, and the
 * rules below are what CI holds it to. The rules live here rather than in the
 * test so a script and the suite cannot read different ones — the same reason
 * `scripts/package-integrity.mjs` exists.
 */

import { createHash } from "node:crypto";

/** SHA-256 of a buffer, as lowercase hex. */
export function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

/**
 * How many pages a PDF has.
 *
 * Read two ways, because one way is a guess. The page tree's root `/Count` is
 * the document's own answer; counting `/Type /Page` objects is independent of
 * it. PowerPoint writes both in the clear even though the file carries object
 * streams. When they disagree, or either is missing, this returns null rather
 * than a number somebody would then trust.
 */
export function pdfPageCount(bytes) {
  const text = Buffer.from(bytes).toString("latin1");
  const counts = [...text.matchAll(/\/Count\s+(\d+)/g)].map((m) => Number(m[1]));
  // PowerPoint writes the compact form, `/Type/Page/Parent`, so the lookahead
  // may not exclude a following slash — only a letter, which is what separates
  // `/Page` from `/Pages`.
  const objects = [...text.matchAll(/\/Type\s*\/Page(?![A-Za-z])/g)].length;
  if (!counts.length || !objects) return null;
  const declared = Math.max(...counts);
  return declared === objects ? declared : null;
}

/** How many slides a .pptx has, from its own `<p:sldIdLst>`. */
export function slideCountOf(presentationXml) {
  return [...presentationXml.matchAll(/<p:sldId\s/g)].length;
}

/** The sidecar a freshly taken print should carry. */
export function stampFor({ deck, print, deckBytes, printBytes, slides, powerPoint, takenAt, how }) {
  return {
    deck,
    deckSha256: sha256(deckBytes),
    slides,
    print,
    printSha256: sha256(printBytes),
    pages: pdfPageCount(printBytes),
    takenAt,
    powerPoint,
    how,
  };
}

/**
 * Everything wrong with a print and its sidecar, as sentences.
 *
 * Empty means the print is of the deck committed beside it, at the size the
 * deck actually is. Anything else names which of the two moved.
 */
export function printProblems({ name, stamp, deckBytes, printBytes, slides }) {
  const problems = [];
  if (!stamp) {
    problems.push(`${name}: no print sidecar, so nothing says which deck the print came from`);
    return problems;
  }

  const deckNow = sha256(deckBytes);
  if (stamp.deckSha256 !== deckNow) {
    problems.push(
      `${name}: the deck has changed since the print was taken ` +
        `(sidecar ${stamp.deckSha256.slice(0, 12)}, deck now ${deckNow.slice(0, 12)}) — re-print it`,
    );
  }

  const printNow = sha256(printBytes);
  if (stamp.printSha256 !== printNow) {
    problems.push(
      `${name}: the print has changed since it was stamped ` +
        `(sidecar ${stamp.printSha256.slice(0, 12)}, print now ${printNow.slice(0, 12)}) — re-stamp it`,
    );
  }

  if (stamp.slides !== slides) {
    problems.push(`${name}: sidecar says ${stamp.slides} slides, the deck has ${slides}`);
  }

  const pages = pdfPageCount(printBytes);
  if (pages === null) {
    problems.push(`${name}: could not read the print's page count two ways that agree`);
  } else {
    if (stamp.pages !== pages) {
      problems.push(`${name}: sidecar says ${stamp.pages} pages, the print has ${pages}`);
    }
    if (pages !== slides) {
      problems.push(
        `${name}: ${pages} pages for ${slides} slides — one page per slide, or the cuts land on the wrong page`,
      );
    }
  }

  for (const field of ["takenAt", "powerPoint", "how"]) {
    if (!stamp[field]) problems.push(`${name}: the sidecar does not say ${field}`);
  }

  return problems;
}
