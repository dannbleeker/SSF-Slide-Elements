/**
 * Who a committed deck says wrote it.
 *
 * PowerPoint stamps the signed-in account into any deck it saves, and a deck
 * carrying a comment gets stamped twice over: `docProps/core.xml` takes the
 * account's display name, and `ppt/authors.xml` — which exists only once there
 * is a modern comment to attribute — takes the display name, the initials, the
 * account's `userId` and the identity provider that issued it. Measured on
 * 2026-09-14 on the deck PowerPoint authored for the probe: those two parts,
 * and no others, named the owner — on THAT machine. `docProps/app.xml` was
 * added to the list on 2026-09-16 for the organisation the same Office writes
 * when it is configured with one, which this one is not; see `APP_FIELDS`.
 *
 * **This repository is public.** A deck is the one file a reviewer cannot read
 * before it is merged, so the check here is an ALLOW LIST rather than a search
 * for anything that looks like a person: the committed deck must name the
 * invented author below and nobody else. Re-authoring the deck from a live
 * PowerPoint and committing it without scrubbing is then a red test rather than
 * a published account id.
 *
 * Both functions take and return the package's TEXT parts as a plain object,
 * so they are pure and the test can feed them a part that never came from a
 * zip.
 */

import { readFileSync, writeFileSync } from "node:fs";
import JSZip from "jszip";
import { isMain } from "./is-main.mjs";

/** The invented author every committed deck is held to. Not a real account. */
export const PROBE_AUTHOR = {
  name: "SSF Probe",
  initials: "SP",
  userId: "ssf-probe",
  // PowerPoint's own value for an author with no identity provider behind it.
  providerId: "None",
};

const AUTHORS = "ppt/authors.xml";
const CORE = "docProps/core.xml";
const APP = "docProps/app.xml";

/**
 * The extended properties that name an ORGANISATION rather than a person.
 *
 * ASSUMED, not measured, and the distinction is the reason this is here. The
 * 2026-09-14 reading found the owner's name in the two parts above and nowhere
 * else — but that was one machine, whose Office install carries no company. The
 * extended-properties schema has carried `Company` and `Manager` since
 * ECMA-376, and an Office configured with them writes them into every file it
 * saves. So an allow list that never looks at this part stays GREEN on the
 * machine where it would matter, which is the shape of gate this repository
 * does not keep. Wanted empty rather than set to the invented author: a deck
 * belongs to nobody.
 */
const APP_FIELDS = ["Company", "Manager"];

/**
 * Every value the named attribute takes in a part.
 *
 * @param {string} xml
 * @param {string} name
 * @returns {string[]}
 */
function attributes(xml, name) {
  /** @type {string[]} */
  const found = [];
  for (const match of xml.matchAll(new RegExp(`\\b${name}="([^"]*)"`, "g"))) {
    if (match[1] !== undefined) found.push(match[1]);
  }
  return found;
}

/**
 * What one element holds, or null when the part has no such element.
 *
 * @param {string} xml
 * @param {string} name
 * @returns {string | null}
 */
function element(xml, name) {
  const match = xml.match(new RegExp(`<${name}[^>]*>([^<]*)</${name}>`));
  return match?.[1] ?? null;
}

/**
 * What is wrong with who this package says wrote it, as sentences.
 *
 * An empty array is the only passing answer. A part that is absent is not a
 * problem — a deck with no comment has no `ppt/authors.xml` — but a part that
 * is PRESENT and names somebody else is.
 *
 * @param {Record<string, string>} parts
 * @param {typeof PROBE_AUTHOR} [author]
 * @returns {string[]}
 */
export function identityProblems(parts, author = PROBE_AUTHOR) {
  /** @type {string[]} */
  const problems = [];
  const authors = parts[AUTHORS];
  if (typeof authors === "string") {
    for (const [attribute, wanted] of [
      ["name", author.name],
      ["initials", author.initials],
      ["userId", author.userId],
      ["providerId", author.providerId],
    ]) {
      for (const value of attributes(authors, attribute)) {
        if (value !== wanted) problems.push(`${AUTHORS} gives ${attribute}="${value}", not "${wanted}"`);
      }
    }
  }
  const core = parts[CORE];
  if (typeof core === "string") {
    for (const tag of ["dc:creator", "cp:lastModifiedBy"]) {
      const value = element(core, tag);
      if (value && value !== author.name) problems.push(`${CORE} gives ${tag} as "${value}", not "${author.name}"`);
    }
  }
  const app = parts[APP];
  if (typeof app === "string") {
    for (const tag of APP_FIELDS) {
      const value = element(app, tag);
      if (value) problems.push(`${APP} gives ${tag} as "${value}", and a committed deck names no organisation`);
    }
  }
  // A second net, under the allow list rather than instead of it: an address is
  // identifying wherever it turns up, including in a part this function has no
  // rule for.
  for (const [name, text] of Object.entries(parts)) {
    if (typeof text !== "string") continue;
    for (const match of text.matchAll(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g)) {
      problems.push(`${name} carries what reads as an address, ${match[0]}`);
    }
  }
  return problems;
}

/**
 * The same parts with the invented author in place of whoever PowerPoint
 * stamped. Parts this does not know about are returned untouched.
 *
 * @param {Record<string, string>} parts
 * @param {typeof PROBE_AUTHOR} [author]
 * @returns {Record<string, string>}
 */
export function anonymise(parts, author = PROBE_AUTHOR) {
  /** @type {Record<string, string>} */
  const out = { ...parts };
  if (typeof out[AUTHORS] === "string") {
    out[AUTHORS] = out[AUTHORS].replace(/\bname="[^"]*"/g, `name="${author.name}"`)
      .replace(/\binitials="[^"]*"/g, `initials="${author.initials}"`)
      .replace(/\buserId="[^"]*"/g, `userId="${author.userId}"`)
      .replace(/\bproviderId="[^"]*"/g, `providerId="${author.providerId}"`);
  }
  if (typeof out[CORE] === "string") {
    out[CORE] = out[CORE].replace(/<dc:creator>[^<]*<\/dc:creator>/, `<dc:creator>${author.name}</dc:creator>`).replace(
      /<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/,
      `<cp:lastModifiedBy>${author.name}</cp:lastModifiedBy>`,
    );
  }
  if (typeof out[APP] === "string") {
    // Emptied rather than replaced. The invented author is a person the deck
    // may claim to be by; there is no invented company to stand in for one, and
    // an empty element is what an Office with no company set writes.
    for (const tag of APP_FIELDS) {
      out[APP] = out[APP].replace(new RegExp(`<${tag}>[^<]*</${tag}>`, "g"), `<${tag}></${tag}>`);
    }
  }
  return out;
}

/**
 * `node scripts/deck-identity.mjs <deck.pptx> [out.pptx]`
 *
 * With one argument it REPORTS, which is how a deck can be checked without
 * being rewritten; with two it writes the scrubbed copy. Only the two parts
 * above are rewritten, so the rest of the package is the bytes PowerPoint
 * saved.
 */
async function main(argv) {
  const [input, output] = argv;
  if (!input) {
    console.error("usage: node scripts/deck-identity.mjs <deck.pptx> [out.pptx]");
    return 2;
  }
  const zip = await JSZip.loadAsync(readFileSync(input));
  const text = {};
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (name.endsWith(".xml") || name.endsWith(".rels")) text[name] = await file.async("string");
  }
  if (!output) {
    const problems = identityProblems(text);
    for (const problem of problems) console.log(problem);
    console.log(problems.length === 0 ? "names nobody but the invented author" : `${problems.length} problems`);
    return problems.length === 0 ? 0 : 1;
  }
  const cleaned = anonymise(text);
  for (const [name, body] of Object.entries(cleaned)) {
    if (body !== text[name]) {
      zip.file(name, body);
      console.log(`rewrote ${name}`);
    }
  }
  writeFileSync(output, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  const left = identityProblems(cleaned);
  for (const problem of left) console.log(`STILL: ${problem}`);
  console.log(left.length === 0 ? `wrote ${output}` : `${output} still names somebody`);
  return left.length === 0 ? 0 : 1;
}

if (isMain(import.meta.url)) {
  process.exitCode = await main(process.argv.slice(2));
}
