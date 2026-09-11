#!/usr/bin/env node
/**
 * What the ENGINE costs on a deck the size of a real one.
 *
 *   npm run build:lib && node scripts/bench-engine.mjs
 *
 * Not in CI and not a gate: it builds a 45 MB deck in memory and takes about a
 * minute. It is here so the numbers in `docs/DESIGN.md` section 11 can be taken
 * again rather than believed.
 *
 * `docs/DESIGN.md` section 13's sixth open question is how long a 50 MB deck
 * takes. Half of that question is `getFileAsync`, which only a host can answer.
 * The other half is everything this repo does with the bytes once it has them,
 * and that half can be measured here: opening the package, sweeping it for the
 * add-in's tags, reading what a slide holds, splicing an element in, and
 * handing the base64 back.
 */
import { readFileSync } from "node:fs";
import { randomFillSync } from "node:crypto";
import { Pkg } from "../dist-lib/core/index.js";
import { harvest } from "../dist-lib/core/catalogue/harvest.js";
import { occupiedBoxes } from "../dist-lib/core/catalogue/boxes.js";
import { slideSize } from "../dist-lib/core/pptx/layout.js";
import { usedInDeck, writeShapeTags, TAG_ELEMENT, TAG_CATALOGUE } from "../dist-lib/core/pptx/tags.js";
import { splice } from "../dist-lib/core/splice/splice.js";

const names = JSON.parse(readFileSync("template/names.en.json", "utf8"));
const library = await harvest(await Pkg.open(new Uint8Array(readFileSync("template/library-16x9.pptx"))), {
  size: "16:9",
  names,
});
const element = library.catalogue.elements.find((e) => e.id === "hvid-kasse-2x1-vertikale");

const ms = async (label, work) => {
  const at = performance.now();
  const out = await work();
  const took = performance.now() - at;
  return { label, ms: Math.round(took), out };
};

/**
 * A deck of `slides` slides and roughly `mb` megabytes of media.
 *
 * Built by cloning the validators' deck's own slides, so every slide is one
 * PowerPoint wrote, and padded with incompressible bytes in `ppt/media/` —
 * a deck is mostly pictures, and pictures are already compressed, so random
 * bytes are a better model of the weight than a repeated pattern.
 */
async function bigDeck(slides, mb) {
  const pkg = await Pkg.open(new Uint8Array(readFileSync("template/validators.pptx")));
  const { cloneSlide } = await import("../dist-lib/core/pptx/clone.js");
  const paths = await pkg.slidePaths();
  while ((await pkg.slidePaths()).length < slides) {
    const source = paths[(await pkg.slidePaths()).length % paths.length];
    await cloneSlide(pkg, source, { list: true });
  }
  if (mb > 0) {
    // Really random: the first version of this filled the buffer with a
    // deterministic sequence, which deflated to nothing and produced a "45 MB"
    // deck of 400 KB. A benchmark whose input is not the size it claims is a
    // benchmark measuring the wrong thing.
    const bytes = new Uint8Array(mb * 1024 * 1024);
    for (let at = 0; at < bytes.length; at += 65536) {
      randomFillSync(bytes, at, Math.min(65536, bytes.length - at));
    }
    pkg.setBytes("ppt/media/weight.bin", bytes);
    await pkg.addContentTypeOverride("/ppt/media/weight.bin", "application/octet-stream");
  }
  // Stamp a few slides so the tag sweep has something to find.
  const all = await pkg.slidePaths();
  for (const path of [all[0], all[Math.floor(all.length / 2)], all[all.length - 1]]) {
    const doc = await pkg.doc(path);
    const shape = doc.getElementsByTagName("p:sp")[0];
    if (shape)
      await writeShapeTags(pkg, path, shape, [
        [TAG_ELEMENT, "markeringer-1"],
        [TAG_CATALOGUE, "bench"],
      ]);
  }
  return pkg.toBase64();
}

const shapes = [
  { slides: 40, mb: 0 },
  { slides: 200, mb: 0 },
  { slides: 60, mb: 20 },
  { slides: 120, mb: 45 },
];

for (const shape of shapes) {
  const base64 = await bigDeck(shape.slides, shape.mb);
  const megabytes = ((base64.length * 3) / 4 / 1024 / 1024).toFixed(1);
  const opened = await ms("Pkg.open", () => Pkg.open(base64));
  const pkg = opened.out;
  const count = (await pkg.slidePaths()).length;
  const used = await ms("usedInDeck", () => usedInDeck(pkg));
  const size = await slideSize(pkg);
  const held = await ms("occupiedBoxes (one slide)", async () =>
    occupiedBoxes(await pkg.doc((await pkg.slidePaths())[0]), size.width, size.height),
  );
  const spliced = await ms("splice", () =>
    splice({
      deck: base64,
      slide: 0,
      element: {
        id: element.id,
        name: element.name,
        kind: element.kind,
        box: element.box,
        landing: element.landing,
        markup: element.markup,
      },
      options: { target: "onto", group: true, colours: "deck" },
      catalogue: { version: "bench", carried: library.catalogue.carried, theme: library.catalogue.theme },
      store: (path) => Promise.resolve(library.parts.get(path)),
    }),
  );
  console.log(
    `${String(count).padStart(3)} slides, ${megabytes.padStart(5)} MB  |  ` +
      `open ${String(opened.ms).padStart(5)} ms  |  usedInDeck ${String(used.ms).padStart(5)} ms (${used.out.length} found)  |  ` +
      `slide boxes ${String(held.ms).padStart(3)} ms  |  splice ${String(spliced.ms).padStart(5)} ms`,
  );
}
