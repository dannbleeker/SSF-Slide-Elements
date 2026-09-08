#!/usr/bin/env node
/**
 * The ribbon icons, drawn in code rather than checked in as binaries.
 *
 * Office wants 16, 32, 64 and 80 pixel PNGs and a manifest that names each one
 * by URL. Four hand-made files is four things to keep in step with a palette
 * that lives in `src/pane/taskpane.css`, and a binary in a diff is a change
 * nobody can review. This writes them from the same two colours the pane uses,
 * so "the icon is off-brand" is a one-line change rather than a round trip
 * through an image editor.
 *
 * The encoder is a few dozen lines because a PNG of a flat shape does not need
 * a library: a raster, one zlib stream, three chunks and a CRC. `node:zlib`
 * does the only hard part.
 *
 *   node scripts/build-icons.mjs
 */
import { deflateSync } from "node:zlib";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { isMain } from "./is-main.mjs";

/** The pane's own colours. Navy ground, orange element, pale slide frame. */
export const NAVY = [0x00, 0x25, 0x4c];
export const ORANGE = [0xed, 0x89, 0x36];
export const PALE = [0xdd, 0xeb, 0xf7];

const CRC = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return (buf) => {
    let c = -1;
    for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  };
})();

function chunk(type, body) {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(body.length, 0);
  head.write(type, 4, "ascii");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(CRC(Buffer.concat([head.subarray(4), body])), 0);
  return Buffer.concat([head, body, crc]);
}

/**
 * An RGBA raster as a PNG.
 *
 * `pixel(x, y)` answers `[r, g, b, a]`. Colour type 6 (RGBA) and filter 0 on
 * every row: a flat shape compresses to nothing either way, and a filter that
 * is always zero is one less thing to get wrong.
 */
export function png(size, pixel) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let at = 0;
  for (let y = 0; y < size; y++) {
    raw[at++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x, y);
      raw[at++] = r;
      raw[at++] = g;
      raw[at++] = b;
      raw[at++] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/**
 * The mark: a navy square holding a pale slide frame with one orange element
 * inside it.
 *
 * That IS the product — an element, dropped onto a slide — and it is the only
 * thing this icon has room to say at sixteen pixels. It is deliberately not
 * the sibling's three-row mark: two add-ins in one ribbon with one icon is a
 * real confusion.
 *
 * Everything is in SIXTEENTHS of the icon rather than in pixels, so the same
 * arithmetic draws 16 and 80 and the proportions do not drift between them.
 */
export function markPixel(size) {
  const u = size / 16;
  const at = (n) => Math.round(n * u);
  const radius = Math.max(1, Math.round(size * 0.16));
  const thick = Math.max(1, at(1));
  // The slide, as a hollow frame; the element, as a filled block inside it.
  const frame = { left: at(3), right: at(13), top: at(4), bottom: at(12) };
  const block = { left: at(6), right: at(10), top: at(7), bottom: at(10) };

  const outside = (x, y) => {
    // Only the four corners can be outside a rounded square.
    const cx = x < radius ? radius - 0.5 : x > size - radius - 1 ? size - radius - 0.5 : x;
    const cy = y < radius ? radius - 0.5 : y > size - radius - 1 ? size - radius - 0.5 : y;
    if (cx === x && cy === y) return false;
    return Math.hypot(x - cx, y - cy) > radius;
  };
  const inFrame = (x, y) => x >= frame.left && x < frame.right && y >= frame.top && y < frame.bottom;
  const inInner = (x, y) =>
    x >= frame.left + thick && x < frame.right - thick && y >= frame.top + thick && y < frame.bottom - thick;

  return (x, y) => {
    if (outside(x, y)) return [0, 0, 0, 0];
    if (x >= block.left && x < block.right && y >= block.top && y < block.bottom) return [...ORANGE, 255];
    if (inFrame(x, y) && !inInner(x, y)) return [...PALE, 255];
    return [...NAVY, 255];
  };
}

/**
 * The sizes Office asks for. 16, 32 and 80 are the ribbon, 64 is the XML
 * manifest's high-resolution store icon, and 192 is the unified manifest's
 * colour icon — its v1.17 schema says 192x192 outright, and a sibling ships the
 * 64 there, which a tenant deployment would refuse.
 */
export const SIZES = [16, 32, 64, 80, 192];

/**
 * The unified manifest's monochrome OUTLINE icon.
 *
 * A different picture rather than the same one recoloured: it is drawn on a
 * transparent ground and stencilled by the host, so a navy square would come
 * back as a solid block. The frame and the element are what survive. Exported
 * so `test/manifest.test.ts` can pin the committed file to it byte for byte,
 * like the colour icons.
 */
export function outlinePixel(size) {
  const mark = markPixel(size);
  return (x, y) => {
    const [r, g, b, a] = mark(x, y);
    const isGround = r === NAVY[0] && g === NAVY[1] && b === NAVY[2];
    return a === 255 && !isGround ? [255, 255, 255, 255] : [0, 0, 0, 0];
  };
}

export function main() {
  const root = dirname(dirname(fileURLToPath(import.meta.url)));
  const out = join(root, "public", "assets");
  mkdirSync(out, { recursive: true });
  for (const size of SIZES) {
    writeFileSync(join(out, `icon-${size}.png`), png(size, markPixel(size)));
  }
  writeFileSync(join(out, "icon-outline-32.png"), png(32, outlinePixel(32)));
  console.log(`icons: ${SIZES.map((s) => `icon-${s}.png`).join(", ")}, icon-outline-32.png`);
}

if (isMain(import.meta.url)) main();
