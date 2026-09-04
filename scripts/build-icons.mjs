/**
 * The four icons, generated rather than committed as art.
 *
 * A ribbon icon at 16px is a shape and two colours; drawing it in a design tool
 * and committing four PNGs means four binaries nobody can review and a set that
 * drifts the moment one is re-exported. This writes them from one description.
 *
 * The mark is the SSF tick over a slide frame — the same tick the pane uses as
 * its one orange element.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateSync } from "node:zlib";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "public", "assets");
mkdirSync(out, { recursive: true });

const NAVY = [0x00, 0x25, 0x4c];
const ORANGE = [0xed, 0x89, 0x36];
const WHITE = [0xff, 0xff, 0xff];

/** A minimal PNG writer: RGBA, no filtering, one IDAT. */
function png(width, height, pixels) {
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (width * 4 + 1)] = 0;
    for (let x = 0; x < width; x++) {
      const p = pixels(x, y);
      const at = y * (width * 4 + 1) + 1 + x * 4;
      raw[at] = p[0];
      raw[at + 1] = p[1];
      raw[at + 2] = p[2];
      raw[at + 3] = p[3];
    }
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

let table;
function crc32(buf) {
  if (!table) {
    table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
  }
  let c = -1;
  for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8);
  return c ^ -1;
}

/** A slide frame in navy with an orange tick across its lower half. */
function icon(size) {
  const s = (v) => Math.round((v * size) / 32);
  const frame = { x0: s(3), y0: s(6), x1: s(29), y1: s(26) };
  const border = Math.max(1, s(2));
  return (x, y) => {
    const inFrame = x >= frame.x0 && x < frame.x1 && y >= frame.y0 && y < frame.y1;
    const onBorder =
      inFrame && (x < frame.x0 + border || x >= frame.x1 - border || y < frame.y0 + border || y >= frame.y1 - border);
    // The tick: two strokes meeting low-left, rising to the right.
    const tx = (x - s(9)) / size;
    const ty = (y - s(20)) / size;
    const up = Math.abs(ty + (x - s(14)) / size) < 2.2 / size && x >= s(14) && x <= s(23);
    const down = Math.abs(ty - (x - s(14)) / size) < 2.2 / size && x >= s(10) && x <= s(14);
    void tx;
    if (up || down) return [...ORANGE, 255];
    if (onBorder) return [...NAVY, 255];
    if (inFrame) return [...WHITE, 255];
    return [0, 0, 0, 0];
  };
}

for (const size of [16, 32, 64, 80]) {
  writeFileSync(join(out, `icon-${size}.png`), png(size, size, icon(size)));
  console.log(`public/assets/icon-${size}.png`);
}
