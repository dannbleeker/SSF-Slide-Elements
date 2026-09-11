/**
 * Base64, through whatever the platform does fastest.
 *
 * Every deck this add-in touches crosses two base64 boundaries: `getFileAsync`
 * hands the presentation over as base64, and `insertSlidesFromBase64` takes it
 * back the same way. Neither is negotiable — they are Office's own signatures.
 * What IS negotiable is who does the conversion.
 *
 * Measured on this machine on 2026-09-11, on 45 MB of incompressible bytes —
 * which is what a deck full of photographs looks like:
 *
 * | route                                    | decode  | encode  |
 * | ---------------------------------------- | ------- | ------- |
 * | JSZip's own (`{ base64: true }`)         | 1608 ms | 2536 ms |
 * | the platform's (`Buffer`, here)          |   27 ms |   24 ms |
 *
 * That is the difference between an insert that costs four seconds of pure
 * character-shuffling on a big deck and one that costs a tenth of a second, and
 * it is the largest single cost in the whole path — bigger than the zip, bigger
 * than the XML, bigger than anything the splice does. `docs/DESIGN.md` section
 * 13's sixth open question asks how a 50 MB deck behaves; this is the half of
 * the answer that does not need a host.
 *
 * Three routes, best first, because no single one is available everywhere:
 *
 * 1. **`Uint8Array.fromBase64` / `toBase64`** — the standard pair. Chromium has
 *    it from 133 and Node from 22.13; the Node running this suite does not, so
 *    it is tried and not assumed.
 * 2. **`Buffer`** — Node, which is the suite, the scripts and nothing a user
 *    runs.
 * 3. **`atob` / `btoa`** — every browser including the WebView the pane lives
 *    in, worked in chunks because `String.fromCharCode.apply` on a megabyte of
 *    arguments overflows the call stack.
 *
 * All three are exercised by `test/base64.test.ts` against each other, because
 * a fallback nobody runs is a fallback nobody has checked — and the one that
 * runs in the PANE is the one this suite would otherwise never take. That
 * sentence was not true when it was first written: the suite's Node has neither
 * half of the standard pair, so the two lines that call it were reached by
 * nothing. The pair is now stood in for with `Buffer` for the length of one
 * case, which checks that this module CALLS it correctly — the static one on
 * the constructor, the instance one on the bytes.
 */

/** Characters per chunk for the `atob`/`btoa` route: a multiple of 3 keeps encoding on byte boundaries. */
const CHUNK = 0xfffc;

type Codec = "standard" | "buffer" | "browser";

/** Which route this platform can take, in order of speed. */
export function routeFor(globals: typeof globalThis = globalThis): Codec | undefined {
  const u8 = globals.Uint8Array as unknown as { fromBase64?: unknown; prototype?: { toBase64?: unknown } };
  if (typeof u8?.fromBase64 === "function" && typeof u8.prototype?.toBase64 === "function") return "standard";
  const buffer = (globals as { Buffer?: { from?: unknown } }).Buffer;
  if (typeof buffer?.from === "function") return "buffer";
  if (typeof globals.atob === "function" && typeof globals.btoa === "function") return "browser";
  return undefined;
}

/**
 * Base64 into bytes.
 *
 * Answers undefined when no route is available, so the caller can hand the
 * string to JSZip as it always did. Being slow is not a failure; being wrong
 * is, and a decoder that guessed would be.
 */
export function bytesFrom(base64: string, globals: typeof globalThis = globalThis): Uint8Array | undefined {
  switch (routeFor(globals)) {
    case "standard": {
      const u8 = globals.Uint8Array as unknown as { fromBase64: (s: string) => Uint8Array };
      return u8.fromBase64(base64);
    }
    case "buffer": {
      const buffer = (globals as unknown as { Buffer: { from: (s: string, e: string) => Uint8Array } }).Buffer;
      const made = buffer.from(base64, "base64");
      // A Node Buffer IS a Uint8Array, but it is a VIEW on a pooled allocation:
      // handing the pool out would let an unrelated allocation change these
      // bytes underneath the package.
      return new Uint8Array(made);
    }
    case "browser": {
      const binary = globals.atob(base64);
      const out = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
      return out;
    }
    default:
      return undefined;
  }
}

/** Bytes into base64, by the same three routes and the same rule about undefined. */
export function base64From(bytes: Uint8Array, globals: typeof globalThis = globalThis): string | undefined {
  switch (routeFor(globals)) {
    case "standard": {
      return (bytes as unknown as { toBase64: () => string }).toBase64();
    }
    case "buffer": {
      const buffer = (
        globals as unknown as { Buffer: { from: (b: Uint8Array) => { toString: (e: string) => string } } }
      ).Buffer;
      return buffer.from(bytes).toString("base64");
    }
    case "browser": {
      let binary = "";
      for (let at = 0; at < bytes.length; at += CHUNK) {
        binary += String.fromCharCode(...bytes.subarray(at, Math.min(at + CHUNK, bytes.length)));
      }
      return globals.btoa(binary);
    }
    default:
      return undefined;
  }
}
