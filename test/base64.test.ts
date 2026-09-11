import { Buffer as NodeBuffer } from "node:buffer";
import { describe, expect, it } from "vitest";
import { base64From, bytesFrom, routeFor } from "../src/core/pptx/base64.js";

/**
 * The three base64 routes, held against each other.
 *
 * This exists because of what it is FOR: every deck crosses base64 twice on
 * every insert, and the pane takes a different route through this file than the
 * suite does — Node has `Buffer`, the WebView has `atob`. A fallback nobody
 * runs is a fallback nobody has checked, so each route is forced here by
 * handing the module a `globalThis` with the faster ones taken away.
 */

/** A global object offering only the routes named. */
function only(...routes: ("standard" | "buffer" | "browser")[]): typeof globalThis {
  const fake = Object.create(globalThis) as Record<string, unknown> & typeof globalThis;
  if (!routes.includes("standard")) {
    // `Uint8Array` with the standard pair removed, and nothing else changed.
    const stripped = function StrippedUint8Array(...args: unknown[]) {
      return new (Uint8Array as unknown as new (...a: unknown[]) => Uint8Array)(...args);
    } as unknown as Uint8ArrayConstructor;
    Object.setPrototypeOf(stripped, Uint8Array);
    // `prototype` is read-only on a function type, so it is defined rather than
    // assigned: a fresh object inheriting from the real one, with the standard
    // pair deleted off the copy and the original left alone.
    Object.defineProperty(stripped, "prototype", { value: Object.create(Uint8Array.prototype) });
    delete (stripped as unknown as { fromBase64?: unknown }).fromBase64;
    delete (stripped.prototype as unknown as { toBase64?: unknown }).toBase64;
    fake.Uint8Array = stripped;
  }
  // Injected rather than inherited: under the test runner `Buffer` is not an
  // own property of `globalThis`, so a fake built on it offers no Buffer route
  // at all and this file would compare the browser route with itself.
  fake.Buffer = (routes.includes("buffer") ? NodeBuffer : undefined) as never;
  if (!routes.includes("browser")) {
    fake.atob = undefined as never;
    fake.btoa = undefined as never;
  }
  return fake;
}

const SAMPLES: [string, Uint8Array][] = [
  ["empty", new Uint8Array(0)],
  ["one byte", new Uint8Array([0])],
  ["two bytes, which pads to one =", new Uint8Array([255, 254])],
  ["three bytes, which pads to none", new Uint8Array([1, 2, 3])],
  ["every byte value", Uint8Array.from({ length: 256 }, (_, i) => i)],
  // Past the chunk the browser route works in, so its seam is crossed.
  ["a quarter of a megabyte", Uint8Array.from({ length: 262_144 }, (_, i) => (i * 31 + (i >> 8)) % 256)],
];

describe("the base64 routes", () => {
  it("takes the fastest route the platform it is on offers", () => {
    // Not hard-coded to one answer: the suite, the scripts and the pane are
    // three platforms. What must hold is the ORDER — the standard pair when it
    // exists, then Buffer, then the browser pair — and that whatever is chosen
    // is actually there.
    const chosen = routeFor();
    expect(chosen).toBeDefined();
    const standard = typeof (Uint8Array as { fromBase64?: unknown }).fromBase64 === "function";
    if (standard) expect(chosen).toBe("standard");
    expect(routeFor(only("standard", "buffer", "browser"))).toBe(standard ? "standard" : "buffer");
    expect(routeFor(only("buffer", "browser"))).toBe("buffer");
    expect(routeFor(only("browser"))).toBe("browser");
  });

  for (const [name, bytes] of SAMPLES) {
    it(`round-trips ${name}, the same way by every route`, () => {
      const available = (["standard", "buffer", "browser"] as const).filter((route) => routeFor(only(route)) === route);
      // At least two, or this case is comparing one implementation with itself.
      expect(available.length, "only one route is available here").toBeGreaterThan(1);
      const encoded = available.map((route) => base64From(bytes, only(route)));
      for (const text of encoded) {
        expect(text, `${name}: a route answered nothing`).toBeTypeOf("string");
        expect(text).toBe(encoded[0]);
      }
      for (const route of available) {
        const back = bytesFrom(encoded[0] as string, only(route));
        expect([...(back ?? [])], `${name} via ${route}`).toEqual([...bytes]);
      }
    });
  }

  it("answers undefined when the platform offers nothing, rather than guessing", () => {
    const barren = only();
    expect(routeFor(barren)).toBeUndefined();
    expect(bytesFrom("AAAA", barren)).toBeUndefined();
    expect(base64From(new Uint8Array([1]), barren)).toBeUndefined();
  });

  it("hands back bytes nothing else owns", () => {
    // Node's `Buffer.from` answers a view on a POOLED allocation. Handing that
    // out would let an unrelated allocation change a package's bytes
    // underneath it, which is the kind of defect that shows up as a corrupt
    // file on somebody else's machine.
    const bytes = bytesFrom("AAECAwQ=", only("buffer"));
    expect(bytes).toBeDefined();
    expect((bytes as Uint8Array).byteOffset).toBe(0);
    expect((bytes as Uint8Array).buffer.byteLength).toBe((bytes as Uint8Array).length);
  });
});
