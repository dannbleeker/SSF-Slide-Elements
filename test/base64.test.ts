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

/** Whether this platform ships the standard pair itself, read once before anything borrows it. */
const ORIGINAL_STANDARD =
  typeof (Uint8Array as unknown as { fromBase64?: unknown }).fromBase64 === "function" &&
  typeof (Uint8Array.prototype as unknown as { toBase64?: unknown }).toBase64 === "function";

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
    // Shadowed with `undefined`, not DELETED. Both the copy and its prototype
    // INHERIT from the real ones, and deleting a property a thing does not own
    // does nothing at all — so on any platform that really has the standard
    // pair, a fake built to have it taken away still had it, and every case
    // below that asks for a slower route would have been handed the fastest.
    // It reads as correct here only because this Node has neither.
    Object.defineProperty(stripped, "fromBase64", { value: undefined });
    Object.defineProperty(stripped.prototype, "toBase64", { value: undefined });
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

/**
 * Run something with the standard pair in place, whether or not this platform
 * has it.
 *
 * Chromium has `Uint8Array.fromBase64` from 133 and Node from 22.13. The Node
 * running this suite has neither — so the route the PANE takes is the one route
 * this file never executed, while the module's own comment says all three are
 * held against each other. Coverage said the same thing out loud: the two lines
 * that call the standard pair were never reached.
 *
 * The stand-in is `Buffer`, and that is deliberate: what is being checked is
 * that the module CALLS the pair correctly — the static one on the constructor,
 * the instance one on the bytes — not that the platform's implementation is
 * right. On a platform that has the real pair this changes nothing and the same
 * case runs against it.
 */
function withStandardPair<T>(run: () => T): T {
  const ctor = Uint8Array as unknown as { fromBase64?: unknown };
  const proto = Uint8Array.prototype as unknown as { toBase64?: unknown };
  if (typeof ctor.fromBase64 === "function" && typeof proto.toBase64 === "function") return run();
  Object.defineProperty(Uint8Array, "fromBase64", {
    configurable: true,
    value: (text: string) => new Uint8Array(NodeBuffer.from(text, "base64")),
  });
  Object.defineProperty(Uint8Array.prototype, "toBase64", {
    configurable: true,
    value(this: Uint8Array) {
      return NodeBuffer.from(this).toString("base64");
    },
  });
  try {
    return run();
  } finally {
    delete ctor.fromBase64;
    delete proto.toBase64;
  }
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

  it("takes the standard pair where there is one, and reaches it the way the platform names it", () => {
    // The route the pane takes on any current Chromium, and the one this file
    // could not run: the suite's Node has neither half of the pair. The lines
    // that call them were reached by nothing until this case existed.
    withStandardPair(() => {
      expect(routeFor(only("standard", "buffer", "browser"))).toBe("standard");
      for (const [name, bytes] of SAMPLES) {
        const viaStandard = base64From(bytes, only("standard", "buffer", "browser"));
        const viaBuffer = base64From(bytes, only("buffer"));
        expect(viaStandard, `${name}: the standard route answered nothing`).toBeTypeOf("string");
        expect(viaStandard, name).toBe(viaBuffer);
        const back = bytesFrom(viaStandard as string, only("standard", "buffer", "browser"));
        expect([...(back ?? [])], name).toEqual([...bytes]);
      }
    });
    // And it is gone again afterwards, or every case above this one would be
    // measuring a different platform from every case below it.
    expect(typeof (Uint8Array as unknown as { fromBase64?: unknown }).fromBase64).toBe(
      ORIGINAL_STANDARD ? "function" : "undefined",
    );
  });

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
