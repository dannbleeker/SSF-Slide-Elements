import { afterEach, describe, expect, it, vi } from "vitest";

/**
 * The one call the pane makes before it has decided anything.
 *
 * `src/office` mostly cannot run in the suite — it calls Office.js — but
 * `hostSupports` is a single expression over `Office.context`, and what it does
 * when that object is not what it expects decides whether an unsupported host
 * gets a sentence or a blank pane. `ready()` uses it to render "this
 * PowerPoint is too old", before anything else, so a raise here leaves the
 * user with nothing to read on exactly the host that needed the message.
 */
async function supportsWith(context: unknown): Promise<(v: string) => boolean> {
  vi.resetModules();
  (globalThis as unknown as { Office: unknown }).Office = { context };
  const mod = await import("../src/office/powerpoint.js");
  return mod.hostSupports;
}

afterEach(() => {
  delete (globalThis as unknown as { Office?: unknown }).Office;
});

describe("asking the host what it supports", () => {
  it("answers what the host says when the host can answer", async () => {
    // The ordinary path, so the guard below is not just "always false".
    const asked: string[] = [];
    const supports = await supportsWith({
      requirements: {
        isSetSupported: (name: string, v: string) => {
          asked.push(`${name} ${v}`);
          return v === "1.2";
        },
      },
    });
    expect(supports("1.2")).toBe(true);
    expect(supports("1.10")).toBe(false);
    expect(asked).toEqual(["PowerPointApi 1.2", "PowerPointApi 1.10"]);
  });

  it("answers false rather than throwing when there is no requirements object", async () => {
    const supports = await supportsWith({});
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when there is no context at all", async () => {
    const supports = await supportsWith(undefined);
    expect(supports("1.2")).toBe(false);
  });

  it("answers false rather than throwing when the host raises", async () => {
    const supports = await supportsWith({
      requirements: {
        isSetSupported: () => {
          throw new Error("no");
        },
      },
    });
    expect(supports("1.2")).toBe(false);
  });

  it("is what ready() asks, so the pane's first sentence is the floor's", async () => {
    vi.resetModules();
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: { isSetSupported: (_n: string, v: string) => v === "1.2" } },
    };
    const mod = await import("../src/office/powerpoint.js");
    expect(mod.ready().ok).toBe(true);
    (globalThis as unknown as { Office: unknown }).Office = { context: {} };
    vi.resetModules();
    const again = await import("../src/office/powerpoint.js");
    const refused = again.ready();
    expect(refused.ok).toBe(false);
    expect(refused.detail).toContain("1.2");
  });
});
