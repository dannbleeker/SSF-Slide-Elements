/**
 * `btoa`/`atob` on the global.
 *
 * The engine uses them so ONE implementation runs in the pane and in the suite
 * — a Node-only `Buffer` path would be code the pane never executes, which is
 * the shape of bug that only shows up in a real PowerPoint.
 */
if (typeof globalThis.btoa !== "function") {
  globalThis.btoa = (s: string): string => Buffer.from(s, "binary").toString("base64");
}
if (typeof globalThis.atob !== "function") {
  globalThis.atob = (s: string): string => Buffer.from(s, "base64").toString("binary");
}
export {};
