/**
 * A TypeScript file with its comments and string literals removed.
 *
 * The architecture guard asks whether a file TOUCHES Office.js, and every file
 * in this repo TALKS about Office.js at length — "hand the result back to
 * PowerPoint", "Office echoes the argument back". A guard that reads prose
 * reports every well-documented file as a violation, which is a gate that has
 * to be switched off within a day.
 *
 * Deliberately not a parser. It is a small state machine over the four things
 * that can contain a `.` — line comments, block comments, quoted strings and
 * template literals — and it is used by exactly one test, which asserts against
 * files whose contents that test also controls.
 */
export function withoutProse(source) {
  let out = "";
  let i = 0;
  const n = source.length;
  while (i < n) {
    const c = source[i];
    const next = source[i + 1];
    if (c === "/" && next === "/") {
      while (i < n && source[i] !== "\n") i += 1;
      continue;
    }
    if (c === "/" && next === "*") {
      i += 2;
      while (i < n && !(source[i] === "*" && source[i + 1] === "/")) i += 1;
      i += 2;
      continue;
    }
    if (c === '"' || c === "'" || c === "`") {
      const quote = c;
      i += 1;
      while (i < n && source[i] !== quote) {
        // A backslash escapes the next character, including the quote itself.
        if (source[i] === "\\") i += 1;
        i += 1;
      }
      i += 1;
      // The string is replaced by an empty one so the code around it still
      // reads as code — dropping it entirely would join two identifiers.
      out += `${quote}${quote}`;
      continue;
    }
    out += c;
    i += 1;
  }
  return out;
}
