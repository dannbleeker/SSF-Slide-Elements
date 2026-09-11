/**
 * A file with its PROSE removed, so a guard reads what the file DOES.
 *
 * Three guards in a sibling repo went wrong the same way, each in a different
 * syntax, each found only when it went red on a file that was perfectly correct
 * or green on one that was not:
 *
 * - `test/architecture.test.ts` forbade Office.js in the engine, and matched the
 *   word "Office.js" in the paragraphs explaining WHY the engine avoids it —
 *   four correct files, red.
 * - `scripts/manifest-rules.mjs` forbade a `<Requirements>` block, and matched
 *   the XML comment explaining why the manifest has none — so the generator
 *   refused to write a file that was exactly right.
 * - `test/release.test.ts` checked that the pre-flight runs before the tag is
 *   created, and matched the YAML comment mentioning `gh release create` — so
 *   it compared the comment's position with the check's and reported the order
 *   backwards.
 *
 * A file that explains itself is not a defect; a guard that cannot tell an
 * explanation from an instruction is. Three strippers in three files is three
 * chances to write a fourth, so they are one module — and the next person
 * reaching for one finds it here rather than inventing it. There are four now,
 * and the fourth arrived the way this module predicted: a new guard needed a
 * stripper one shade different from an existing one, and the difference
 * mattered.
 *
 * None of these is a PARSER. Each is the smallest thing that makes its own
 * guard honest, and each is checked by `test/without-prose.test.ts` against the
 * case that caught it.
 */

/** `<!-- … -->`. */
export function withoutXmlComments(text) {
  return text.replace(/<!--[\s\S]*?-->/g, "");
}

/** `#` to end of line, for YAML and anything shell-shaped. */
export function withoutHashComments(text) {
  return text
    .split("\n")
    .filter((line) => !/^\s*#/.test(line))
    .join("\n");
}

/**
 * Block comments and comment-only lines, for TypeScript and JavaScript alike.
 *
 * The literals STAY, which is the whole reason this is separate from
 * `withoutTsProse`. A guard over code that builds its output with template
 * literals — a sibling script read half its output inside one — would
 * otherwise have every one of those reads stripped out from under it and report
 * the file as not doing what it plainly does. Reach for this when the thing
 * being checked can legitimately live in a string, and for `withoutTsProse`
 * when a name in a string is prose about a thing rather than a use of it.
 */
export function withoutTsComments(text) {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split("\n")
    .filter((line) => !/^\s*(\/\/|\*)/.test(line))
    .join("\n");
}

/**
 * One template literal's INTERPOLATIONS, with its text dropped.
 *
 * `` `${size} library, scaled to ${nameOfRatio(ratio)} slides.` `` answers
 * `size nameOfRatio(ratio)`: the prose goes, the calls stay.
 *
 * Brace depth is counted rather than matched with a regex, because an
 * interpolation routinely contains braces of its own — `${exact ? {} : x}` is
 * ordinary, and a pattern stopping at the first `}` would cut the expression in
 * half and lose whatever came after it.
 *
 * A template literal NESTED inside an interpolation keeps its text. That is the
 * lenient direction for every guard that uses this — a name is reported as used
 * rather than as dead — and unpicking it would mean a parser.
 */
function interpolationsOf(literal) {
  const out = [];
  let at = 1; // past the opening backtick
  while (at < literal.length - 1) {
    if (literal[at] === "\\") {
      at += 2;
      continue;
    }
    if (literal[at] === "$" && literal[at + 1] === "{") {
      let depth = 1;
      let end = at + 2;
      while (end < literal.length && depth > 0) {
        if (literal[end] === "{") depth += 1;
        else if (literal[end] === "}") depth -= 1;
        end += 1;
      }
      out.push(literal.slice(at + 2, end - 1));
      at = end;
      continue;
    }
    at += 1;
  }
  return out.join(" ");
}

/**
 * Comments and literal TEXT, with interpolated EXPRESSIONS kept.
 *
 * The fourth stripper, and the one `withoutTsProse` cannot be: it blanks a
 * template literal whole, and a call written inside one is a call.
 * `src/pane/catalogue.ts` reaches `nameOfRatio` exactly once, from inside
 * `` `${size} library, scaled to ${nameOfRatio(ratio)} slides.` `` — so the
 * dead-export sweep, reading `withoutTsProse`, reported a function the pane
 * runs on every borrowed deck as reached by nothing but its test.
 *
 * Reach for this when the question is WHAT THIS FILE CALLS. Reach for
 * `withoutTsProse` when the question is what it mentions, where a name inside a
 * string is a sentence about a thing rather than a use of it.
 */
export function withoutTsText(text) {
  return (
    withoutTsComments(text)
      .replace(/`(?:[^`\\]|\\.)*`/g, (literal) => ` ${interpolationsOf(literal)} `)
      .replace(/"(?:[^"\\\n]|\\.)*"/g, '""')
      .replace(/'(?:[^'\\\n]|\\.)*'/g, '""')
      // Trailing comments, for the reason `withoutTsProse` spells out below.
      .replace(/(?<!\\)\/\/[^\n]*/g, "")
  );
}

/**
 * Block comments, comment-only lines, and STRING LITERALS, for TypeScript.
 *
 * The literals go too, and for the same reason the comments do: a verdict that
 * names `office-js#6105` in its text is a sentence about an issue, not a
 * dependency on one. An import specifier is a string literal as well, so a
 * guard that reads IMPORTS must read the raw source instead —
 * `test/architecture.test.ts` does exactly that, deliberately, and says so.
 */
export function withoutTsProse(text) {
  return (
    withoutTsComments(text)
      .replace(/`(?:[^`\\]|\\.)*`/g, '""')
      .replace(/"(?:[^"\\\n]|\\.)*"/g, '""')
      .replace(/'(?:[^'\\\n]|\\.)*'/g, '""')
      // A TRAILING comment, last and only here.
      //
      // `withoutTsComments` drops block comments and lines that BEGIN with `//`
      // or `*`, and stops — so `const x = 1; // never calls PowerPoint.run`
      // survived, and the architecture guard went red on a file that was
      // entirely correct. That is the exact failure this module exists to
      // prevent, in the one position it did not cover.
      //
      // It cannot go in `withoutTsComments`, which keeps literals on purpose: a
      // URL in a string is full of `//`. Here the literals are already blanked,
      // so a `//` that is left is nearly always a comment.
      //
      // NEARLY: a REGEX literal is not blanked, and one that ENDS in an escaped
      // slash puts two together — `/https:\/\//` is `…\` `/` `/`, where the
      // second slash is the terminator. The lookbehind is what separates them,
      // and it was written after getting this wrong: the first version claimed
      // an escaped slash could never be adjacent to another and truncated that
      // very pattern.
      .replace(/(?<!\\)\/\/[^\n]*/g, "")
  );
}
