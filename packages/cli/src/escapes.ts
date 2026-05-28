/**
 * Decode backslash escapes in free-text args the way `printf` does.
 *
 * Common case we're protecting against: an agent calls
 *
 *   outlook mail draft --body "Hi\n\nfoo"
 *
 * Bash double-quoting does not interpret backslash escapes, so the CLI
 * receives the literal sequence `\` + `n`. Without this helper, those
 * characters reach Graph verbatim and the rendered email shows them as-is.
 *
 * Decodes only `\n`, `\r`, `\t`, `\\`. Any other backslash sequence is left
 * intact. Real newlines, tabs, etc. already present in the input pass through
 * untouched because we only react to a literal backslash + escape char.
 *
 * Mirrors `interpret_escapes` in the original Python CLI.
 */
export function decodeEscapes(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) {
      const nxt = s[i + 1];
      if (nxt === 'n') {
        out += '\n';
        i++;
        continue;
      }
      if (nxt === 'r') {
        out += '\r';
        i++;
        continue;
      }
      if (nxt === 't') {
        out += '\t';
        i++;
        continue;
      }
      if (nxt === '\\') {
        out += '\\';
        i++;
        continue;
      }
    }
    out += ch;
  }
  return out;
}
