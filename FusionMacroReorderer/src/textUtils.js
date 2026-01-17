export function isSpace(ch) {
  return ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n';
}

export function isIdentStart(ch) {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '_' || ch === '[' || ch === ']';
}

export function isIdentPart(ch) {
  return isIdentStart(ch) || (ch >= '0' && ch <= '9') || ch === '.';
}

export function getLineIndent(text, pos) {
  const lineStart = text.lastIndexOf('\n', pos) + 1;
  const line = text.slice(lineStart, pos);
  const m = line.match(/^[\t ]*/);
  return m ? m[0] : '';
}

export function isQuoteEscaped(text, index) {
  let backslashes = 0;
  for (let i = index - 1; i >= 0 && text[i] === '\\'; i--) backslashes++;
  return (backslashes % 2) === 1;
}

export function findMatchingBrace(text, openIndex) {
  let depth = 0;
  let inStr = false;
  for (let i = openIndex; i < text.length; i++) {
    const ch = text[i];
    if (inStr) {
      if (ch === '"' && !isQuoteEscaped(text, i)) inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; continue; }
    if (ch === '{') { depth++; continue; }
    if (ch === '}') { depth--; if (depth === 0) return i; continue; }
  }
  return -1;
}

export function humanizeName(s) {
  if (!s) return s;
  let out = String(s)
    .replace(/_/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Za-z])([0-9])/g, '$1 $2')
    .replace(/([0-9])([A-Za-z])/g, '$1 $2')
    .trim();
  out = out.replace(/\b([a-z])/g, (m, a) => a.toUpperCase());
  return out;
}

export function escapeQuotes(s) {
  return String(s).replace(/"/g, '\\"');
}
