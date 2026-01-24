const LUA_KEYWORDS = new Set([
  'and', 'break', 'do', 'else', 'elseif', 'end', 'false', 'for', 'function',
  'if', 'in', 'local', 'nil', 'not', 'or', 'repeat', 'return', 'then', 'true',
  'until', 'while',
]);

const LUA_BUILTINS = new Set(['true', 'false', 'nil']);

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isIdentStart(ch) {
  return (ch >= 'A' && ch <= 'Z')
    || (ch >= 'a' && ch <= 'z')
    || ch === '_';
}

function isIdentPart(ch) {
  return isIdentStart(ch) || (ch >= '0' && ch <= '9');
}

function readWhile(text, start, predicate) {
  let i = start;
  while (i < text.length && predicate(text[i])) i++;
  return i;
}

function normalizeNameSet(names) {
  if (!names) return null;
  if (names instanceof Set) return names;
  if (Array.isArray(names)) return new Set(names);
  return null;
}

export function highlightLua(source, options = {}) {
  if (!source) return '';
  const toolNames = normalizeNameSet(options.toolNames);
  const controlNames = normalizeNameSet(options.controlNames);
  const tokens = [];
  let i = 0;
  while (i < source.length) {
    const ch = source[i];
    if (ch === '-' && source[i + 1] === '-') {
      if (source[i + 2] === '[' && source[i + 3] === '[') {
        const end = source.indexOf(']]', i + 4);
        const stop = end >= 0 ? end + 2 : source.length;
        tokens.push({ type: 'comment', value: source.slice(i, stop) });
        i = stop;
        continue;
      }
      const lineEnd = source.indexOf('\n', i + 2);
      const stop = lineEnd >= 0 ? lineEnd : source.length;
      tokens.push({ type: 'comment', value: source.slice(i, stop) });
      i = stop;
      continue;
    }
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < source.length) {
        const cj = source[j];
        if (cj === '\\') { j += 2; continue; }
        if (cj === ch) { j++; break; }
        j++;
      }
      tokens.push({ type: 'string', value: source.slice(i, j) });
      i = j;
      continue;
    }
    if (ch === '[' && source[i + 1] === '[') {
      const end = source.indexOf(']]', i + 2);
      const stop = end >= 0 ? end + 2 : source.length;
      tokens.push({ type: 'string', value: source.slice(i, stop) });
      i = stop;
      continue;
    }
    if (ch >= '0' && ch <= '9') {
      const slice = source.slice(i);
      const m = slice.match(/^(0x[0-9a-fA-F]+|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)/);
      if (m) {
        tokens.push({ type: 'number', value: m[0] });
        i += m[0].length;
        continue;
      }
    }
    if (isIdentStart(ch)) {
      const end = readWhile(source, i + 1, isIdentPart);
      const word = source.slice(i, end);
      if (LUA_KEYWORDS.has(word)) {
        tokens.push({ type: LUA_BUILTINS.has(word) ? 'builtin' : 'keyword', value: word });
      } else if (toolNames && toolNames.has(word)) {
        tokens.push({ type: 'tool', value: word });
      } else if (controlNames && controlNames.has(word)) {
        tokens.push({ type: 'control', value: word });
      } else {
        tokens.push({ type: 'ident', value: word });
      }
      i = end;
      continue;
    }
    tokens.push({ type: 'text', value: ch });
    i++;
  }

  return tokens.map((tok) => {
    const escaped = escapeHtml(tok.value);
    if (tok.type === 'text' || tok.type === 'ident') return escaped;
    return `<span class="lua-token lua-${tok.type}">${escaped}</span>`;
  }).join('');
}
