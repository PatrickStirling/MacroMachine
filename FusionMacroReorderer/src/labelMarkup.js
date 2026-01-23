const NAMED_COLORS = {
  red: '#ff0000',
  green: '#00ff00',
  blue: '#0000ff',
  yellow: '#ffff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  white: '#ffffff',
  black: '#000000',
  gray: '#808080',
  grey: '#808080',
  orange: '#ffa500',
  purple: '#800080',
  pink: '#ff69b4',
};

function clampByte(value) {
  const num = Number.parseInt(value, 10);
  if (!Number.isFinite(num)) return null;
  return Math.min(255, Math.max(0, num));
}

function toHex(value) {
  return value.toString(16).padStart(2, '0');
}

function hexToRgbString(hex) {
  const clean = String(hex || '').trim().toLowerCase();
  if (!/^#[0-9a-f]{6}$/.test(clean)) return null;
  const r = parseInt(clean.slice(1, 3), 16);
  const g = parseInt(clean.slice(3, 5), 16);
  const b = parseInt(clean.slice(5, 7), 16);
  return `rgb(${r},${g},${b})`;
}

function normalizeColor(value) {
  if (!value) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  if (raw[0] === '#') {
    const hex = raw.slice(1).trim();
    if (/^[0-9a-f]{3}$/i.test(hex)) {
      const r = hex[0];
      const g = hex[1];
      const b = hex[2];
      return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
    }
    if (/^[0-9a-f]{6}$/i.test(hex)) {
      return `#${hex}`.toLowerCase();
    }
    return null;
  }
  const rgbMatch = raw.match(/rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/i);
  if (rgbMatch) {
    const r = clampByte(rgbMatch[1]);
    const g = clampByte(rgbMatch[2]);
    const b = clampByte(rgbMatch[3]);
    if (r == null || g == null || b == null) return null;
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
  }
  const named = NAMED_COLORS[raw.toLowerCase()];
  return named || null;
}

function decodeEntities(text) {
  return String(text)
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&amp;/gi, '&');
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function normalizeLabelStyle(style = {}) {
  const base = (style && typeof style === 'object') ? style : {};
  return {
    bold: !!base.bold,
    italic: !!base.italic,
    underline: !!base.underline,
    center: !!base.center,
    color: normalizeColor(base.color),
  };
}

export function labelStyleEquals(a, b) {
  const left = normalizeLabelStyle(a);
  const right = normalizeLabelStyle(b);
  return left.bold === right.bold
    && left.italic === right.italic
    && left.underline === right.underline
    && left.center === right.center
    && left.color === right.color;
}

export function stripLabelMarkup(markup) {
  if (markup == null) return '';
  let out = String(markup);
  out = out.replace(/<br\s*\/?>/gi, ' ');
  out = out.replace(/<\/p\s*>/gi, ' ');
  out = out.replace(/<p\b[^>]*>/gi, ' ');
  out = out.replace(/<[^>]*>/g, '');
  out = decodeEntities(out);
  out = out.replace(/\s+/g, ' ').trim();
  return out;
}

export function parseLabelMarkup(markup) {
  const raw = markup == null ? '' : String(markup);
  const lower = raw.toLowerCase();
  const style = {
    bold: /<\s*(b|strong)\b/.test(lower),
    italic: /<\s*(i|em)\b/.test(lower),
    underline: /<\s*u\b/.test(lower),
    center: /<\s*center\b/.test(lower) || /text-align\s*:\s*center/.test(lower),
    color: null,
  };
  let colorMatch = raw.match(/<\s*font\b[^>]*\bcolor\s*=\s*['"]?([^'">\s]+)['"]?/i);
  if (!colorMatch) {
    const styleMatch = raw.match(/style\s*=\s*['"]([^'"]+)['"]/i);
    if (styleMatch) {
      const styleText = styleMatch[1];
      const colorInStyle = styleText.match(/color\s*:\s*([^;]+)\s*/i);
      if (colorInStyle) colorMatch = colorInStyle;
    }
  }
  if (colorMatch) {
    style.color = normalizeColor(colorMatch[1]);
  }
  return {
    text: stripLabelMarkup(raw),
    style: normalizeLabelStyle(style),
  };
}

export function buildLabelMarkup(text, style) {
  const safeText = escapeHtml(text || '');
  const normalized = normalizeLabelStyle(style);
  let out = safeText;
  if (normalized.bold) out = `<b>${out}</b>`;
  if (normalized.italic) out = `<i>${out}</i>`;
  if (normalized.underline) out = `<u>${out}</u>`;
  if (normalized.color) {
    const rgb = hexToRgbString(normalized.color) || normalized.color;
    out = `<p style='color:${rgb};'>${out}</p>`;
  }
  if (normalized.center) out = `<center>${out}</center>`;
  return out;
}
