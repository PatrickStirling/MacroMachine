import {
  getLineIndent,
  findMatchingBrace,
  isSpace,
  isIdentStart,
  isIdentPart,
  humanizeName,
  isQuoteEscaped,
} from './textUtils.js';
import { parseLabelMarkup, normalizeLabelStyle } from './labelMarkup.js';

export function parseSetting(text) {
  const dataLink = extractFmrDataLink(text);
  let blocks = findAllInputsBlocksWithInstanceInputs(text);
  let preferGroupInputs = false;
  if (!blocks.length) {
    preferGroupInputs = true;
    const fb = findInputsBlockAnywhere(text) || findInputsBlockByBacktrack(text);
    if (fb) blocks = [fb];
  }
  if (!blocks.length) {
    const all = findAllInputsBlocks(text);
    if (all && all.length) {
      preferGroupInputs = true;
      blocks = all;
    }
  }
  if (!blocks.length) throw new Error('Could not find an Inputs block with InstanceInput entries within any macro.');
  let chosen = null;
  for (const b of blocks) {
    const grp = findEnclosingGroupForIndex(text, b.openIndex);
    if (grp) { chosen = { ...b, macroName: grp.name, operatorType: grp.operatorType || 'GroupOperator', groupOpen: grp.groupOpenIndex, groupClose: grp.groupCloseIndex }; break; }
  }
  if (!chosen) chosen = { ...blocks[0], macroName: 'Unknown', operatorType: 'GroupOperator' };
  if (!chosen.macroName || chosen.macroName === 'Unknown') {
    try {
      const reGroup = /(^|\n)\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{/g;
      const m = reGroup.exec(text);
      if (m && m[2]) {
        chosen.macroName = m[2];
        if (m[3]) chosen.operatorType = m[3];
      }
    } catch (_) {}
  }
  if (chosen.groupOpen == null || chosen.groupClose == null) {
    const best = findGroupBoundsByName(text, chosen.macroName || chosen.macroNameOriginal || null) || findFirstGroupBounds(text);
    if (best) {
      chosen.groupOpen = best.groupOpenIndex;
      chosen.groupClose = best.groupCloseIndex;
    }
  }
  if (preferGroupInputs && chosen.groupOpen != null && chosen.groupClose != null) {
    const macroInputs = findGroupLevelInputsBlock(text, chosen.groupOpen, chosen.groupClose);
    if (macroInputs) {
      chosen.inputsHeaderStart = macroInputs.headerStart;
      chosen.openIndex = macroInputs.openIndex;
      chosen.closeIndex = macroInputs.closeIndex;
    }
  }
  const indent = getLineIndent(text, chosen.inputsHeaderStart);
  const entries = parsePublishedControls(text, chosen.openIndex, chosen.closeIndex);
  const labelMap = chosen.groupOpen != null ? parseLabelsInGroup(text, chosen.groupOpen, chosen.groupClose) : new Map();
  const displayMap = chosen.groupOpen != null ? parseControlDisplayNamesInGroup(text, chosen.groupOpen, chosen.groupClose) : new Map();
  for (const e of entries) {
    if (isStructuralEntryName(e.key)) e.locked = true;
    if (!e.sourceOp || !e.source) continue;
    const key = `${e.sourceOp}.${e.source}`;
    const meta = labelMap.get(key);
    const dispRaw = displayMap.get(key);
    if (meta) { e.isLabel = true; e.labelCount = meta.numInputs || 0; }
    if (e.isLabel) {
      let labelRaw = dispRaw;
      if (!labelRaw && e.name && /<[^>]+>/.test(e.name)) labelRaw = e.name;
      if (labelRaw) {
        const parsed = parseLabelMarkup(labelRaw);
        e.displayName = parsed.text || labelRaw;
        e.labelStyle = normalizeLabelStyle(parsed.style);
        e.labelStyleOriginal = { ...e.labelStyle };
      } else if (!e.displayName) {
        e.displayName = e.name || dispRaw || humanizeName(e.source);
      }
      if (!e.labelStyle) {
        e.labelStyle = normalizeLabelStyle(null);
        e.labelStyleOriginal = { ...e.labelStyle };
      }
    } else if (!e.name) {
      const disp = dispRaw || humanizeName(e.source);
      if (disp) e.displayName = disp;
    }
    const fallback = e.displayName || e.name || `${e.sourceOp || ''}${e.source ? '.' + e.source : ''}`;
    e.displayName = e.displayName || fallback;
    e.displayNameOriginal = e.displayName;
  }
  const order = entries.map((_, i) => i);
  return {
    macroName: chosen.macroName,
    macroNameOriginal: chosen.macroName,
    operatorType: chosen.operatorType || 'GroupOperator',
    operatorTypeOriginal: chosen.operatorType || 'GroupOperator',
    inputs: { inputsHeaderStart: chosen.inputsHeaderStart, openIndex: chosen.openIndex, closeIndex: chosen.closeIndex, innerStart: chosen.openIndex + 1, innerEnd: chosen.closeIndex, indent },
    entries,
    order,
    originalOrder: [...order],
    dataLink,
  };
}

function extractFmrDataLink(text) {
  try {
    if (!text) return null;
    const re = /--\\s*FMR_DATA_LINK_BEGIN\\s*\\n--\\s*([\\s\\S]*?)\\n--\\s*FMR_DATA_LINK_END/;
    const m = text.match(re);
    if (!m || !m[1]) return null;
    const json = m[1].replace(/^\\s*--\\s?/gm, '').trim();
    if (!json) return null;
    return JSON.parse(json);
  } catch (_) {
    return null;
  }
}

export function findEnclosingGroupForIndex(text, index) {
  let searchEnd = index;
  while (searchEnd >= 0) {
    let bestPos = -1;
    let opType = null;
    const candidates = ['GroupOperator', 'MacroOperator'];
    for (const cand of candidates) {
      const pos = text.lastIndexOf(cand, searchEnd);
      if (pos > bestPos) { bestPos = pos; opType = cand; }
    }
    if (bestPos < 0) break;
    let j = bestPos + opType.length;
    while (j < text.length && isSpace(text[j])) j++;
    if (text[j] !== '{') { searchEnd = bestPos - 1; continue; }
    const groupOpenIndex = j;
    const groupCloseIndex = findMatchingBrace(text, groupOpenIndex);
    if (groupCloseIndex > groupOpenIndex && index > groupOpenIndex && index < groupCloseIndex) {
      let left = bestPos - 1;
      while (left >= 0 && isSpace(text[left])) left--;
      if (text[left] !== '=') return { name: 'Unknown', groupOpenIndex, groupCloseIndex };
      left--;
      while (left >= 0 && isSpace(text[left])) left--;
      let nameEnd = left + 1;
      while (left >= 0 && isIdentPart(text[left])) left--;
      const name = text.slice(left + 1, nameEnd).trim();
      return { name: name || 'Unknown', operatorType: opType, groupOpenIndex, groupCloseIndex };
    }
    searchEnd = bestPos - 1;
  }
  return null;
}

export function extractQuotedProp(text, prop) {
  const re = new RegExp(prop.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*=\\s*\"([^\"]*)\"');
  const m = text.match(re);
  return m ? m[1] : null;
}

export function extractNumericProp(text, prop) {
  const re = new RegExp(prop.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*=\\s*([0-9]+)');
  const m = text.match(re);
  return m ? parseInt(m[1], 10) : null;
}

export function findAllInputsBlocksWithInstanceInputs(text) {
  const results = [];
  const re = new RegExp('Inputs\\s*=\\s*(?:ordered\\(\\)\\s*)?\\{', 'gi');
  let m;
  while ((m = re.exec(text)) !== null) {
    const matchStart = m.index;
    const braceRel = m[0].lastIndexOf('{');
    const openIndex = matchStart + braceRel;
    const closeIndex = findMatchingBrace(text, openIndex);
    if (closeIndex < 0) continue;
    const inner = text.slice(openIndex + 1, closeIndex);
    if (inner.includes('InstanceInput')) {
      results.push({ inputsHeaderStart: matchStart, openIndex, closeIndex });
    }
    if (re.lastIndex === m.index) re.lastIndex++;
  }
  return results;
}

export function findInputsBlockAnywhere(text) {
  let i = 0;
  let inStr = false;
  while (i < text.length) {
    const ch = text[i];
    if (inStr) { if (ch === '"' && !isQuoteEscaped(text, i)) inStr = false; i++; continue; }
    if (ch === '"') { inStr = true; i++; continue; }
    if (text.slice(i, i + 6).toLowerCase() === 'inputs') {
      let j = i + 6;
      while (j < text.length && isSpace(text[j])) j++;
      if (text[j] !== '=') { i++; continue; }
      j++;
      while (j < text.length && isSpace(text[j])) j++;
      const maybe = text.slice(j, j + 8).toLowerCase();
      if (maybe === 'ordered(') {
        j += 8;
        while (j < text.length && isSpace(text[j])) j++;
        if (text[j] !== ')') { i++; continue; }
        j++;
        while (j < text.length && isSpace(text[j])) j++;
      }
      if (text[j] === '{') {
        const openIndex = j;
        const closeIndex = findMatchingBrace(text, openIndex);
        if (closeIndex < 0) break;
        return { inputsHeaderStart: i, openIndex, closeIndex };
      }
    }
    i++;
  }
  return null;
}

export function findInputsBlockByBacktrack(text) {
  const re = /InstanceInput\b/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const instIdx = m.index;
    let searchFrom = instIdx;
    while (searchFrom >= 0) {
      const inputsIdx = text.lastIndexOf('Inputs', searchFrom);
      if (inputsIdx < 0) break;
      const chBefore = text[inputsIdx - 1] || ' ';
      const chAfter = text[inputsIdx + 6] || ' ';
      if (/[^A-Za-z0-9_]/.test(chBefore) && /[^A-Za-z0-9_]/.test(chAfter)) {
        let j = inputsIdx + 6;
        while (j < text.length && isSpace(text[j])) j++;
        if (text[j] !== '=') { searchFrom = inputsIdx - 1; continue; }
        j++;
        while (j < text.length && isSpace(text[j])) j++;
        const maybe = text.slice(j, j + 7).toLowerCase();
        if (maybe.startsWith('ordered')) {
          while (j < text.length && text[j] !== ')') j++;
          if (text[j] === ')') j++;
          while (j < text.length && isSpace(text[j])) j++;
        }
        if (text[j] === '{') {
          const openIndex = j;
          const closeIndex = findMatchingBrace(text, openIndex);
          if (closeIndex > openIndex && instIdx > openIndex && instIdx < closeIndex) {
            return { inputsHeaderStart: inputsIdx, openIndex, closeIndex };
          }
        }
      }
      searchFrom = inputsIdx - 1;
    }
  }
  return null;
}

function findAllInputsBlocks(text) {
  const results = [];
  const re = new RegExp('Inputs\\s*=\\s*(?:ordered\\(\\)\\s*)?\\{', 'gi');
  let m;
  while ((m = re.exec(text)) !== null) {
    const openIndex = m.index + m[0].lastIndexOf('{');
    const closeIndex = findMatchingBrace(text, openIndex);
    if (closeIndex < 0) continue;
    results.push({ inputsHeaderStart: m.index, openIndex, closeIndex });
    if (re.lastIndex === m.index) re.lastIndex++;
  }
  return results;
}

function findGroupBoundsByName(text, macroName) {
  try {
    if (!macroName) return null;
    const re = new RegExp(`(^|\\n)\\s*${macroName.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}\\s*=\\s*(GroupOperator|MacroOperator)\\s*\\{`);
    const match = re.exec(text);
    if (!match) return null;
    const openIndex = match.index + match[0].lastIndexOf('{');
    const closeIndex = findMatchingBrace(text, openIndex);
    if (closeIndex < 0) return null;
    return { groupOpenIndex: openIndex, groupCloseIndex: closeIndex };
  } catch (_) {
    return null;
  }
}

function findFirstGroupBounds(text) {
  try {
    const re = /([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{/g;
    const match = re.exec(text);
    if (!match) return null;
    const openIndex = match.index + match[0].lastIndexOf('{');
    const closeIndex = findMatchingBrace(text, openIndex);
    if (closeIndex < 0) return null;
    return { groupOpenIndex: openIndex, groupCloseIndex: closeIndex };
  } catch (_) {
    return null;
  }
}

function findGroupLevelInputsBlock(text, groupOpen, groupClose) {
  try {
    if (groupOpen == null || groupClose == null) return null;
    const toolsPos = text.indexOf('Tools', groupOpen);
    const limit = (toolsPos >= 0 && toolsPos < groupClose) ? toolsPos : groupClose;
    if (limit <= groupOpen) return null;
    const segment = text.slice(groupOpen, limit);
    const re = /Inputs\s*=\s*(?:ordered\(\)\s*)?\{/g;
    const match = re.exec(segment);
    if (!match) return null;
    const headerStart = groupOpen + match.index;
    const openIndex = groupOpen + match.index + match[0].lastIndexOf('{');
    const closeIndex = findMatchingBrace(text, openIndex);
    if (closeIndex < 0 || closeIndex > groupClose) return null;
    return { headerStart, openIndex, closeIndex };
  } catch (_) {
    return null;
  }
}

function parseLabelsInGroup(text, groupOpen, groupClose) {
  const map = new Map();
  const toolsPos = text.indexOf('Tools = ordered()', groupOpen);
  if (toolsPos < 0 || toolsPos > groupClose) return map;
  const open = text.indexOf('{', toolsPos);
  if (open < 0) return map;
  const close = findMatchingBrace(text, open);
  if (close < 0 || close > groupClose) return map;
  const inner = text.slice(open + 1, close);
  let i = 0, depth = 0, inStr = false;
  while (i < inner.length) {
    const ch = inner[i];
    if (inStr) { if (ch === '"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
    if (ch === '"') { inStr = true; i++; continue; }
    if (ch === '{') { depth++; i++; continue; }
    if (ch === '}') { depth--; i++; continue; }
    if (depth === 0 && isIdentStart(ch)) {
      const nameStart = i; i++;
      while (i < inner.length && isIdentPart(inner[i])) i++;
      const toolName = inner.slice(nameStart, i);
      while (i < inner.length && isSpace(inner[i])) i++;
      if (inner[i] !== '=') { i++; continue; }
      i++;
      while (i < inner.length && isSpace(inner[i])) i++;
      while (i < inner.length && isIdentPart(inner[i])) i++;
      while (i < inner.length && isSpace(inner[i])) i++;
      if (inner[i] !== '{') { i++; continue; }
      const tOpen = i;
      const tClose = findMatchingBrace(inner, tOpen);
      if (tClose < 0) break;
      const body = inner.slice(tOpen + 1, tClose);
      const ucPos = body.indexOf('UserControls = ordered()');
      if (ucPos >= 0) {
        const ucOpen = body.indexOf('{', ucPos);
        const ucClose = ucOpen >= 0 ? findMatchingBrace(body, ucOpen) : -1;
        if (ucOpen >= 0 && ucClose > ucOpen) {
          const uc = body.slice(ucOpen + 1, ucClose);
          let j = 0, d = 0, s = false;
          while (j < uc.length) {
            const c = uc[j];
            if (s) { if (c === '"' && !isQuoteEscaped(uc, j)) s = false; j++; continue; }
            if (c === '"') { s = true; j++; continue; }
            if (c === '{') { d++; j++; continue; }
            if (c === '}') { d--; j++; continue; }
            if (d === 0 && isIdentStart(c)) {
              const cnStart = j; j++;
              while (j < uc.length && isIdentPart(uc[j])) j++;
              const controlName = uc.slice(cnStart, j);
              while (j < uc.length && isSpace(uc[j])) j++;
              if (uc[j] !== '=') { j++; continue; }
              j++;
              while (j < uc.length && isSpace(uc[j])) j++;
              if (uc[j] !== '{') { j++; continue; }
              const cOpen = j;
              const cClose = findMatchingBrace(uc, cOpen);
              if (cClose < 0) break;
              const cBody = uc.slice(cOpen + 1, cClose);
              if (/INPID_InputControl\s*=\s*"LabelControl"/.test(cBody)) {
                const num = extractNumericProp(cBody, 'LBLC_NumInputs');
                const disp = extractQuotedProp(cBody, 'LINKS_Name');
                map.set(`${toolName}.${controlName}`, { numInputs: num || 0, name: disp || controlName });
              }
              j = cClose + 1;
              continue;
            }
            j++;
          }
        }
      }
      i = tClose + 1;
      continue;
    }
    i++;
  }
  return map;
}

function parseControlDisplayNamesInGroup(text, groupOpen, groupClose) {
  const map = new Map();
  const toolsPos = text.indexOf('Tools = ordered()', groupOpen);
  if (toolsPos < 0 || toolsPos > groupClose) return map;
  const open = text.indexOf('{', toolsPos);
  if (open < 0) return map;
  const close = findMatchingBrace(text, open);
  if (close < 0 || close > groupClose) return map;
  const inner = text.slice(open + 1, close);
  let i = 0, depth = 0, inStr = false;
  while (i < inner.length) {
    const ch = inner[i];
    if (inStr) { if (ch === '"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
    if (ch === '"') { inStr = true; i++; continue; }
    if (ch === '{') { depth++; i++; continue; }
    if (ch === '}') { depth--; i++; continue; }
    if (depth === 0 && isIdentStart(ch)) {
      const nameStart = i; i++;
      while (i < inner.length && isIdentPart(inner[i])) i++;
      const toolName = inner.slice(nameStart, i);
      while (i < inner.length && isSpace(inner[i])) i++;
      if (inner[i] !== '=') { i++; continue; }
      i++;
      while (i < inner.length && isSpace(inner[i])) i++;
      while (i < inner.length && isIdentPart(inner[i])) i++;
      while (i < inner.length && isSpace(inner[i])) i++;
      if (inner[i] !== '{') { i++; continue; }
      const tOpen = i;
      const tClose = findMatchingBrace(inner, tOpen);
      if (tClose < 0) break;
      const body = inner.slice(tOpen + 1, tClose);
      const ucPos = body.indexOf('UserControls = ordered()');
      if (ucPos >= 0) {
        const ucOpen = body.indexOf('{', ucPos);
        const ucClose = ucOpen >= 0 ? findMatchingBrace(body, ucOpen) : -1;
        if (ucOpen >= 0 && ucClose > ucOpen) {
          const uc = body.slice(ucOpen + 1, ucClose);
          let j = 0, d = 0, s = false;
          while (j < uc.length) {
            const c = uc[j];
            if (s) { if (c === '"' && !isQuoteEscaped(uc, j)) s = false; j++; continue; }
            if (c === '"') { s = true; j++; continue; }
            if (c === '{') { d++; j++; continue; }
            if (c === '}') { d--; j++; continue; }
            if (d === 0 && isIdentStart(c)) {
              const cnStart = j; j++;
              while (j < uc.length && isIdentPart(uc[j])) j++;
              const controlName = uc.slice(cnStart, j);
              while (j < uc.length && isSpace(uc[j])) j++;
              if (uc[j] !== '=') { j++; continue; }
              j++;
              while (j < uc.length && isSpace(uc[j])) j++;
              if (uc[j] !== '{') { j++; continue; }
              const cOpen = j;
              const cClose = findMatchingBrace(uc, cOpen);
              if (cClose < 0) break;
              const cBody = uc.slice(cOpen + 1, cClose);
              const disp = extractQuotedProp(cBody, 'LINKS_Name');
              if (disp) map.set(`${toolName}.${controlName}`, disp);
              j = cClose + 1;
              continue;
            }
            j++;
          }
        }
      }
      i = tClose + 1;
      continue;
    }
    i++;
  }
  return map;
}

function parsePublishedControls(text, inputsOpen, inputsClose) {
  const inner = text.slice(inputsOpen + 1, inputsClose);
  const entries = [];
  let i = 0;
  let depth = 0;
  let inStr = false;
  while (i < inner.length) {
    const ch = inner[i];
    if (inStr) {
      if (ch === '"' && !isQuoteEscaped(inner, i)) inStr = false;
      i++;
      continue;
    }
    if (ch === '"') { inStr = true; i++; continue; }
    if (ch === '{') { depth++; i++; continue; }
    if (ch === '}') { depth--; i++; continue; }
    if (depth === 0) {
      if (isIdentStart(ch)) {
        const keyStart = i;
        i++;
        while (i < inner.length && isIdentPart(inner[i])) i++;
        const key = inner.slice(keyStart, i);
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner[i] !== '=') continue;
        i++;
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner.slice(i, i + 13) !== 'InstanceInput') continue;
        i += 13;
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner[i] !== '{') continue;
        const instOpen = i;
        const instClose = findMatchingBrace(inner, instOpen);
        if (instClose < 0) throw new Error('Malformed InstanceInput block.');
        let end = instClose + 1;
        if (inner[end] === ',') end++;
        const raw = inner.slice(keyStart, end);
        const propStr = inner.slice(instOpen + 1, instClose);
        const name = extractQuotedProp(propStr, 'Name');
        const page = extractQuotedProp(propStr, 'Page');
        const sourceOp = extractQuotedProp(propStr, 'SourceOp');
        const source = extractQuotedProp(propStr, 'Source');
        const displayName = name || (sourceOp && source ? `${sourceOp}.${source}` : (source || key));
        const controlGroup = extractNumericProp(propStr, 'ControlGroup');
        entries.push({ key, name: name || null, page: page || null, sourceOp: sourceOp || null, source: source || null, displayName, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), raw });
        i = end;
        continue;
      }
    }
    i++;
  }
  return entries;
}

function isStructuralEntryName(key) {
  if (!key) return false;
  const k = String(key).toLowerCase();
  return /^maininput\d+$/.test(k) || /^main\d+input$/.test(k);
}
