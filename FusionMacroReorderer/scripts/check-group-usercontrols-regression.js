#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const mmRoot = path.resolve(repoRoot, '..');

const knownInputControls = new Set([
  'buttoncontrol',
  'checkboxcontrol',
  'combocontrol',
  'colorcontrol',
  'labelcontrol',
  'multibuttoncontrol',
  'offsetcontrol',
  'screwcontrol',
  'slidercontrol',
  'texteditcontrol',
]);

const systemControlIds = new Set([
  'FMR_FileMeta',
  'FMR_DataLink',
  'FMR_PresetData',
  'FMR_PresetScript',
  'FMR_Preset',
]);

const samples = [
  {
    name: 'Published_ALL',
    file: path.join(mmRoot, 'Published_ALL.setting'),
    expectBlocks: 0,
    expectControls: 0,
    expectDuplicateIds: [],
    expectUnknownCount: 0,
    expectSystemIds: [],
  },
  {
    name: 'Published_ALL_Modifiers',
    file: path.join(mmRoot, 'Published_ALL_Modifiers.setting'),
    expectBlocks: 0,
    expectControls: 0,
    expectDuplicateIds: [],
    expectUnknownCount: 0,
    expectSystemIds: [],
  },
  {
    name: 'Published_Connect',
    file: path.join(mmRoot, 'Published_Connect.setting'),
    expectBlocks: 0,
    expectControls: 0,
    expectDuplicateIds: [],
    expectUnknownCount: 0,
    expectSystemIds: [],
  },
  {
    name: 'BG_XF',
    file: path.join(mmRoot, 'BG_XF.setting'),
    expectBlocks: 0,
    expectControls: 0,
    expectDuplicateIds: [],
    expectUnknownCount: 0,
    expectSystemIds: [],
  },
  {
    name: 'ProtoV3',
    file: path.join(mmRoot, 'Macro_ref', '_STORE', '64_ProtoV3', 'V1.01', 'Edit', 'Generators', 'Stirling Supply Co', 'ProtoV3.setting'),
    expectBlocks: 1,
    expectControls: 3,
    expectDuplicateIds: [],
    expectUnknownCount: 0,
    expectSystemIds: [],
  },
  {
    name: 'Pre_Test18',
    file: path.join(mmRoot, 'Pre_Test18.setting'),
    expectBlocks: 2,
    expectControls: 4,
    expectDuplicateIds: ['FMR_FileMeta'],
    expectUnknownCount: 0,
    expectSystemIds: ['FMR_DataLink', 'FMR_FileMeta', 'FMR_PresetData'],
  },
];

function findMatchingBrace(text, openIndex) {
  let depth = 0;
  let inString = false;
  for (let i = openIndex; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"' && !isEscapedQuote(text, i)) inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') {
      depth++;
      continue;
    }
    if (ch === '}') {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function isEscapedQuote(text, quoteIndex) {
  let backslashes = 0;
  for (let i = quoteIndex - 1; i >= 0 && text[i] === '\\'; i--) {
    backslashes++;
  }
  return (backslashes % 2) === 1;
}

function getTopGroupBounds(text) {
  const match = /([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{/.exec(text);
  if (!match) return null;
  const openIndex = text.indexOf('{', match.index);
  if (openIndex < 0) return null;
  const closeIndex = findMatchingBrace(text, openIndex);
  if (closeIndex < 0) return null;
  return {
    name: match[1],
    openIndex,
    closeIndex,
  };
}

function getGroupUserControlsBlocks(text, groupOpenIndex, groupCloseIndex) {
  const blocks = [];
  let depth = 0;
  let inString = false;
  for (let i = groupOpenIndex + 1; i < groupCloseIndex; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"' && !isEscapedQuote(text, i)) inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') {
      depth++;
      continue;
    }
    if (ch === '}') {
      depth--;
      continue;
    }
    if (depth === 0 && text.startsWith('UserControls = ordered()', i)) {
      const braceIndex = text.indexOf('{', i);
      if (braceIndex < 0 || braceIndex > groupCloseIndex) continue;
      const closeIndex = findMatchingBrace(text, braceIndex);
      if (closeIndex < 0 || closeIndex > groupCloseIndex) continue;
      blocks.push({ openIndex: braceIndex, closeIndex });
      i = closeIndex;
    }
  }
  return blocks;
}

function getUcControls(text, ucOpenIndex, ucCloseIndex) {
  const controls = [];
  let depth = 0;
  let inString = false;
  for (let i = ucOpenIndex + 1; i < ucCloseIndex; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"' && !isEscapedQuote(text, i)) inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') {
      depth++;
      continue;
    }
    if (ch === '}') {
      depth--;
      continue;
    }
    if (depth === 0 && /[A-Za-z_]/.test(ch)) {
      const start = i;
      i++;
      while (i < ucCloseIndex && /[A-Za-z0-9_]/.test(text[i])) i++;
      const id = text.slice(start, i);
      while (i < ucCloseIndex && /\s/.test(text[i])) i++;
      if (text[i] !== '=') continue;
      i++;
      while (i < ucCloseIndex && /\s/.test(text[i])) i++;
      if (text[i] !== '{') continue;
      const openIndex = i;
      const closeIndex = findMatchingBrace(text, openIndex);
      if (closeIndex < 0 || closeIndex > ucCloseIndex) continue;
      const body = text.slice(openIndex + 1, closeIndex);
      const inputControlMatch = /INPID_InputControl\s*=\s*"([^"]*)"/.exec(body);
      controls.push({
        id,
        inputControl: inputControlMatch ? inputControlMatch[1] : '',
      });
      i = closeIndex;
    }
  }
  return controls;
}

function measureGroupUserControls(filePath) {
  const text = fs.readFileSync(filePath, 'utf8');
  const group = getTopGroupBounds(text);
  if (!group) {
    throw new Error(`No top-level GroupOperator/MacroOperator found in ${filePath}`);
  }
  const blocks = getGroupUserControlsBlocks(text, group.openIndex, group.closeIndex);
  const controls = [];
  for (const block of blocks) {
    controls.push(...getUcControls(text, block.openIndex, block.closeIndex));
  }
  const duplicateIds = Array.from(
    controls.reduce((map, control) => {
      map.set(control.id, (map.get(control.id) || 0) + 1);
      return map;
    }, new Map()).entries()
  )
    .filter(([, count]) => count > 1)
    .map(([id]) => id)
    .sort();
  const unknownControls = controls.filter((control) => {
    const inputControl = String(control.inputControl || '').trim().toLowerCase();
    return !systemControlIds.has(control.id) && !knownInputControls.has(inputControl);
  });
  const systemIds = Array.from(new Set(controls.filter((control) => systemControlIds.has(control.id)).map((control) => control.id))).sort();
  return {
    groupName: group.name,
    blocks: blocks.length,
    controls: controls.length,
    duplicateIds,
    unknownCount: unknownControls.length,
    unknownControls,
    systemIds,
    details: controls,
  };
}

function sameArray(a, b) {
  return JSON.stringify([...a].sort()) === JSON.stringify([...b].sort());
}

const verbose = process.argv.includes('--verbose');
const failures = [];
const rows = [];

for (const sample of samples) {
  if (!fs.existsSync(sample.file)) {
    failures.push(`${sample.name}: missing sample ${sample.file}`);
    continue;
  }

  const measured = measureGroupUserControls(sample.file);
  rows.push({
    Sample: sample.name,
    Group: measured.groupName,
    Blocks: measured.blocks,
    Controls: measured.controls,
    DuplicateIds: measured.duplicateIds.join(', '),
    Unknown: measured.unknownCount,
    SystemIds: measured.systemIds.join(', '),
  });

  if (measured.blocks !== sample.expectBlocks) {
    failures.push(`${sample.name}: expected blocks=${sample.expectBlocks}, got ${measured.blocks}`);
  }
  if (measured.controls !== sample.expectControls) {
    failures.push(`${sample.name}: expected controls=${sample.expectControls}, got ${measured.controls}`);
  }
  if (!sameArray(measured.duplicateIds, sample.expectDuplicateIds)) {
    failures.push(`${sample.name}: expected duplicateIds=[${sample.expectDuplicateIds.join(', ')}], got [${measured.duplicateIds.join(', ')}]`);
  }
  if (measured.unknownCount !== sample.expectUnknownCount) {
    failures.push(`${sample.name}: expected unknownCount=${sample.expectUnknownCount}, got ${measured.unknownCount}`);
  }
  if (!sameArray(measured.systemIds, sample.expectSystemIds)) {
    failures.push(`${sample.name}: expected systemIds=[${sample.expectSystemIds.join(', ')}], got [${measured.systemIds.join(', ')}]`);
  }

  if (verbose) {
    console.log(`\n[${sample.name}] ${sample.file}`);
    measured.details.forEach((control) => {
      console.log(`  ${control.id}:${control.inputControl || '<none>'}`);
    });
  }
}

console.table(rows);

if (failures.length) {
  console.error('\nRegression failures:');
  failures.forEach((failure) => console.error(`  ${failure}`));
  process.exit(1);
}

console.log('\nGroup UserControls regression check passed.');
