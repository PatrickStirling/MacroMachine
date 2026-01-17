#!/usr/bin/env node
/**
 * Simple sanity check for Blend exports.
 * Usage: node scripts/check-blend-toggle.js <file.setting> [controlName]
 */
const fs = require('fs');
const path = require('path');

function printUsage() {
  console.error('Usage: node scripts/check-blend-toggle.js <file.setting> [controlName]');
  process.exit(1);
}

if (process.argv.length < 3) {
  printUsage();
}

const filePath = process.argv[2];
const controlName = process.argv[3] || 'Blend';

const fullPath = path.resolve(process.cwd(), filePath);
if (!fs.existsSync(fullPath)) {
  console.error(`File not found: ${fullPath}`);
  process.exit(1);
}

const content = fs.readFileSync(fullPath, 'utf8');

function findMatchingBrace(text, openIndex) {
  let depth = 0;
  let inString = false;
  for (let i = openIndex; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (ch === '"' && text[i - 1] !== '\\') inString = false;
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === '{') { depth++; continue; }
    if (ch === '}') {
      depth--;
      if (depth === 0) return i;
      continue;
    }
  }
  return -1;
}

function extractBlendBlocks(text, targetControl) {
  const blocks = [];
  const regex = /InstanceInput\s*\{/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    const open = text.indexOf('{', match.index);
    if (open < 0) break;
    const close = findMatchingBrace(text, open);
    if (close < 0) break;
    const block = text.slice(match.index, close + 1);
    const controlPattern = new RegExp(`Source\\s*=\\s*"${targetControl}"`, 'i');
    if (controlPattern.test(block)) blocks.push(block);
    regex.lastIndex = close + 1;
  }
  return blocks;
}

const blocks = extractBlendBlocks(content, controlName);
if (!blocks.length) {
  console.error(`No InstanceInput blocks found for control "${controlName}" in ${filePath}`);
  process.exit(2);
}

const failing = blocks.filter(block => !/INPID_InputControl\s*=\s*"CheckboxControl"/i.test(block));
if (failing.length) {
  console.error(`Found ${failing.length} block(s) without CheckboxControl:`);
  failing.forEach((block, idx) => {
    console.error(`--- Block ${idx + 1} ---`);
    console.error(block.trim().split('\n').map(l => `  ${l}`).join('\n'));
  });
  process.exit(3);
}

console.log(`All ${blocks.length} "${controlName}" InstanceInput block(s) use CheckboxControl ✅`);
