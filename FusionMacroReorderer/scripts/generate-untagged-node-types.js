const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(__dirname, '..');
const nodesPanePath = path.join(projectRoot, 'src', 'nodesPane.js');
const catalogPath = path.join(projectRoot, 'FusionNodeCatalog.cleaned.json');
const docsOutPath = path.join(projectRoot, 'docs', 'untagged-node-types.txt');
const rootOutPath = path.resolve(projectRoot, '..', '..', 'untagged-node-types.txt');

function readExplicitTypeMap() {
  const source = fs.readFileSync(nodesPanePath, 'utf8');
  const blockMatch = source.match(/const explicitTypeMap = new Map\(\[(?<body>[\s\S]*?)\]\);/);
  if (!blockMatch || !blockMatch.groups || !blockMatch.groups.body) {
    throw new Error('Could not locate explicitTypeMap in src/nodesPane.js');
  }
  const body = blockMatch.groups.body;
  const entryRe = /\[\s*'([^']+)'\s*,\s*\{\s*key:\s*'([^']+)'\s*,\s*label:\s*'([^']+)'\s*\}\s*\]/g;
  const map = new Map();
  let m;
  while ((m = entryRe.exec(body)) !== null) {
    map.set(String(m[1]).toLowerCase(), { key: m[2], label: m[3] });
  }
  return map;
}

function classifyType(type, name, explicitTypeMap) {
  const typeStr = String(type || '').trim();
  const nameStr = String(name || '').trim();
  const typeLower = typeStr.toLowerCase();
  if (!typeLower) return null;

  if (explicitTypeMap.has(typeLower)) return explicitTypeMap.get(typeLower);

  const shapeSystemTypes = new Set([
    'sboolean',
    'sbspline',
    'schangestyle',
    'sduplicate',
    'sellipse',
    'sexpand',
    'sgrid',
    'sjitter',
    'smerge',
    'sngon',
    'soutline',
    'spolygon',
    'srectangle',
    'srender',
    'sstar',
    'stext',
    'stransform',
    'sbitmap',
    'sshape',
  ]);

  const blob = `${typeStr} ${nameStr}`.toLowerCase();
  const has = (re) => re.test(blob);
  const hasType = (re) => re.test(typeLower);

  if (has(/\b3d\b|camera3d|renderer3d|light3d|merge3d|transform3d|material|imageplane3d|shape3d/)) {
    return { key: 'three-d', label: '3D' };
  }
  if (hasType(/^u[a-z0-9]/)) {
    return { key: 'three-d', label: '3D' };
  }
  if (hasType(/^(?:light(?:ambient|directional|dome|point|spot|trim)|mtl|dimension\.)/)) {
    return { key: 'three-d', label: '3D' };
  }
  const looksLikeParticlePrefix = /^p[A-Z]/.test(typeStr);
  if (
    hasType(
      /^p(?:emitter|render|turbulence|bounce|directionalforce|drag|friction|imageemitter|kill|line|pointforce|randomforce|sprite|stylize|vortex|wind)\b/
    ) ||
    has(/\bparticle\b/)
  ) {
    return { key: 'particle', label: 'Particle' };
  }
  if (looksLikeParticlePrefix) {
    return { key: 'particle', label: 'Particle' };
  }
  if (hasType(/^ofx\./)) {
    return { key: 'ofx', label: 'OFX' };
  }
  if (hasType(/^fuse[._]/)) {
    return { key: 'fuse', label: 'Fuse' };
  }
  const looksLikeShapeSystemPrefix = /^s[A-Z]/.test(typeStr);
  if (shapeSystemTypes.has(typeLower) || looksLikeShapeSystemPrefix) {
    return { key: 'shape-system', label: 'Shape' };
  }
  if (
    hasType(
      /^(?:multipoly|polymask|maskpaint|bitmapmask|paintmask|rectangle(?:mask)?|ellipse(?:mask)?|polygon(?:mask)?|polyline(?:mask)?|bspline(?:mask)?|spline(?:mask)?|polypath|mask|outline)\b/
    )
  ) {
    return { key: 'mask', label: 'Mask' };
  }
  if (has(/multitext|loader|mediain|background|textplus|text3d|text|generator|solid|checker|gradient|fastnoise|noise/)) {
    return { key: 'generator', label: 'Generator' };
  }
  if (hasType(/^(?:multimerge|merge|dissolve|switch|wireless)\b/)) {
    return { key: 'flow', label: 'Flow' };
  }
  if (has(/transform|blur|glow|color|correct|resize|crop|key|tracker|optical|lens|vignette|channel|levels|brightness|contrast|sharpen|erodedilate|displace|letterbox/)) {
    return { key: 'effect', label: 'Effect' };
  }
  return null;
}

function buildReport() {
  const explicitTypeMap = readExplicitTypeMap();
  const catalog = JSON.parse(fs.readFileSync(catalogPath, 'utf8'));
  const uniqueTypes = new Set();
  for (const [key, value] of Object.entries(catalog || {})) {
    const type = String((value && value.type) || key || '').trim();
    if (type) uniqueTypes.add(type);
  }

  const tagCounts = new Map();
  const untagged = [];
  for (const type of Array.from(uniqueTypes).sort((a, b) => a.localeCompare(b))) {
    const meta = classifyType(type, type, explicitTypeMap);
    if (!meta || !meta.key) {
      untagged.push(type);
      continue;
    }
    tagCounts.set(meta.key, (tagCounts.get(meta.key) || 0) + 1);
  }

  const tagLines = Array.from(tagCounts.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, count]) => `  ${key}: ${count}`);

  const out = [
    `Total untagged types: ${untagged.length}`,
    '',
    'Tag counts:',
    ...tagLines,
    '',
    'Untagged types:',
    ...untagged,
    '',
  ].join('\n');

  return out;
}

function main() {
  const report = buildReport();
  fs.writeFileSync(docsOutPath, report, 'utf8');
  fs.writeFileSync(rootOutPath, report, 'utf8');
  process.stdout.write(`Updated:\\n- ${docsOutPath}\\n- ${rootOutPath}\\n`);
}

main();

