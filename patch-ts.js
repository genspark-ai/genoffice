const fs = require('fs');

function patchRecordLangString(filePath) {
  if (!fs.existsSync(filePath)) return;
  let c = fs.readFileSync(filePath, 'utf8');
  // Look for 'zh-TW': 'something',
  c = c.replace(/'zh-TW': (['"].*?['"]|.*?),(\s*\n?\s*\})/g, "'zh-TW': $1,\n  vi: $1,$2");
  fs.writeFileSync(filePath, c, 'utf8');
}

function patchRecordLangObject(filePath) {
  if (!fs.existsSync(filePath)) return;
  let c = fs.readFileSync(filePath, 'utf8');
  // Look for 'zh-TW': { ... },
  c = c.replace(/'zh-TW': \{([^}]*)\},/g, "'zh-TW': {$1},\n  vi: {$1},");
  fs.writeFileSync(filePath, c, 'utf8');
}

patchRecordLangString('apps/docs/src/renderer/i18n/locale.tsx');
patchRecordLangObject('apps/markdown/src/renderer/i18n/locale.tsx');
patchRecordLangString('apps/markdown/src/renderer/i18n/locale.tsx');
patchRecordLangObject('packages/ui/src/strings-ai-panel-side.ts');
patchRecordLangObject('packages/ui/src/strings-crop-edges.ts');
patchRecordLangString('packages/ui/src/strings-image-viewer.ts');

console.log("Patched additional files");
