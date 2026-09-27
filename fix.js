const fs = require('fs');
const glob = require('glob');
const files = glob.sync('apps/*/src/renderer*/i18n*/locale.tsx');
files.push('apps/shell/src/renderer/src/locale.tsx');
for (const file of files) {
  if (!fs.existsSync(file)) continue;
  let content = fs.readFileSync(file, 'utf8');
  
  if (!content.includes('vi: ')) {
    content = content.replace(/(['`"]zh-TW['`"]:.*?)(,\n|\n)/g, (match, p1, p2) => {
      return p1 + p2 + '  vi: ' + p1.substring(p1.indexOf(':') + 1).trim() + p2;
    });
    fs.writeFileSync(file, content);
    console.log('Patched', file);
  } else {
    // maybe vi exists in AI_LANG_DIRECTIVES but not in DATE_LOCALES?
    // Let's force it if it's missing in DATE_LOCALES
    if (content.includes('DATE_LOCALES') && !content.match(/vi:\s*['"]zh-TW['"]/)) {
      content = content.replace(/(['`"]zh-TW['"]:\s*['"]zh-TW['"])(,\n|\n)/g, (match, p1, p2) => {
         return p1 + p2 + '  vi: \'zh-TW\'' + p2;
      });
      fs.writeFileSync(file, content);
      console.log('Patched DATE_LOCALES', file);
    }
  }
}
