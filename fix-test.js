const fs = require('fs');
let c = fs.readFileSync('packages/i18n/tests/i18n.test.ts', 'utf8');
c = c.replace(/'zh-TW': \{ hello: '你好 \{name\}', plain: '檔案' \},/g, "'zh-TW': { hello: '你好 {name}', plain: '檔案' },\n    vi: { hello: 'Xin chao {name}', plain: 'Tep' },");
fs.writeFileSync('packages/i18n/tests/i18n.test.ts', c, 'utf8');
