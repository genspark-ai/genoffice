const fs = require("fs");
const files = [
  "apps/docs/src/main/docs-main.ts",
  "apps/docs/src/main/updater.ts",
  "apps/html/src/main/html-main.ts",
  "apps/markdown/src/main/markdown-main.ts",
  "apps/pdf/src/main/pdf-main.ts",
  "apps/sheets/src/main/sheets-main.ts",
  "apps/shell/src/main/index.ts",
  "apps/shell/src/main/updater.ts",
  "apps/slides/src/main/i18n-main.ts",
  "packages/i18n/tests/i18n.test.ts"
];

for (const file of files) {
  let code = fs.readFileSync(file, "utf8");
  if (code.includes("vi: {")) continue;
  
  // Find "en: {"
  const enStart = code.indexOf("en: {");
  if (enStart === -1) continue;
  
  let braceCount = 0;
  let enEnd = -1;
  for (let i = enStart + 4; i < code.length; i++) {
    if (code[i] === "{") braceCount++;
    if (code[i] === "}") {
      braceCount--;
      if (braceCount === 0) {
        enEnd = i + 1;
        break;
      }
    }
  }
  
  if (enEnd !== -1) {
    const enBlock = code.substring(enStart, enEnd);
    const viBlock = enBlock.replace(/^en:/, "vi:");
    code = code.slice(0, enEnd) + ",\n    " + viBlock + code.slice(enEnd);
    fs.writeFileSync(file, code);
    console.log("Patched", file);
  }
}

