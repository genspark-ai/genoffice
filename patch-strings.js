const fs = require("fs");
const files = [
  "apps/markdown/src/renderer/i18n/strings.ts",
  "apps/pdf/src/renderer/i18n/strings.ts",
  "apps/shell/src/renderer/src/strings.ts"
];

for (const file of files) {
  let code = fs.readFileSync(file, "utf8");
  if (code.includes("vi: {")) {
    console.log("Already has vi in", file);
    continue;
  }
  
  // Find "en: {"
  const enStart = code.indexOf("en: {");
  if (enStart === -1) {
    console.log("Could not find en: { in", file);
    continue;
  }
  
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
    code = code.slice(0, enEnd) + ",\n  " + viBlock + code.slice(enEnd);
    fs.writeFileSync(file, code);
    console.log("Patched", file);
  }
}

