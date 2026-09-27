const fs = require("fs");
const glob = require("glob");

const files = glob.sync("apps/*/src/renderer*/i18n*/locale.tsx");
files.push("apps/shell/src/renderer/src/locale.tsx");

for (const file of files) {
  if (!fs.existsSync(file)) continue;
  let content = fs.readFileSync(file, "utf8");
  if (!content.includes("vi: ")) {
    content = content.replace(/'zh-TW': (.*?),\n/g, "'zh-TW': $1,\n  vi: $1,\n");
    content = content.replace(/"zh-TW": (.*?),\n/g, "\"zh-TW\": $1,\n  vi: $1,\n");
    content = content.replace(/'zh-TW'/g, "\"zh-TW\"");
    content = content.replace(/'zh-TW': (.*?),/g, "\"zh-TW\": $1,\n  vi: $1,");
    
    // Actually the regex should just match:
    content = content.replace(/.\bzh-TW\b.: (.*?),\n/g, "\x27zh-TW\x27: $1,\n  vi: $1,\n");
    
    fs.writeFileSync(file, content);
    console.log("Patched", file);
  }
}

