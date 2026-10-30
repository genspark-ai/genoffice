#!/usr/bin/env node
// Shadow-mode report for packages/docx-engine/src/resolve-props.ts: parse every
// docx under the given directories and print resolver vs legacy disagreements
// by surface and key, with example files.
//   node --import tsx tools/resolve-props-shadow.mjs <dir>... [--examples N] [--json out.json]
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { join, basename, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = fileURLToPath(new URL('.', import.meta.url))
const { parseDocx } = await import(resolve(here, '../packages/docx-engine/src/index.ts'))

const args = process.argv.slice(2)
const dirs = []
let examples = 3
let jsonOut
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--examples') examples = parseInt(args[++i], 10) || 3
  else if (args[i] === '--json') jsonOut = args[++i]
  else dirs.push(args[i])
}
if (dirs.length === 0) {
  console.error('usage: node --import tsx tools/resolve-props-shadow.mjs <dir>...')
  process.exit(2)
}

const files = []
for (const dir of dirs) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name)
    if (name.endsWith('.docx') && !name.startsWith('~$') && statSync(p).isFile()) files.push(p)
  }
}
files.sort()

const byKey = new Map()
let paragraphs = 0
let failed = 0
for (const file of files) {
  let doc
  try {
    doc = await parseDocx(new Uint8Array(readFileSync(file)), { shadowResolve: true })
  } catch (err) {
    failed++
    console.error(`parse failed: ${basename(file)}: ${err?.message ?? err}`)
    continue
  }
  paragraphs += doc.blocks.length
  for (const d of doc.extras.resolveDiff ?? []) {
    const key = `${d.surface ?? 'body'}:${d.key}`
    let e = byKey.get(key)
    if (!e) byKey.set(key, (e = { count: 0, files: new Map() }))
    e.count++
    const f = e.files.get(file)
    if (!f) e.files.set(file, { count: 1, sample: d })
    else f.count++
  }
}

const rows = [...byKey.entries()].sort((a, b) => b[1].count - a[1].count)
console.log(
  `files ${files.length} (failed ${failed}), blocks ${paragraphs}, disagreement keys ${rows.length}`,
)
console.log('')
console.log('| key | count | files | examples |')
console.log('|---|---|---|---|')
for (const [key, e] of rows) {
  const ex = [...e.files.entries()]
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, examples)
    .map(([file, f]) => {
      const s = f.sample
      const old = JSON.stringify(s.old)
      const eff = JSON.stringify(s.eff)
      return `${basename(file).replace(/_[0-9a-f]{12}\.docx$/, '')} (${f.count}: old=${old} eff=${eff}${s.text ? ` "${s.text}"` : ''})`
    })
  console.log(`| ${key} | ${e.count} | ${e.files.size} | ${ex.join('; ')} |`)
}
if (jsonOut) {
  const out = {}
  for (const [key, e] of rows) {
    out[key] = {
      count: e.count,
      files: Object.fromEntries([...e.files.entries()].map(([f, v]) => [basename(f), v])),
    }
  }
  writeFileSync(jsonOut, JSON.stringify(out, null, 2))
}
