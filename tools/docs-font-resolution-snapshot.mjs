#!/usr/bin/env node
/**
 * Font-resolution decision snapshot for the font-rules consolidation waves.
 *
 *   node tools/docs-font-resolution-snapshot.mjs <out.json> <corpus-dir> [<corpus-dir>...]
 *   node tools/docs-font-resolution-snapshot.mjs --diff <before.json> <after.json> [--ignore <key>...]
 *   node tools/docs-font-resolution-snapshot.mjs --stats <snapshot.json>
 *       (decision / factor source histograms, and the names behind every
 *       non-resolver factor source so a probe round can target them)
 *
 * The snapshot runs apps/docs/tests/tools/font-resolution-snapshot.test.ts under
 * vitest (jsdom: no canvas and no font IPC, so "installed" answers come from the
 * word-faces.json macOS snapshot and availability probes answer "missing").
 * It proves "same decisions", not "same pixels". A refactor wave must diff empty;
 * `--ignore decisions` skips the new entry point's record when the baseline
 * checkout predates it.
 */
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const args = process.argv.slice(2)

function load(p) {
  return JSON.parse(readFileSync(p, 'utf8'))
}

if (args[0] === '--diff') {
  const [a, b] = args.slice(1, 3).map(load)
  const ignore = new Set()
  for (let i = 3; i < args.length; i++) if (args[i] === '--ignore') ignore.add(args[++i])
  let diffs = 0
  const byKey = new Map()
  const walk = (x, y, path) => {
    if (JSON.stringify(x) === JSON.stringify(y)) return
    if (x && y && typeof x === 'object' && typeof y === 'object') {
      const keys = Array.isArray(x)
        ? [...Array(Math.max(x.length, y.length)).keys()]
        : [...new Set([...Object.keys(x), ...Object.keys(y)])]
      for (const k of keys) {
        if (ignore.has(String(k))) continue
        walk(x[k], y[k], `${path}/${k}`)
      }
      return
    }
    diffs++
    const key = path
      .replace(/^\/[^/]+\/[^/]+/, '/<doc>')
      .replace(/\/(faces|decisions|runPairs|cs)\/[^/]+/, '/$1/<name>')
    byKey.set(key, (byKey.get(key) ?? 0) + 1)
    if (diffs <= 80) console.log(`${path}: ${JSON.stringify(x)} -> ${JSON.stringify(y)}`)
  }
  walk(a, b, '')
  for (const [k, n] of [...byKey].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${k}`)
  console.log(diffs === 0 ? 'no differences' : `${diffs} differences`)
  process.exit(diffs === 0 ? 0 : 1)
}

if (args[0] === '--stats') {
  const snap = load(args[1])
  const docs = Object.values(snap)
  const names = new Set()
  const sources = new Map()
  const factorSources = new Map()
  const factorNames = new Map()
  let errors = 0
  let pairs = 0
  for (const d of docs) {
    if (d.error) {
      errors++
      continue
    }
    for (const n of Object.keys(d.faces ?? {})) names.add(n)
    pairs += Object.keys(d.runPairs ?? {}).length
    for (const dec of Object.values(d.decisions ?? {})) {
      sources.set(dec.source, (sources.get(dec.source) ?? 0) + 1)
      factorSources.set(dec.factorSource, (factorSources.get(dec.factorSource) ?? 0) + 1)
      if (dec.factorSource !== 'resolver') {
        if (!factorNames.has(dec.factorSource)) factorNames.set(dec.factorSource, new Map())
        const names = factorNames.get(dec.factorSource)
        names.set(dec.declared, (names.get(dec.declared) ?? 0) + 1)
      }
    }
  }
  console.log(
    `${docs.length} documents (${errors} unparsed), ${names.size} distinct names, ${pairs} run slot pairs`,
  )
  if (sources.size) {
    console.log('decision sources (name x document):')
    for (const [k, n] of [...sources].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${k}`)
    console.log('factor sources:')
    for (const [k, n] of [...factorSources].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${k}`)
    for (const [source, names] of factorNames) {
      console.log(`names by factor source: ${source} (name: documents)`)
      for (const [name, n] of [...names].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${name}`)
    }
  }
  process.exit(0)
}

const [out, ...dirs] = args
if (!out || dirs.length === 0) {
  console.error(
    'usage: node tools/docs-font-resolution-snapshot.mjs <out.json> <corpus-dir>... | --diff <a> <b> [--ignore key] | --stats <a>',
  )
  process.exit(2)
}
const r = spawnSync('npx', ['vitest', 'run', 'tests/tools/font-resolution-snapshot.test.ts'], {
  cwd: resolve('apps/docs'),
  stdio: 'inherit',
  env: {
    ...process.env,
    FONT_SNAPSHOT_CORPUS: dirs.map((d) => resolve(d)).join(','),
    FONT_SNAPSHOT_OUT: resolve(out),
  },
})
process.exit(r.status ?? 1)
