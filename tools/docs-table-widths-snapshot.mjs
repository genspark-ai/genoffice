#!/usr/bin/env node
/**
 * Table column-width decision snapshot for the table-layout consolidation waves.
 *
 *   node tools/docs-table-widths-snapshot.mjs [--branch heuristic|measured|both] <out.json> <corpus-dir>...
 *   node tools/docs-table-widths-snapshot.mjs --diff <before.json> <after.json>
 *
 * Runs apps/docs/tests/tools/table-widths-snapshot.test.ts under vitest (jsdom) and records
 * every body / nested table's resolved widths on two branches: the heuristic text measure
 * (what unit tests exercise) and, under `measured`, the real advances of the bundled /
 * installed faces (what the live app's canvas branch computes; ttf and woff2 @font-face files,
 * then the per-script system cascade for uncovered characters). A refactor wave must diff
 * empty; a rule change lists each moved table and a histogram by field.
 */
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const args = process.argv.slice(2)
if (args[0] === '--diff') {
  const [a, b] = args.slice(1).map((p) => JSON.parse(readFileSync(p, 'utf8')))
  let diffs = 0
  const byKey = new Map()
  const tables = new Set()
  const walk = (x, y, path) => {
    if (JSON.stringify(x) === JSON.stringify(y)) return
    if (x && y && typeof x === 'object' && typeof y === 'object') {
      const keys = Array.isArray(x)
        ? [...Array(Math.max(x.length, y.length)).keys()]
        : [...new Set([...Object.keys(x), ...Object.keys(y)])]
      for (const k of keys) walk(x[k], y[k], `${path}/${k}`)
      return
    }
    diffs++
    tables.add(path.split('/').slice(1, 4).join('/'))
    const key = path.replace(/^\/[^/]+\/[^/]+\/t\d+/, '/<doc>/<table>').replace(/\/\d+/g, '/<i>')
    byKey.set(key, (byKey.get(key) ?? 0) + 1)
    if (diffs <= 60) console.log(`${path}: ${JSON.stringify(x)} -> ${JSON.stringify(y)}`)
  }
  walk(a, b, '')
  for (const [k, n] of [...byKey].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${k}`)
  console.log(diffs === 0 ? 'no differences' : `${diffs} differences in ${tables.size} tables`)
  process.exit(diffs === 0 ? 0 : 1)
}
let branch = 'both'
if (args[0] === '--branch') {
  args.shift()
  branch = args.shift() ?? ''
}
const [out, ...dirs] = args
if (!out || dirs.length === 0 || !['heuristic', 'measured', 'both'].includes(branch)) {
  console.error(
    'usage: node tools/docs-table-widths-snapshot.mjs [--branch heuristic|measured|both] <out.json> <corpus-dir>... | --diff <a> <b>',
  )
  process.exit(2)
}
const r = spawnSync('npx', ['vitest', 'run', 'tests/tools/table-widths-snapshot.test.ts'], {
  cwd: resolve('apps/docs'),
  stdio: 'inherit',
  env: {
    ...process.env,
    TABLE_WIDTHS_CORPUS: dirs.map((d) => resolve(d)).join(','),
    TABLE_WIDTHS_OUT: resolve(out),
    TABLE_WIDTHS_BRANCH: branch,
  },
})
if (r.status === 0) {
  try {
    const stats = JSON.parse(
      readFileSync(resolve(out).replace(/\.json$/, '') + '.stats.json', 'utf8'),
    )
    const { real, system = 0, fallback } = stats.measuredChars
    console.log(
      `table widths snapshot (${stats.branch}): ${stats.docs} docs, ${stats.tables} tables -> ${out}` +
        (stats.branch === 'heuristic'
          ? ''
          : `; measured chars ${real} chain / ${system} system cascade / ${fallback} fallback` +
            (stats.fallbackFamilies.length ? ` (${stats.fallbackFamilies.join(', ')})` : '') +
            (stats.scripts
              ? `; by script ${Object.entries(stats.scripts)
                  .map(([k, v]) => `${k} ${v.system}/${v.unresolved}`)
                  .join(', ')}`
              : '')),
    )
  } catch {
    /* no sidecar: the test was skipped */
  }
}
process.exit(r.status ?? 1)
