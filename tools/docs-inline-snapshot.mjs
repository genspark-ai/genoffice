#!/usr/bin/env node
/**
 * Inline-layout decision snapshot for the text-layout consolidation waves.
 *
 *   node tools/docs-inline-snapshot.mjs <out.json> <corpus-dir> [<corpus-dir>...]
 *   node tools/docs-inline-snapshot.mjs --diff <before.json> <after.json>
 *
 * The snapshot runs apps/docs/tests/tools/inline-snapshot.test.ts under vitest
 * (jsdom: installed-font probes answer "missing", so availability-dependent
 * answers are stable but not the live app's). A refactor wave must diff empty;
 * every remaining difference is listed per document / face / paragraph.
 */
import { spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const args = process.argv.slice(2)
if (args[0] === '--diff') {
  const [a, b] = args.slice(1).map((p) => JSON.parse(readFileSync(p, 'utf8')))
  let diffs = 0
  const byKey = new Map()
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
    const key = path
      .replace(/\/prod_[^/]+/, '/<doc>')
      .replace(/\/paras\/\d+/, '/paras/<i>')
      .replace(/\/faces\/[^/]+/, '/faces/<name>')
    byKey.set(key, (byKey.get(key) ?? 0) + 1)
    if (diffs <= 60) console.log(`${path}: ${JSON.stringify(x)} -> ${JSON.stringify(y)}`)
  }
  walk(a, b, '')
  for (const [k, n] of [...byKey].sort((p, q) => q[1] - p[1])) console.log(`  ${n}\t${k}`)
  console.log(diffs === 0 ? 'no differences' : `${diffs} differences`)
  process.exit(diffs === 0 ? 0 : 1)
}
const [out, ...dirs] = args
if (!out || dirs.length === 0) {
  console.error(
    'usage: node tools/docs-inline-snapshot.mjs <out.json> <corpus-dir>... | --diff <a> <b>',
  )
  process.exit(2)
}
const r = spawnSync('npx', ['vitest', 'run', 'tests/tools/inline-snapshot.test.ts'], {
  cwd: resolve('apps/docs'),
  stdio: 'inherit',
  env: {
    ...process.env,
    INLINE_SNAPSHOT_CORPUS: dirs.map((d) => resolve(d)).join(','),
    INLINE_SNAPSHOT_OUT: resolve(out),
  },
})
process.exit(r.status ?? 1)
