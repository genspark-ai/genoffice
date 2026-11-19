#!/usr/bin/env node
// Shadow audit of Word's face resolver against the hand line-factor table.
// Input: a JSON array of fontTable names with hints
//   [{ name, docs, altName?, panose?, family?, charset? }, ...]
// Installed metrics come from this machine's font index (the renderer's IPC
// source), so run it on the Mac the truths were probed on.
// Usage: npx tsx tools/font-resolve-shadow.mjs <names.json> [--all]
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const [, , namesPath, ...flags] = process.argv
if (!namesPath) {
  console.error('usage: npx tsx tools/font-resolve-shadow.mjs <names.json> [--all]')
  process.exit(2)
}
const showAll = flags.includes('--all')

const { familyVerticalMetrics } = await import('../packages/font-metrics/src/index.ts')
const { faceKey, resolveFace, storeContext } =
  await import('../packages/font-metrics/src/resolve.ts')
const wordFaces = JSON.parse(
  readFileSync(join(here, '../packages/font-metrics/word-faces.json'), 'utf8'),
)
const lm = await import('../apps/docs/src/renderer/line-metrics.ts')

const rows = JSON.parse(readFileSync(namesPath, 'utf8'))
const hints = new Map(rows.map((r) => [faceKey(r.name), r]))

const EA = 0x3e0000
const ctx = storeContext(wordFaces, {
  installed: (key) => {
    const m = familyVerticalMetrics(key)
    if (!m) return null
    const u = m.unitsPerEm
    return {
      family: key,
      source: 'installed',
      hheaTotal: (m.hheaAscender - m.hheaDescender + m.hheaLineGap) / u,
      ascent: m.hheaAscender / u,
      descent: -m.hheaDescender / u,
      xAvg: m.xAvgCharWidth == null ? null : m.xAvgCharWidth / u,
      eastAsian: ((m.codePageRange1 ?? 0) & EA) !== 0,
    }
  },
  hints: (key) => {
    const h = hints.get(key)
    if (!h) return undefined
    const out = {}
    for (const k of ['altName', 'panose', 'family', 'charset']) if (h[k]) out[k] = h[k]
    return out
  },
})

lm.setDocFontTable(rows.map((r) => ({ ...r })))
const r4 = (v) => Math.round(v * 1e4) / 1e4
const out = []
for (const r of rows) {
  const legacy = lm.tableLineHeightFactor(r.name)
  const now = lm.lineHeightFactor(r.name)
  const face = resolveFace(r.name, ctx)
  const resolver = face ? face.lineFactor : null
  const agree = resolver !== null && Math.abs(resolver - legacy) <= 0.0005
  if (!showAll && agree) continue
  out.push({
    name: r.name,
    docs: r.docs,
    legacy: r4(legacy),
    now: r4(now),
    resolver,
    face: face ? `${face.family} (${face.source}${face.substituted ? ', subst' : ''})` : '-',
    ea: face?.metrics.eastAsian ? 'EA' : '',
  })
}
out.sort((a, b) => b.docs - a.docs)
console.log('| name | docs | table | resolver | now | target face | EA |')
console.log('|---|---|---|---|---|---|---|')
for (const o of out)
  console.log(
    `| ${o.name} | ${o.docs} | ${o.legacy} | ${o.resolver ?? '-'} | ${o.now} | ${o.face} | ${o.ea} |`,
  )
console.error(`${out.length} of ${rows.length} names listed`)
