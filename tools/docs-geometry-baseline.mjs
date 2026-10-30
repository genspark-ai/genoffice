#!/usr/bin/env node
/**
 * Page-count and deviation summary of a Docs fidelity run directory, the gate
 * numbers the anchor / floating-table / table-width geometry stages are
 * measured against. The corpus and its run directories are private and stay
 * outside the repository; this prints the aggregate so a PR can quote
 * "before -> after" in words and numbers without copying sample ids.
 *
 * A run directory is what the docs-fidelity shooter writes:
 *   <run>/pagecounts.json   { <sample>: { word, gen, errors } }  Word vs GenOffice page counts
 *   <run>/wps-metrics.json  { pages: [{ doc, page, ours, label }] }  per-page pixel deviation
 *                           of the GenOffice render from Word (docs_wps_report.py)
 *
 *   node tools/docs-geometry-baseline.mjs <run-dir>                 summary
 *   node tools/docs-geometry-baseline.mjs <base-run> <after-run>    summary of both + per-sample
 *                                                                   page-count changes
 *   --json                                                          machine-readable summary
 *
 * Numbers reported: samples with both counts, page-count agreement, summed
 * absolute page difference, mean per-page deviation (0..1) and pages whose
 * deviation exceeds 5 % (the report's "deviating page" threshold).
 */
import { existsSync, readFileSync } from 'node:fs'
import { basename, join } from 'node:path'

const DEVIATING = 0.05

const args = process.argv.slice(2)
const asJson = args.includes('--json')
const runs = args.filter((a) => !a.startsWith('--'))
if (runs.length === 0 || runs.length > 2) {
  console.error('usage: node tools/docs-geometry-baseline.mjs <run-dir> [<after-run-dir>] [--json]')
  process.exit(2)
}

function summarise(dir) {
  const countsPath = join(dir, 'pagecounts.json')
  if (!existsSync(countsPath)) throw new Error(`${countsPath} not found`)
  const counts = JSON.parse(readFileSync(countsPath, 'utf8'))
  const samples = Object.entries(counts)
    .filter(([, v]) => Number.isFinite(v.word) && Number.isFinite(v.gen))
    .map(([id, v]) => ({ id, word: v.word, gen: v.gen }))
    .sort((a, b) => a.id.localeCompare(b.id))
  const agree = samples.filter((s) => s.word === s.gen).length
  const pageDelta = samples.reduce((n, s) => n + Math.abs(s.word - s.gen), 0)
  const metricsPath = join(dir, 'wps-metrics.json')
  let deviation = null
  if (existsSync(metricsPath)) {
    const pages = JSON.parse(readFileSync(metricsPath, 'utf8')).pages ?? []
    const scored = pages.filter((p) => typeof p.ours === 'number')
    deviation = {
      pages: scored.length,
      mean: scored.length ? scored.reduce((n, p) => n + p.ours, 0) / scored.length : 0,
      deviating: scored.filter((p) => p.ours > DEVIATING).length,
    }
  }
  return { run: basename(dir), samples, agree, pageDelta, deviation }
}

const pct = (x) => `${(x * 100).toFixed(2)}%`

function print(s) {
  console.log(`${s.run}: ${s.samples.length} samples with Word and GenOffice page counts`)
  console.log(
    `  page count agrees: ${s.agree}/${s.samples.length}, summed |gen - word| = ${s.pageDelta}`,
  )
  if (s.deviation)
    console.log(
      `  deviation vs Word: mean ${pct(s.deviation.mean)} over ${s.deviation.pages} pages, ${s.deviation.deviating} pages > ${pct(DEVIATING)}`,
    )
  else console.log('  deviation vs Word: no wps-metrics.json (run docs_wps_report.py first)')
}

const summaries = runs.map(summarise)
if (asJson) {
  console.log(
    JSON.stringify(
      summaries.map(({ samples, ...s }) => ({ ...s, samples: samples.length })),
      null,
      2,
    ),
  )
  process.exit(0)
}
for (const s of summaries) print(s)
if (summaries.length === 2) {
  const [base, after] = summaries
  const before = new Map(base.samples.map((s) => [s.id, s]))
  const moved = after.samples.filter((s) => before.get(s.id) && before.get(s.id).gen !== s.gen)
  console.log(`page-count changes ${base.run} -> ${after.run}: ${moved.length}`)
  for (const s of moved)
    console.log(`  ${s.id}: gen ${before.get(s.id).gen} -> ${s.gen} (word ${s.word})`)
}
