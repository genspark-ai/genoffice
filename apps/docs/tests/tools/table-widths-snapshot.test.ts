/**
 * Table column-width decision snapshot over a docx corpus (tools/docs-table-widths-snapshot.mjs).
 * Records, for every body table and its nested tables, the widths resolveTableWidth draws
 * with the section budget, so the table-layout consolidation waves can prove "same widths"
 * without a screenshot gate. Skipped unless TABLE_WIDTHS_CORPUS names the corpus directories.
 * Two branches per table: `heuristic` (jsdom has no canvas; what CI's unit tests exercise) and
 * `measured` (NodeFontMetrics: the bundled / installed faces' real advances, what the live app's
 * canvas branch sees; characters outside every face take the per-script system cascade). TABLE_WIDTHS_BRANCH=heuristic|measured|both (default both).
 */
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, type TableModel } from '@genoffice/docx-engine'
import { pmDocOptions } from '../../src/renderer/editor/convert'
import { resolveTableWidth } from '../../src/renderer/editor/table-geometry'
import { budgetAvail, sectionWidthBudget } from '../../src/renderer/table-layout'
import { setDocFontTable } from '../../src/renderer/text-layout/font-rules'
import { NodeFontMetrics } from './node-text-measure'

const corpus = (process.env.TABLE_WIDTHS_CORPUS ?? '').split(',').filter(Boolean)
const out = process.env.TABLE_WIDTHS_OUT ?? ''
const branch = process.env.TABLE_WIDTHS_BRANCH ?? 'both'
const nodeMetrics = new NodeFontMetrics()

function nestedWidths(model: TableModel): unknown[] {
  const list: unknown[] = []
  model.rows.forEach((row, r) =>
    row.forEach((cell, c) =>
      (cell.nestedTables ?? []).forEach((nt, k) =>
        list.push({
          at: `${r}/${c}/${k}`,
          colTwips: nt.colWidthsTwips ?? null,
          widthPct: nt.widthPct ?? null,
          nested: nestedWidths(nt),
        }),
      ),
    ),
  )
  return list
}

async function snapshotDoc(path: string) {
  const parsed = await parseDocx(new Uint8Array(readFileSync(path)))
  // the measured branch resolves faces through the document's fontTable like the app
  setDocFontTable(parsed.fontTable ?? [])
  nodeMetrics.resetDocument()
  const sections = readSections(parsed)
  const options = pmDocOptions(parsed)
  const tables: Record<string, unknown> = {}
  for (const block of parsed.blocks) {
    if (block.type !== 'table' || !block.table || block.hidden) continue
    const budget = sectionWidthBudget(sections, block.docxIndex ?? null)
    const resolve = (metrics: typeof nodeMetrics | undefined) => {
      const resolved = resolveTableWidth(
        block.table!,
        budget
          ? { avail: budgetAvail(block.table!, budget), fit: budget.fit, paper: budget.paper }
          : null,
        {
          legacyIndent: options.legacyTableIndent ?? false,
          defaults: options.defaultRun,
          ...(metrics ? { metrics } : {}),
        },
      )
      const { model, log, ...geom } = resolved
      return { ...geom, log, nested: nestedWidths(model) }
    }
    const heuristic = branch !== 'measured' ? resolve(undefined) : null
    const measured = branch !== 'heuristic' ? resolve(nodeMetrics) : null
    tables[`t${block.docxIndex ?? '?'}`] = {
      declared: block.table.colWidthsTwips ?? null,
      ...(heuristic ?? {}),
      ...(measured ? { measured } : {}),
    }
  }
  return tables
}

describe.skipIf(!corpus.length || !out)('table widths snapshot', () => {
  it('records every table width decision over the corpus', async () => {
    const files: string[] = []
    for (const dir of corpus)
      for (const f of readdirSync(dir).sort()) if (/\.docx$/i.test(f)) files.push(join(dir, f))
    const result: Record<string, unknown> = {}
    let tables = 0
    for (const file of files) {
      const key = `${basename(join(file, '..'))}/${basename(file, '.docx')}`
      try {
        const doc = await snapshotDoc(file)
        tables += Object.keys(doc).length
        result[key] = doc
      } catch (e) {
        result[key] = { error: String(e) }
      }
    }
    writeFileSync(out, JSON.stringify(result, null, 1))
    // vitest keeps console output of passing tests; the tool prints this sidecar
    const { resolved, system, unresolved, unresolvedFamilies, scripts } = nodeMetrics.stats
    writeFileSync(
      out.replace(/\.json$/, '') + '.stats.json',
      JSON.stringify(
        {
          branch,
          docs: files.length,
          tables,
          measuredChars: { real: resolved, system, fallback: unresolved },
          fallbackFamilies: [...unresolvedFamilies].sort(),
          scripts: Object.fromEntries([...scripts].sort()),
        },
        null,
        1,
      ),
    )
    expect(files.length).toBeGreaterThan(0)
  }, 3_600_000)
})
