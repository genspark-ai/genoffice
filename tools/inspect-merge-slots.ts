/**
 * Offline mail-merge slot inspection: point it at a template .docx and see the
 * blocks the merge counts as record slots, so a "page holds N" report can be
 * checked without launching the editor.
 *
 * Usage: npx tsx tools/inspect-merge-slots.ts <template.docx>
 *
 * It reuses describeRecordSlots/countRecordSlots from the renderer's template
 * module. Those only walk a plain { type, text, marks, content } tree, so the
 * script rebuilds the same shape from the parsed Block[] instead of starting
 * ProseMirror: one paragraph node per block, one text node per run, with
 * run.instrField becoming the instrField mark. Nested table rows collapse into
 * their table's block (same granularity the dialog's diagnostics report at).
 */
import { readFileSync } from 'node:fs'
import { mergeFieldName, parseDocx, splitInstruction } from '@genoffice/docx-engine'
import type { Block, Run, TableModel } from '@genoffice/docx-engine'
import { describeRecordSlots } from '../apps/docs/src/renderer/mail-merge/template'
import type { PmNode } from '../apps/docs/src/renderer/editor/convert'

const FIELD_MARK = 'instrField'

/** one paragraph's runs -> the text nodes ProseMirror would hold */
function runsToNodes(runs: readonly Run[] = []): PmNode[] {
  const nodes: PmNode[] = []
  for (const run of runs) {
    // ProseMirror rejects empty text nodes, and convert drops the empty runs
    if (run.text === '') continue
    const node: PmNode = { type: 'text', text: run.text }
    if (run.instrField !== undefined) {
      node.marks = [{ type: FIELD_MARK, attrs: { instr: run.instrField } }]
    }
    nodes.push(node)
  }
  return nodes
}

function paragraphOf(runs?: readonly Run[]): PmNode {
  const content = runsToNodes(runs)
  return content.length > 0 ? { type: 'paragraph', content } : { type: 'paragraph' }
}

/** the parsed Block[] as the flat node tree countRecordSlots walks */
export function blocksToPmDoc(blocks: readonly Block[]): PmNode {
  const content: PmNode[] = []
  for (const block of blocks) {
    if (block.hidden) continue
    if (block.type === 'table') {
      const rows = block.table?.rows ?? []
      content.push({
        type: 'docTable',
        content: rows.map((row) => ({
          type: 'docTableRow',
          content: row.map((cell) => ({
            type: 'docTableCell',
            content:
              cell.richParas && cell.richParas.length > 0
                ? cell.richParas.map((para) => paragraphOf(para.runs))
                : [
                    {
                      type: 'paragraph',
                      content: cell.paras
                        .filter((text) => text !== '')
                        .map((text) => ({ type: 'text', text })),
                    },
                  ],
          })),
        })),
      })
      continue
    }
    if (block.type === 'paragraph' || block.type === 'heading' || block.type === 'listItem') {
      content.push(paragraphOf(block.runs))
      continue
    }
    // image / passthrough: no text ProseMirror renders, so an invisible block
    content.push({ type: block.type === 'image' ? 'docImage' : 'docHtmlBlock' })
  }
  return { type: 'doc', content }
}

/** per-row detail for tables: record rows hide inside one top-level block */
function printTableRows(blocks: readonly Block[]): void {
  for (const block of blocks) {
    if (block.type !== 'table' || !block.table) continue
    console.log(`table block rows: ${block.table.rows.length}`)
    block.table.rows.forEach((row, index) => {
      const runs = row.flatMap((cell) => cell.richParas?.flatMap((para) => para.runs) ?? [])
      const next = runs.filter((run) => /^\s*NEXT\b/i.test(run.instrField ?? '')).length
      const fields = runs.filter((run) => run.instrField !== undefined).length
      const text = runs
        .map((run) => run.text)
        .join('')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 50)
      console.log(`  row #${index} next=${next} fields=${fields} text=${JSON.stringify(text)}`)
    })
  }
}

/** --full: every row/cell/paragraph, plus the record each field reads */
function printFull(parsed: Awaited<ReturnType<typeof parseDocx>>): void {
  let cursor = 0
  let highest = -1
  let lastWasNext = false
  const hits: Array<{ path: string; name: string; record: number }> = []
  const nextHits: string[] = []

  const walkRuns = (runs: readonly Run[], path: string) => {
    const fields: string[] = []
    let nextInfo: string | null = null
    runs.forEach((run, index) => {
      const instr = run.instrField
      if (instr === undefined) {
        if ((run.text ?? '').trim() !== '') lastWasNext = false
        return
      }
      const [keyword] = splitInstruction(instr)
      if (keyword === 'NEXT') {
        if (!lastWasNext) cursor++
        lastWasNext = true
        const textAfter = runs.slice(index + 1).some((later) => (later.text ?? '').trim() !== '')
        nextInfo = `NEXT 于 run ${index + 1}/${runs.length}（${textAfter ? '段中，其后仍有文本' : '段末'}）→ 游标推进到 ${cursor}`
        nextHits.push(`${path} · ${nextInfo}`)
        return
      }
      lastWasNext = false
      const name = mergeFieldName(instr) ?? keyword
      highest = Math.max(highest, cursor)
      fields.push(`${name}→第${cursor + 1}条`)
      hits.push({ path, name, record: cursor + 1 })
    })
    return { fields, nextInfo }
  }

  const cellText = (cell: { richParas?: Array<{ runs: readonly Run[] }>; paras: string[] }): string =>
    (cell.richParas && cell.richParas.length > 0
      ? cell.richParas.map((para) => para.runs.map((run) => run.text).join('')).join(' / ')
      : cell.paras.join(' / ')
    )
      .replace(/\s+/g, ' ')
      .trim()

  const printTable = (table: TableModel, prefix: string): void => {
    console.log(`${prefix}table rows=${table.rows.length} cols=${table.colWidthsPct?.length ?? '?'}`)
    table.rows.forEach((row, rowIndex) => {
      const flags = row
        .map((cell, cellIndex) => {
          const marks: string[] = []
          if (cell.colSpan && cell.colSpan > 1) marks.push(`colSpan=${cell.colSpan}`)
          if (cell.vMerge) marks.push(`vMerge=${cell.vMerge}`)
          if (cell.gridGap) marks.push('gridGap')
          if (cell.nestedTables?.length) marks.push(`nested=${cell.nestedTables.length}`)
          if (cell.anchoredBoxes?.length) marks.push(`boxes=${cell.anchoredBoxes.length}`)
          return marks.length > 0 ? ` cell#${cellIndex}[${marks.join(' ')}]` : ''
        })
        .join('')
      console.log(`${prefix}  row #${rowIndex} cells=${row.length}${flags}`)
      row.forEach((cell, cellIndex) => {
        const path = `${prefix}row#${rowIndex} cell#${cellIndex}`
        const paras =
          cell.richParas && cell.richParas.length > 0
            ? cell.richParas.map((para, paraIndex) => ({
                paraIndex,
                text: para.runs.map((run) => run.text).join('').replace(/\s+/g, ' ').trim().slice(0, 40),
                result: walkRuns(para.runs, `${path} para#${paraIndex}`),
              }))
            : []
        console.log(`${prefix}    cell#${cellIndex} text=${JSON.stringify(cellText(cell).slice(0, 60))}`)
        for (const para of paras) {
          if (para.result.fields.length === 0 && !para.result.nextInfo) continue
          const parts: string[] = []
          if (para.result.fields.length > 0) parts.push(`域: ${para.result.fields.join(', ')}`)
          if (para.result.nextInfo) parts.push(para.result.nextInfo)
          console.log(`${prefix}      para#${para.paraIndex} ${parts.join(' · ')}`)
        }
        for (const nested of cell.nestedTables ?? []) printTable(nested, `${prefix}    cell#${cellIndex} `)
      })
    })
  }

  console.log('')
  console.log('=== full structure: top-level blocks ===')
  parsed.blocks.forEach((block, index) => {
    if (block.type === 'table' && block.table) {
      const text = block.table.rows
        .flatMap((row) => row.map(cellText))
        .join('')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 60)
      console.log(`#${index} table text=${JSON.stringify(text)}`)
      printTable(block.table, '  ')
      return
    }
    const result = walkRuns(block.runs ?? [], `#${index} ${block.type}`)
    const text = (block.runs ?? [])
      .map((run) => run.text)
      .join('')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 60)
    console.log(`#${index} ${block.type}${block.hidden ? ' hidden' : ''} text=${JSON.stringify(text)}`)
    if (result.fields.length > 0 || result.nextInfo) {
      const parts: string[] = []
      if (result.fields.length > 0) parts.push(`域: ${result.fields.join(', ')}`)
      if (result.nextInfo) parts.push(result.nextInfo)
      console.log(`   ${parts.join(' · ')}`)
    }
  })

  console.log('')
  console.log('=== header / footer fields ===')
  for (const [name, paras] of [
    ['header', parsed.headerParas],
    ['footer', parsed.footerParas],
  ] as const) {
    const found = (paras ?? []).flatMap((para, paraIndex) =>
      para.runs
        .filter((run) => run.instrField !== undefined)
        .map((run) => `${name} para#${paraIndex}: ${run.instrField}`),
    )
    console.log(found.length > 0 ? found.join('\n') : `${name}: 无域（${name}Paras ${paras ? paras.length : 0} 段）`)
  }

  console.log('')
  console.log('=== cursor walk result ===')
  console.log(`final cursor=${cursor} highest=${highest} → record slots per page=${Math.max(1, highest + 1)}`)
  const quantity = hits.filter((hit) => hit.name === '数量')
  console.log(`数量 域: ${quantity.map((hit) => `${hit.path} → 第${hit.record}条`).join(' | ')}`)
  console.log(`NEXT 位置: ${nextHits.join(' | ')}`)
}

async function main(): Promise<void> {
  const path = process.argv[2]
  if (!path) {
    console.error('usage: npx tsx tools/inspect-merge-slots.ts <template.docx>')
    process.exitCode = 2
    return
  }
  const bytes = new Uint8Array(readFileSync(path))
  const parsed = await parseDocx(bytes)
  const trace = describeRecordSlots(blocksToPmDoc(parsed.blocks))

  console.log(`template: ${path}`)
  console.log(`blocks: ${trace.blocks.length} · record slots per page: ${trace.slots}`)
  for (const block of trace.blocks) {
    const slot = block.contributed ? `slot ${block.slotAfter}` : '-'
    const fields = block.fields.length > 0 ? ` fields=${JSON.stringify(block.fields)}` : ''
    console.log(
      `#${String(block.index).padStart(3)} ${block.type.padEnd(10)} ${
        block.visible ? 'visible' : 'empty  '
      } next=${block.nextFields} ${slot.padEnd(7)} text=${JSON.stringify(block.text)}${fields}`,
    )
  }
  if (trace.nextFields.length > 0) {
    console.log(
      'NEXT fields: ' +
        trace.nextFields
          .map((entry) => `#${entry.block} text=${JSON.stringify(entry.text)} instr=${JSON.stringify(entry.instr)}`)
          .join(' | '),
    )
  }
  printTableRows(parsed.blocks)
  if (process.argv.includes('--full')) printFull(parsed)
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err))
  process.exitCode = 1
})
