/**
 * Plain-text / Markdown serialization of the docx block tree.
 *
 * Pure engine code (no Electron, no DOM): feed it the document's Block[]
 * (the same interchange model the editor renders) and get a shareable
 * Markdown document or a plain-text strip back. Export semantics:
 * - hidden blocks are skipped, tracked deletions are dropped, insertions kept
 * - images embed their data URL when available (self-contained .md)
 * - tables use GitHub-flavored pipes; the first row becomes the header when
 *   it looks like one (all cells bold or shaded — mirrors the editor heuristic)
 * - formulas render as $..$ / $$..$$, superscript as ^..^, subscript as ~..~
 */

import type { Block, Run, TableCell, TableModel } from './types'

function escapeInline(text: string): string {
  // '|' is only special inside tables, where escapeCell handles it
  return text.replace(/([\\`*_[\]{}<>#+!])/g, '\\$1')
}

/** escape only the pipe char (cells inside a pipe table) */
function escapeCell(text: string): string {
  return text.replace(/\|/g, '\\|').replace(/\n/g, '<br>')
}

function runsToMarkdown(runs: Run[] | undefined): string {
  let out = ''
  for (const run of runs ?? []) {
    if (run.del) continue
    if (run.math) {
      out += `$${escapeInline(run.text)}$`
      continue
    }
    if (run.ruby) {
      out += escapeInline(run.text)
      continue
    }
    if (run.xeTerm !== undefined) continue
    let text = escapeInline(run.text)
    if (run.vertAlign === 'superscript') text = `^${text}^`
    else if (run.vertAlign === 'subscript') text = `~${text}~`
    if (run.highlight) text = `<mark>${text}</mark>`
    if (run.strike) text = `~~${text}~~`
    if (run.underline) text = `<u>${text}</u>`
    if (run.italic) text = `*${text}*`
    if (run.bold) text = `**${text}**`
    if (run.link?.href) text = `[${text}](${escapeInline(run.link.href)})`
    out += text
  }
  return out
}

function runsToPlainText(runs: Run[] | undefined): string {
  let out = ''
  for (const run of runs ?? []) {
    if (run.del) continue
    out += run.text
  }
  return out
}

function hasHeaderRow(model: TableModel): boolean {
  return model.rows.length > 1 && model.rows[0].every((cell) => cell.bold || cell.fill !== undefined)
}

function cellParas(cell: TableCell): string[] {
  if (cell.richParas?.length) {
    return cell.richParas.map((p) => runsToMarkdown(p.runs ?? []))
  }
  return cell.paras
}

function cellPlainParas(cell: TableCell): string[] {
  if (cell.richParas?.length) {
    return cell.richParas.map((p) => runsToPlainText(p.runs ?? []))
  }
  return cell.paras
}

function tableToMarkdown(model: TableModel): string {
  const rows = model.rows.map((row) => row.filter((cell) => cell.vMerge !== 'continue'))
  const columns = Math.max(1, ...rows.map((r) => r.length))
  const header = hasHeaderRow(model)
  const lines: string[] = []
  const rowLine = (row: TableCell[]) => {
    const cells = row.map((cell) => escapeCell(cellParas(cell).join('<br>')))
    while (cells.length < columns) cells.push('')
    return `| ${cells.join(' | ')} |`
  }
  if (rows.length === 0) return ''
  const first = header && rows[0].length > 0 ? rows.shift()! : null
  lines.push(rowLine(first ?? rows[0]))
  const dashes = Array.from({ length: columns }, () => '---')
  lines.push(`| ${dashes.join(' | ')} |`)
  const start = first ? 0 : 1
  for (let i = start; i < rows.length; i++) lines.push(rowLine(rows[i]))
  return lines.join('\n')
}

function tableToPlainText(model: TableModel): string {
  return model.rows
    .map((row) =>
      row
        .filter((cell) => cell.vMerge !== 'continue')
        .map((cell) => cellPlainParas(cell).join(' '))
        .join('\t'),
    )
    .join('\n')
}

/** Serialize the block tree to a GitHub-flavored Markdown document. */
export function serializeBlocksToMarkdown(blocks: Block[]): string {
  const parts: string[] = []
  let prev: Block | null = null
  for (const block of blocks) {
    if (block.hidden) continue
    if (block.blockRevision?.kind === 'del') continue
    let part: string | null = null
    switch (block.type) {
      case 'heading': {
        const level = Math.max(1, Math.min(6, block.level ?? 1))
        part = `${'#'.repeat(level)} ${runsToMarkdown(block.runs)}`
        break
      }
      case 'paragraph': {
        const texts: string[] = []
        if (block.formulaDisplay) texts.push(`$$${block.formulaDisplay.tokens.join(' ')}$$`)
        texts.push(runsToMarkdown(block.runs))
        for (const box of block.textboxes ?? []) {
          for (const para of box.paras) texts.push(`> ${runsToMarkdown(para.runs)}`)
        }
        part = texts.filter(Boolean).join('\n')
        break
      }
      case 'listItem': {
        const indent = '  '.repeat(block.list?.ilvl ?? 0)
        const marker = block.list?.kind === 'ordered' ? '1. ' : '- '
        part = `${indent}${marker}${runsToMarkdown(block.runs)}`
        break
      }
      case 'table':
        part = tableToMarkdown(block.table!)
        break
      case 'image': {
        const alt = escapeInline(block.previewText ?? block.label ?? 'image')
        part = block.imageDataUrl ? `![${alt}](${block.imageDataUrl})` : `![${alt}]()`
        break
      }
      default: {
        // passthrough / protected blocks (charts, OLE, field displays, …)
        part = block.previewText ?? block.label ?? ''
        break
      }
    }
    if (part === null || part === '') continue
    // adjacent list items and tables read as one structure, not separate paragraphs
    const tight = prev && prev.type === block.type && (block.type === 'listItem' || block.type === 'table')
    if (tight) parts[parts.length - 1] += '\n' + part
    else parts.push(part)
    prev = block
  }
  return parts.join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}

/** Serialize the block tree to plain text (formatting stripped, structure kept). */
export function serializeBlocksToPlainText(blocks: Block[]): string {
  const parts: string[] = []
  for (const block of blocks) {
    if (block.hidden) continue
    if (block.blockRevision?.kind === 'del') continue
    switch (block.type) {
      case 'heading':
        parts.push(runsToPlainText(block.runs))
        break
      case 'paragraph': {
        const texts: string[] = [runsToPlainText(block.runs)]
        for (const box of block.textboxes ?? []) {
          for (const para of box.paras) texts.push(runsToPlainText(para.runs))
        }
        parts.push(texts.filter(Boolean).join('\n'))
        break
      }
      case 'listItem':
        parts.push(runsToPlainText(block.runs))
        break
      case 'table':
        parts.push(tableToPlainText(block.table!))
        break
      case 'image':
        parts.push(block.previewText ?? block.label ?? 'image')
        break
      default:
        if (block.previewText) parts.push(block.previewText)
        break
    }
  }
  return parts.join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n'
}
