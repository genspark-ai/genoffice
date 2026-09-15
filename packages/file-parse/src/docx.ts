import { parseDocx } from '@genoffice/docx-engine'

/** flatten a parsed docx into readable text (structure markers preserved) */
export async function docxToText(bytes: Uint8Array): Promise<string> {
  const doc = await parseDocx(new Uint8Array(bytes))
  const lines: string[] = []
  for (const block of doc.blocks) {
    if (block.hidden) continue
    switch (block.type) {
      case 'heading':
        lines.push(`${'#'.repeat(Math.min(block.level ?? 1, 6))} ${runText(block)}`)
        break
      case 'listItem':
        lines.push(`- ${runText(block)}`)
        break
      case 'paragraph':
        lines.push(runText(block))
        break
      case 'table': {
        const rows = block.table?.rows ?? []
        for (const row of rows) {
          lines.push(row.map((cell) => cell.paras.join(' ')).join(' | '))
        }
        break
      }
      default:
        if (block.textboxes?.length) {
          for (const box of block.textboxes) {
            for (const para of box.paras) lines.push(para.runs.map((r) => r.text).join(''))
          }
        } else if (block.previewText) {
          lines.push(block.previewText)
        }
    }
  }
  // Footnotes/endnotes live outside the block tree but are part of the
  // document's readable text (and what AI Q&A over attachments should index).
  if (doc.footnotes.length > 0) {
    lines.push('## Footnotes')
    for (const note of doc.footnotes) lines.push(note.text)
  }
  if (doc.endnotes.length > 0) {
    lines.push('## Endnotes')
    for (const note of doc.endnotes) lines.push(note.text)
  }
  return lines.join('\n')
}

function runText(block: { runs?: { text: string }[] }): string {
  return (block.runs ?? []).map((r) => r.text).join('')
}
