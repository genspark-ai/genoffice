import type { Editor } from '@tiptap/core'
import { parseDocx, type Block } from '@genoffice/docx-engine'
import { fetchDocBytes } from '../doc-bytes'
import { insertFieldMarks, occurrenceEnds } from './index-entry'

/** one row of a concordance file: the text to look for and the index entry it is marked as */
export interface ConcordanceRow {
  find: string
  entry: string
}

/**
 * The rows of a concordance file, as Word's AutoMark reads one: each row of
 * its tables is the text to find and, in the second column, the entry
 * ("Main:Sub" for a subentry). A paragraph outside a table is a text marked
 * under itself, or two of them when a tab separates them.
 */
export function concordanceRows(blocks: Block[]): ConcordanceRow[] {
  const out: ConcordanceRow[] = []
  const add = (find: string, entry: string) => {
    const text = find.trim()
    if (text) out.push({ find: text, entry: entry.replace(/"/g, '').trim() || text })
  }
  for (const block of blocks) {
    if (block.type === 'table' && block.table) {
      for (const row of block.table.rows) {
        const cells = row.filter((cell) => !cell.gridGap).map((cell) => cell.paras.join(' '))
        add(cells[0] ?? '', cells[1] ?? '')
      }
    } else if (block.type === 'paragraph' || block.type === 'heading') {
      const [find, entry] = (block.runs ?? [])
        .map((r) => r.text)
        .join('')
        .split('\t')
      add(find ?? '', entry ?? '')
    }
  }
  return out
}

/**
 * Marks the first occurrence of each row's text in every paragraph with its
 * entry, capitals and all. An occurrence that already carries the entry is
 * left alone, so marking from the same file twice adds nothing. Returns how
 * many entries were marked.
 */
export function autoMarkIndex(editor: Editor, rows: ConcordanceRow[]): number {
  const { doc } = editor.state
  const marks: Array<{ at: number; attrs: Record<string, unknown> }> = []
  for (const row of rows) {
    for (const at of occurrenceEnds(editor, row.find)) {
      const next = doc.nodeAt(at)
      const marked =
        next?.type.name === 'docXeMark' && !next.attrs.ta && !next.attrs.tc
          ? String(next.attrs.term)
          : null
      if (marked !== row.entry && !marks.some((m) => m.at === at && m.attrs.term === row.entry))
        marks.push({ at, attrs: { term: row.entry, switches: null } })
    }
  }
  insertFieldMarks(editor, marks)
  return marks.length
}

export type AutoMarkResult =
  | { status: 'canceled' | 'password' }
  | { status: 'marked'; count: number }
  | { status: 'failed'; error: string }

/** asks for a concordance file and marks its entries in the document */
export async function pickAndAutoMark(editor: Editor): Promise<AutoMarkResult> {
  const picked = await window.desktop.pickDocx()
  if (!picked) return { status: 'canceled' }
  if ('needsPassword' in picked) return { status: 'password' }
  try {
    const parsed = await parseDocx(await fetchDocBytes(picked.dataUrl))
    return { status: 'marked', count: autoMarkIndex(editor, concordanceRows(parsed.blocks)) }
  } catch (err) {
    return { status: 'failed', error: String(err) }
  }
}
