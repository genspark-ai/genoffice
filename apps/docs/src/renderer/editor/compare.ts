/**
 * Review → Compare: paragraph-level diff of two documents (pure, unit-testable).
 * Word builds a merged revision document; this lightweight version reports
 * per-paragraph additions / removals / edits for a side panel.
 */
import type { Block } from '@genoffice/docx-engine'
import { inlineToRuns, type PmNode } from './convert'

export interface CompareEntry {
  kind: 'same' | 'removed' | 'added' | 'changed'
  /** paragraph text in the current document */
  left?: string
  /** paragraph text in the compared document */
  right?: string
}

/** visible block -> comparable plain text (tables/objects fall back to previews) */
export function blockTexts(blocks: Block[]): string[] {
  return blocks
    .filter((b) => !b.hidden)
    .map((b) => {
      if (b.runs) return b.runs.map((r) => r.text).join('')
      return b.previewText ?? ''
    })
}

/** Comparable text from the live editor, including paragraphs not yet saved to the DOCX. */
export function editorBlockTexts(doc: PmNode, originalBlocks: Block[]): string[] {
  const originalByIndex = new Map(
    originalBlocks
      .filter((block) => block.docxIndex !== null)
      .map((block) => [block.docxIndex, block]),
  )
  return (doc.content ?? []).map((node) => {
    if (node.type === 'docProtected') return String(node.attrs?.previewText ?? '')
    if (node.type === 'docTable') {
      const original = originalByIndex.get(node.attrs?.docxIndex as number)
      return original?.previewText ?? ''
    }
    return inlineToRuns(node.content ?? [])
      .map((run) => run.text)
      .join('')
  })
}

/** LCS-based paragraph diff; a removal directly followed by an addition merges into 'changed' */
export function compareParagraphs(left: string[], right: string[]): CompareEntry[] {
  const n = left.length
  const m = right.length
  // lcs[i][j] = LCS length of left[i:], right[j:]
  const lcs: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] =
        left[i] === right[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1])
    }
  }
  const raw: CompareEntry[] = []
  let i = 0
  let j = 0
  while (i < n && j < m) {
    if (left[i] === right[j]) {
      raw.push({ kind: 'same', left: left[i], right: right[j] })
      i++
      j++
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      raw.push({ kind: 'removed', left: left[i] })
      i++
    } else {
      raw.push({ kind: 'added', right: right[j] })
      j++
    }
  }
  while (i < n) raw.push({ kind: 'removed', left: left[i++] })
  while (j < m) raw.push({ kind: 'added', right: right[j++] })

  const out: CompareEntry[] = []
  for (const entry of raw) {
    const prev = out[out.length - 1]
    if (entry.kind === 'added' && prev?.kind === 'removed' && prev.right === undefined) {
      prev.kind = 'changed'
      prev.right = entry.right
      continue
    }
    out.push(entry)
  }
  return out
}

export interface CompareSummary {
  added: number
  removed: number
  changed: number
}

export function summarize(entries: CompareEntry[]): CompareSummary {
  return {
    added: entries.filter((e) => e.kind === 'added').length,
    removed: entries.filter((e) => e.kind === 'removed').length,
    changed: entries.filter((e) => e.kind === 'changed').length,
  }
}
