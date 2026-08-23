import type { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection, type Selection } from '@tiptap/pm/state'
import { queueAnchorRange } from '../editor/aiQueueAnchors'
import { blockIndexRange } from './tools'

/**
 * Selection-scoped AI edit queue (docs parity): the user annotates passages
 * with short instructions, they pile up in the AI panel, and the batch is
 * submitted as one agent run. Anchors live as decorations (aiQueueAnchors.ts),
 * so they migrate through edits character-precisely; block indexes are derived
 * only at render/submit time and never stored.
 */

/** Hard cap on queued edits: keeps one submission inside the agent's turn budget and the card readable */
export const EDIT_QUEUE_MAX = 10
/** Soft cap on one instruction; longer requests belong in the main composer */
export const EDIT_INSTRUCTION_MAX = 500

export interface EditQueueItem {
  qid: string
  instruction: string
  /** text snapshot at annotation time; the label of last resort once the anchor is gone */
  capturedText: string
}

export interface ResolvedQueueItem {
  item: EditQueueItem
  /** null = the anchored text has been deleted since annotation */
  target: { startIndex: number; endIndex: number; excerpt: string } | null
}

export function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/** locate an item's anchor in the current document (block indexes + live excerpt) */
export function resolveQueueItem(editor: Editor, item: EditQueueItem): ResolvedQueueItem {
  const range = queueAnchorRange(editor.state, item.qid)
  if (!range) return { item, target: null }
  const indexes = blockIndexRange(editor.state.doc, range.from, range.to)
  const text = editor.state.doc
    .textBetween(range.from, range.to, '\n', ' ')
    .replace(/\s+/g, ' ')
    .trim()
  // a textless anchor (image/table node selection) is still a live target —
  // label it by block type instead of declaring it orphaned
  const excerpt = text || `(${editor.state.doc.child(indexes.startIndex).type.name} block, no text)`
  return { item, target: { ...indexes, excerpt } }
}

export function resolveQueue(editor: Editor, items: EditQueueItem[]): ResolvedQueueItem[] {
  return items.map((item) => resolveQueueItem(editor, item))
}

/**
 * Selection that focuses an anchor: a text selection cannot cover a block
 * atom (image/math — TextSelection.create throws there), so an anchored
 * atom selects the node itself; TextSelection.between never throws.
 */
export function selectionForAnchor(editor: Editor, qid: string): Selection | null {
  const range = queueAnchorRange(editor.state, qid)
  if (!range) return null
  const doc = editor.state.doc
  const node = doc.nodeAt(range.from)
  if (node && !node.isText && !node.isTextblock && range.to <= range.from + node.nodeSize) {
    return NodeSelection.create(doc, range.from)
  }
  return TextSelection.between(doc.resolve(range.from), doc.resolve(range.to))
}

type LiveItem = ResolvedQueueItem & { target: NonNullable<ResolvedQueueItem['target']> }

export function liveItems(resolved: ResolvedQueueItem[]): LiveItem[] {
  return resolved.filter((r): r is LiveItem => r.target !== null)
}

const blocksLabel = (t: LiveItem['target']): string =>
  t.startIndex === t.endIndex ? `Block ${t.startIndex}` : `Blocks ${t.startIndex}-${t.endIndex}`

/**
 * The batch instruction handed to the model (English regardless of UI
 * language). Edits are listed bottom-up so applying one never shifts the
 * block indexes of those still ahead in the list.
 */
export function buildQueueInstruction(entries: LiveItem[]): string {
  const ordered = [...entries].sort((a, b) => b.target.startIndex - a.target.startIndex)
  const lines = ordered.map(
    (entry, i) =>
      `${i + 1}. ${blocksLabel(entry.target)}, target text: "${truncate(entry.target.excerpt, 160)}"\n` +
      `   Requested change: ${entry.item.instruction}`,
  )
  return [
    'The user marked passages in the document and queued one edit per passage; apply them all now as a single batch.',
    '',
    'Edits to apply (block indexes refer to the current document block list; the list is ordered bottom-up so applying one edit does not shift the indexes of the following ones):',
    ...lines,
    '',
    'Apply exactly these changes to exactly the listed target passages — verify the quoted target text before rewriting, and re-read with get_document_context if block counts changed unexpectedly. Do not modify content outside the listed targets. Finish with a short summary of what was changed.',
  ].join('\n')
}

/** user-facing echo of a submission, shown as the chat bubble text */
export function buildQueueSummary(header: string, entries: LiveItem[]): string {
  const lines = entries.map(
    (entry, i) => `${i + 1}. ${truncate(entry.target.excerpt, 24)} — ${entry.item.instruction}`,
  )
  return [header, ...lines].join('\n')
}
