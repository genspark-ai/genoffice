import { sourceText, type ElementEntry, type ParseMap } from '../document/parse-map'

/**
 * Element-scoped AI edit queue (slides/docs parity): the user annotates
 * selected elements with short instructions, they pile up in the AI panel,
 * and the batch is submitted as one agent run. Items anchor to sids, which
 * the parse map keeps stable across rebuilds by tag + path; an item whose
 * element is gone stays visible as stale until removed.
 */

/** hard cap on queued edits: keeps one submission inside the agent's turn budget and the card readable */
export const EDIT_QUEUE_MAX = 10
/** soft cap on one instruction; longer requests belong in the main composer */
export const EDIT_INSTRUCTION_MAX = 500

export interface EditQueueItem {
  qid: string
  sid: number
  tag: string
  /** excerpt at annotation time; the label of last resort once the element is gone */
  capturedText: string
  instruction: string
}

export interface QueueTarget {
  sid: number
  tag: string
  excerpt: string
  /** source offset of the start tag, for bottom-up ordering */
  start: number
}

export interface ResolvedQueueItem {
  item: EditQueueItem
  /** null = the element no longer exists in the source */
  target: QueueTarget | null
}

export type LiveQueueItem = ResolvedQueueItem & { target: QueueTarget }

export function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/** collapsed text content, or the tag when the element carries none (img, hr, empty div) */
export function excerptOf(text: string, entry: ElementEntry): string {
  const content = sourceText(text, entry).trim()
  return content || `<${entry.tag}>`
}

export function resolveQueueItem(
  text: string,
  map: ParseMap,
  item: EditQueueItem,
): ResolvedQueueItem {
  const entry = map.bySid.get(item.sid)
  if (!entry || entry.tag !== item.tag) return { item, target: null }
  return {
    item,
    target: {
      sid: entry.sid,
      tag: entry.tag,
      excerpt: excerptOf(text, entry),
      start: entry.range[0],
    },
  }
}

export function resolveQueue(
  text: string,
  map: ParseMap,
  items: EditQueueItem[],
): ResolvedQueueItem[] {
  return items.map((item) => resolveQueueItem(text, map, item))
}

export function liveItems(resolved: ResolvedQueueItem[]): LiveQueueItem[] {
  return resolved.filter((r): r is LiveQueueItem => r.target !== null)
}

/** "Send now" from the popover: the instruction pinned to its element so a later selection change cannot redirect it */
export function buildSelectionInstruction(target: QueueTarget, instruction: string): string {
  return `Apply this to the element sid=${target.sid} <${target.tag}> ("${truncate(target.excerpt, 120)}"), addressing it by sid in apply_ops:\n${instruction}`
}

/**
 * The batch instruction handed to the model (English regardless of UI
 * language). Edits are listed bottom-up so applying one never shifts the
 * source of those still ahead in the list.
 */
export function buildQueueInstruction(entries: LiveQueueItem[]): string {
  const ordered = [...entries].sort((a, b) => b.target.start - a.target.start)
  const lines = ordered.map(
    (entry, i) =>
      `${i + 1}. sid=${entry.target.sid} <${entry.target.tag}> "${truncate(entry.target.excerpt, 160)}"\n` +
      `   Requested change: ${entry.item.instruction}`,
  )
  return [
    'The user marked elements in the page preview and queued one edit per element; apply them all now as a single batch.',
    '',
    'Edits to apply (sid = stable element id from the outline; the list is ordered bottom-up so applying one edit does not disturb the following ones):',
    ...lines,
    '',
    'Apply exactly these changes to exactly the listed elements — address each by its sid in apply_ops (read_source sid=N first when you need the exact markup) and batch them into as few apply_ops calls as possible. Do not modify content outside the listed elements. Finish with a short summary of what was changed.',
  ].join('\n')
}

/** user-facing echo of a submission, shown as the chat bubble text */
export function buildQueueSummary(header: string, entries: LiveQueueItem[]): string {
  const lines = entries.map(
    (entry, i) => `${i + 1}. ${truncate(entry.target.excerpt, 24)} — ${entry.item.instruction}`,
  )
  return [header, ...lines].join('\n')
}
