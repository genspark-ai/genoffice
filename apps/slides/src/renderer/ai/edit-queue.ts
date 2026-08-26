/**
 * Element-scoped AI edit queue: the user annotates individual elements with a
 * short instruction, the annotations pile up in the AI panel, and the batch is
 * submitted in one action.
 *
 * Anchors are ids only — never coordinates, and never a page index treated as
 * authority. A RenderSlide carries no id of its own, so a stored slideIndex
 * breaks as soon as a page is deleted or reordered; it is kept as a cache and
 * re-derived by searching the deck for the anchor (see resolveQueueItem).
 * Staleness is decided at submit time rather than stored, so undo/redo of a
 * delete heals itself.
 */
import type {
  GroupRenderNode,
  RenderNode,
  RenderNodeType,
  RenderSlide,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'
import type { StringKey } from '../i18n/locale'

/** Localized noun for an element kind, shared by the popover title and the queue rows */
export const NODE_NOUN_KEY: Record<RenderNodeType, StringKey> = {
  text: 'aiElText',
  shape: 'aiElShape',
  picture: 'aiElPicture',
  table: 'aiElTable',
  chart: 'aiElChart',
  group: 'aiElGroup',
  'placeholder-chip': 'aiElShape',
}

/** Hard cap on queued edits: keeps one submission inside the agent's turn budget and the card readable */
export const EDIT_QUEUE_MAX = 10
/** Soft cap on one instruction; longer requests belong in the main composer */
export const EDIT_INSTRUCTION_MAX = 500

export type EditQueueStatus = 'pending' | 'running' | 'done' | 'failed' | 'skipped'

export interface EditTargetRef {
  /** durableId ?? sourceId — the id space the deck outline and the edit tools speak */
  id: string
  /** Render-tree id: deck-unique for as long as the app stays open, so it disambiguates first */
  sourceId: string
  type: RenderNodeType
}

export interface EditQueueItem {
  key: string
  /** Cache/hint only; resolveQueueItem refreshes it */
  slideIndex: number
  /** More than one when the annotation was made from a multi-selection */
  targets: EditTargetRef[]
  instruction: string
  status: EditQueueStatus
  error?: string
}

/** Structured element summary; the UI localizes it, the prompt uses promptLabel */
export interface NodeDescriptor {
  type: RenderNodeType
  /** Trimmed first line of the element's text, when it has any */
  text?: string
  rows?: number
  cols?: number
}

/** Durable when the element's bytes carry one, so the anchor survives regenerate/ungroup/save */
export function anchorId(node: RenderNode): string {
  return node.durableId ?? node.sourceId
}

function plainText(node: RenderNode): string {
  if (node.type === 'shape' || node.type === 'text') {
    return ((node as ShapeRenderNode).text?.lines ?? [])
      .map((line) => line.runs.map((r) => r.text).join(''))
      .join(' ')
      .trim()
  }
  if (node.type === 'table') {
    return (
      (node as TableRenderNode).cells
        .map((c) => (c.text?.lines ?? []).map((l) => l.runs.map((r) => r.text).join('')).join(' '))
        .find((t) => t.trim().length > 0)
        ?.trim() ?? ''
    )
  }
  return ''
}

export function describeNode(node: RenderNode): NodeDescriptor {
  const text = plainText(node).replace(/\s+/g, ' ')
  const desc: NodeDescriptor = { type: node.type }
  if (text) desc.text = text
  if (node.type === 'table') {
    // gridX/gridY are line offsets: nCols+1 and nRows+1 entries
    const table = node as TableRenderNode
    desc.rows = Math.max(0, table.gridY.length - 1)
    desc.cols = Math.max(0, table.gridX.length - 1)
  }
  return desc
}

export function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/** English one-liner handed to the model (the prompt is English regardless of UI language) */
export function promptLabel(desc: NodeDescriptor): string {
  const kind =
    desc.type === 'table' && desc.rows && desc.cols
      ? `${desc.rows}x${desc.cols} table`
      : desc.type === 'text'
        ? 'text box'
        : desc.type
  return desc.text ? `${kind}, "${truncate(desc.text, 60)}"` : kind
}

/**
 * Freeze a selection-scoped "Send now" request into the model instruction.
 * The popover already resolved canvas source ids to durable ids; carrying those
 * refs here prevents a later live-selection read from substituting a nearby
 * element while the AI panel opens.
 */
export function buildSelectionInstruction(
  slideIndex: number,
  targets: ReadonlyArray<{ id: string; desc: NodeDescriptor }>,
  instruction: string,
): string {
  return [
    'The user invoked "Send now" for a fixed canvas selection.',
    `Apply the request on page ${slideIndex + 1} (slideIndex=${slideIndex}) to exactly these elements:`,
    ...targets.map((target) => `- ${target.id} (${promptLabel(target.desc)})`),
    '',
    `Requested change: ${instruction}`,
    '',
    'Use the listed ids directly. Do not infer or substitute a nearby element, and do not modify unlisted elements.',
  ].join('\n')
}

interface FoundNode {
  node: RenderNode
  /** 0 = top level, 1 = child of a top-level group, 2+ = deeper (edit tools patch one level only) */
  depth: number
}

/**
 * Element kind is part of every match, not a nicety: decks authored without
 * a16:creationId fall back to a durable id that is unique per slide but
 * repeats across slides, so an id alone can name an unrelated element on
 * another page. Requiring the kind to agree makes a mismatch fail closed.
 */
function findNode(
  nodes: RenderNode[],
  target: EditTargetRef,
  by: 'source' | 'durable',
  depth = 0,
): FoundNode | null {
  for (const n of nodes) {
    const idHit = by === 'source' ? n.sourceId === target.sourceId : n.durableId === target.id
    if (idHit && n.type === target.type) return { node: n, depth }
    if (n.type === 'group') {
      const hit = findNode((n as GroupRenderNode).children, target, by, depth + 1)
      if (hit) return hit
    }
  }
  return null
}

export type ResolveFailure = 'deleted' | 'ambiguous' | 'nested'

export type ResolvedItem =
  | { ok: true; item: EditQueueItem; slideIndex: number; nodes: RenderNode[] }
  | { ok: false; item: EditQueueItem; reason: ResolveFailure }

function findAllOnSlide(
  slide: RenderSlide,
  targets: EditTargetRef[],
  by: 'source' | 'durable',
): FoundNode[] | null {
  const found: FoundNode[] = []
  for (const t of targets) {
    const hit = findNode(slide.nodes, t, by)
    if (!hit) return null
    found.push(hit)
  }
  return found
}

function scanDeck(
  slides: RenderSlide[],
  targets: EditTargetRef[],
  by: 'source' | 'durable',
): { index: number; found: FoundNode[] }[] {
  const hits: { index: number; found: FoundNode[] }[] = []
  slides.forEach((slide, index) => {
    const found = findAllOnSlide(slide, targets, by)
    if (found) hits.push({ index, found })
  })
  return hits
}

/**
 * Locate an item's targets in the current deck, page hint included but never
 * trusted: the whole deck is searched, so page reordering, insertion, and
 * deletion all resolve correctly. Render-tree ids come first because they are
 * deck-unique; durable ids are the fallback for elements that re-materialized
 * (AI regeneration, ungroup) and lost their session id. A durable match on
 * several pages is reported as ambiguous rather than guessed.
 */
export function resolveQueueItem(slides: RenderSlide[], item: EditQueueItem): ResolvedItem {
  let hits = scanDeck(slides, item.targets, 'source')
  if (hits.length === 0) hits = scanDeck(slides, item.targets, 'durable')
  if (hits.length === 0) return { ok: false, item, reason: 'deleted' }
  if (hits.length > 1) {
    // Prefer the page the annotation was made on when it is among the candidates
    const cached = hits.find((h) => h.index === item.slideIndex)
    if (!cached) return { ok: false, item, reason: 'ambiguous' }
    hits = [cached]
  }
  const { index, found } = hits[0]!
  if (found.some((f) => f.depth > 1)) return { ok: false, item, reason: 'nested' }
  return { ok: true, item, slideIndex: index, nodes: found.map((f) => f.node) }
}

export function resolveQueue(slides: RenderSlide[], items: EditQueueItem[]): ResolvedItem[] {
  return items.map((item) => resolveQueueItem(slides, item))
}

/** Same-page edits ship together (they affect each other's layout); pages run one after another */
export function groupByPage(
  resolved: ResolvedItem[],
): { slideIndex: number; entries: Extract<ResolvedItem, { ok: true }>[] }[] {
  const byPage = new Map<number, Extract<ResolvedItem, { ok: true }>[]>()
  for (const r of resolved) {
    if (!r.ok) continue
    const list = byPage.get(r.slideIndex) ?? []
    list.push(r)
    byPage.set(r.slideIndex, list)
  }
  return [...byPage.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([slideIndex, entries]) => ({ slideIndex, entries }))
}

function describeEntry(entry: Extract<ResolvedItem, { ok: true }>): string {
  const targets = entry.nodes
    .map((n) => `${anchorId(n)} (${promptLabel(describeNode(n))})`)
    .join('; ')
  return `- ${targets}\n  Requested change: ${entry.item.instruction}`
}

/**
 * One page's share of the batch. The whole batch is listed as background so the
 * agent keeps global intent in view while being told to touch only this page —
 * the remaining pages are handled by their own runs.
 */
export function buildPageInstruction(
  group: { slideIndex: number; entries: Extract<ResolvedItem, { ok: true }>[] },
  totalPages: number,
  pageOrdinal: number,
): string {
  const lines = [
    `The user marked specific elements and submitted them as one batch of edits.`,
    `This request covers page ${group.slideIndex + 1} (slideIndex=${group.slideIndex}) only` +
      (totalPages > 1 ? `, part ${pageOrdinal} of ${totalPages} in the batch.` : '.'),
    '',
    'Edits to apply, each naming the element it targets:',
    ...group.entries.map(describeEntry),
    '',
    'Apply exactly these changes to exactly these elements. Read other elements on',
    'the page when alignment or consistency requires it, but do not modify anything',
    'outside the listed targets, and do not touch other pages — they have their own',
    'requests in this batch.',
  ]
  return lines.join('\n')
}

/** Short user-facing echo of a submission, shown as the chat bubble text */
export function buildSubmissionSummary(
  resolved: ResolvedItem[],
  label: (slideIndex: number, instruction: string) => string,
): string {
  return resolved
    .filter((r): r is Extract<ResolvedItem, { ok: true }> => r.ok)
    .map((r) => label(r.slideIndex, r.item.instruction))
    .join('\n')
}
