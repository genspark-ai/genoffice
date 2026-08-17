/**
 * Word-style margin annotations, drawn as an absolute overlay on the page wrap
 * (same pattern as .page-cut-overlays: zero layout height, rebuilt after each
 * pagination remeasure, cleared wholesale).
 *
 * - Comment bubbles: one bubble per open thread in a markup column right of the
 *   paper, top-aligned to its anchor line, stacked downward on overlap, with a
 *   dashed leader from the anchor. The wrap gains `has-markup-area` +
 *   `--markup-w` so the column takes real width (screenshots/regression capture
 *   the wrap box); percentage-centered wrap overlays are re-centered in CSS.
 * - Change bars: vertical segments in the left page margin covering every line
 *   that carries a tracked revision (Word's changed-line marks).
 */
import type { CommentInfo } from '@genoffice/docx-engine'

export const MARKUP_AREA_W = 200
const BUBBLE_W = 168
const BUBBLE_ENTRY_X = 12
const BUBBLE_STACK_GAP = 6
const CHANGE_BAR_X = 24

const REV_SELECTOR =
  '.doc-ins, .doc-del, .has-move-from, .has-move-to, .has-rpr-change, .has-ppr-change'

const SVG_NS = 'http://www.w3.org/2000/svg'

type Seg = { top: number; bottom: number }

function mergeSegs(segs: Seg[]): Seg[] {
  segs.sort((a, b) => a.top - b.top)
  const out: Seg[] = []
  for (const s of segs) {
    const last = out[out.length - 1]
    if (last && s.top <= last.bottom + 3) last.bottom = Math.max(last.bottom, s.bottom)
    else out.push({ ...s })
  }
  return out
}

function flashEls(targets: HTMLElement[]): void {
  targets[0]?.scrollIntoView({ behavior: 'smooth', block: 'center' })
  for (const t of targets) {
    t.classList.remove('doc-comment-flash')
    void t.offsetWidth
    t.classList.add('doc-comment-flash')
  }
}

function flashAnchors(pm: HTMLElement, id: string): void {
  flashEls(
    [...pm.querySelectorAll<HTMLElement>('.doc-comment')].filter((s) =>
      (s.dataset.commentIds ?? '').split(' ').includes(id),
    ),
  )
}

function makeBubble(root: CommentInfo, replies: CommentInfo[], onJump: () => void): HTMLElement {
  const bubble = document.createElement('div')
  bubble.className = 'comment-bubble'
  const addEntry = (c: CommentInfo, cls?: string) => {
    const entry = document.createElement('div')
    if (cls) entry.className = cls
    const author = document.createElement('div')
    author.className = 'comment-bubble-author'
    author.textContent = c.author
    const text = document.createElement('div')
    text.className = 'comment-bubble-text'
    text.textContent = c.text
    entry.append(author, text)
    bubble.appendChild(entry)
  }
  addEntry(root)
  for (const r of replies) addEntry(r, 'comment-bubble-reply')
  bubble.addEventListener('click', onJump)
  return bubble
}

export function clearMarginAnnotations(wrap: HTMLElement): void {
  wrap.querySelector(':scope > .page-margin-annotations')?.remove()
  wrap.classList.remove('has-markup-area')
}

/** parsed block subset for anchoring comments that never produced a text mark */
export interface AnchorBlock {
  docxIndex?: number | null
  originalXml?: string | null
}

export function syncMarginAnnotations(
  wrap: HTMLElement,
  pm: HTMLElement,
  comments: CommentInfo[],
  zoomFactor: number,
  blocks?: AnchorBlock[],
): void {
  const f = zoomFactor
  const wrapRect = wrap.getBoundingClientRect()
  const pmRect = pm.getBoundingClientRect()

  // change bars only in All Markup view (Word hides them in No Markup / Original)
  const segs: Seg[] = []
  if (!wrap.closest('.rev-display-none, .rev-display-original')) {
    for (const el of pm.querySelectorAll<HTMLElement>(REV_SELECTOR)) {
      for (const r of el.getClientRects()) {
        if (r.height > 0)
          segs.push({ top: (r.top - wrapRect.top) / f, bottom: (r.bottom - wrapRect.top) / f })
      }
    }
  }
  const bars = mergeSegs(segs)

  // one bubble per open top-level thread that has an anchor range in the body
  const anchorOf = new Map<string, HTMLElement>()
  for (const span of pm.querySelectorAll<HTMLElement>('.doc-comment')) {
    for (const id of (span.dataset.commentIds ?? '').split(' ')) {
      if (id && !anchorOf.has(id)) anchorOf.set(id, span)
    }
  }
  type Thread = {
    root: CommentInfo
    replies: CommentInfo[]
    top: number
    y: number
    x: number
    jump: () => void
  }
  const threads: Thread[] = []
  const localThread = (c: CommentInfo, rect: DOMRect, blockAnchor: HTMLElement | null): Thread => ({
    root: c,
    replies: comments.filter((r) => r.parentId === c.id),
    top: (rect.top - wrapRect.top) / f,
    // leader start: end of the marked range, or start of the anchor paragraph's
    // first line for range-less comments (bare w:commentReference on an empty run)
    y: blockAnchor
      ? (rect.top - wrapRect.top) / f + Math.min(rect.height / f, 16)
      : (rect.bottom - wrapRect.top) / f - 1,
    x: ((blockAnchor ? rect.left : rect.right) - wrapRect.left) / f,
    // block-anchored threads have no .doc-comment span: flash the block itself
    jump: blockAnchor ? () => flashEls([blockAnchor]) : () => flashAnchors(pm, c.id),
  })
  for (const c of comments) {
    if (c.parentId || c.done) continue
    const el = anchorOf.get(c.id)
    const rect = el ? [...el.getClientRects()].find((r) => r.height > 0) : undefined
    if (rect) {
      threads.push(localThread(c, rect, null))
      continue
    }
    // no mark in the body: anchor to the block whose XML carries the reference
    const idRe = new RegExp(
      `<w:comment(?:Reference|RangeStart)\\b[^>]*w:id="${c.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`,
    )
    const block = blocks?.find(
      (b) => b.docxIndex != null && b.originalXml && idRe.test(b.originalXml),
    )
    const blockEl =
      block && (pm.querySelector(`[data-idx="${block.docxIndex}"]`) as HTMLElement | null)
    const blockRect = blockEl && [...blockEl.getClientRects()].find((r) => r.height > 0)
    if (blockEl && blockRect) threads.push(localThread(c, blockRect, blockEl))
  }
  threads.sort((a, b) => a.top - b.top)

  if (bars.length === 0 && threads.length === 0) {
    clearMarginAnnotations(wrap)
    return
  }

  let layer = wrap.querySelector(':scope > .page-margin-annotations') as HTMLElement | null
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-margin-annotations'
    wrap.appendChild(layer)
  }
  layer.textContent = ''

  wrap.classList.toggle('has-markup-area', threads.length > 0)
  if (threads.length > 0) wrap.style.setProperty('--markup-w', `${MARKUP_AREA_W}px`)

  const barX = (pmRect.left - wrapRect.left) / f + CHANGE_BAR_X
  for (const b of bars) {
    const el = document.createElement('div')
    el.className = 'change-bar'
    el.style.left = `${barX}px`
    el.style.top = `${b.top}px`
    el.style.height = `${b.bottom - b.top}px`
    layer.appendChild(el)
  }

  if (threads.length === 0) return
  const paperRight = (pmRect.right - wrapRect.left) / f
  const bubbleLeft = paperRight + BUBBLE_ENTRY_X
  const svg = document.createElementNS(SVG_NS, 'svg')
  svg.setAttribute('class', 'comment-leaders')
  layer.appendChild(svg)
  let prevBottom = -Infinity
  for (const t of threads) {
    const bubble = makeBubble(t.root, t.replies, t.jump)
    bubble.style.left = `${bubbleLeft}px`
    bubble.style.width = `${BUBBLE_W}px`
    const top = Math.max(t.top, prevBottom + BUBBLE_STACK_GAP)
    bubble.style.top = `${top}px`
    layer.appendChild(bubble)
    prevBottom = top + bubble.offsetHeight
    const path = document.createElementNS(SVG_NS, 'path')
    path.setAttribute(
      'd',
      `M ${t.x} ${t.y} L ${paperRight + BUBBLE_ENTRY_X / 2} ${t.y} L ${bubbleLeft} ${top + 9}`,
    )
    svg.appendChild(path)
  }
}
