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
 * - Revision bubbles (balloon mode): "Deleted:"/"Formatted:" balloons for
 *   tracked deletions and format changes, which leave the text flow in that
 *   mode (Word's All Markup balloons). Bubbles stack per page and overflow
 *   past a page bottom is dropped, like Word's overflow pane.
 * - Change bars: vertical segments in the left page margin covering every line
 *   that carries a tracked revision (Word's changed-line marks).
 */
import type { CommentInfo } from '@genoffice/docx-engine'
import type { EditorView } from '@tiptap/pm/view'
import type { Mark as PmMark } from '@tiptap/pm/model'
import { t } from '../i18n/locale'

export const MARKUP_AREA_W = 200
const BUBBLE_W = 168
const BUBBLE_ENTRY_X = 12
const BUBBLE_STACK_GAP = 6
const CHANGE_BAR_X = 24
const REV_TEXT_MAX = 220

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

function makeRevBubble(kind: 'del' | 'fmt', text: string): HTMLElement {
  const bubble = document.createElement('div')
  bubble.className = `comment-bubble rev-bubble rev-bubble-${kind}`
  const label = document.createElement('span')
  label.className = 'rev-bubble-label'
  label.textContent = `${t(kind === 'del' ? 'editorRevDeleted' : 'editorRevFormatted')}: `
  const body = document.createElement('span')
  body.className = 'rev-bubble-text'
  body.textContent = text.length > REV_TEXT_MAX ? `${text.slice(0, REV_TEXT_MAX)}…` : text
  bubble.append(label, body)
  return bubble
}

/** "Formatted:" balloon text: the changed run props' new values (Word describes what the run became) */
function fmtDesc(old: Record<string, unknown>, marks: readonly PmMark[]): string {
  const ts = marks.find((m) => m.type.name === 'docTextStyle')
  const bold = marks.some((m) => m.type.name === 'bold')
  const italic = marks.some((m) => m.type.name === 'italic')
  const parts: string[] = []
  const fontBits: string[] = []
  const newFont = (ts?.attrs.fontAscii ?? ts?.attrs.font) as string | null | undefined
  const oldFont = (old.fontAscii ?? old.font) as string | undefined
  if (newFont && newFont !== oldFont) fontBits.push(newFont)
  const newSz = ts?.attrs.sizeHalfPoints as number | null | undefined
  if (newSz && newSz !== old.sizeHalfPoints) fontBits.push(`${newSz / 2} pt`)
  if (fontBits.length > 0) parts.push(`${t('editorRevFont')}: ${fontBits.join(', ')}`)
  if (bold !== !!old.bold) parts.push(t(bold ? 'editorRevBold' : 'editorRevNotBold'))
  if (italic !== !!old.italic) parts.push(t(italic ? 'editorRevItalic' : 'editorRevNotItalic'))
  const newHl = (ts?.attrs.highlight as string | null | undefined) ?? null
  if (newHl !== (((old.highlight as string | undefined) ?? null) as string | null)) {
    parts.push(t('editorRevHighlight'))
  }
  return parts.join(', ') || t('editorFormatRevision')
}

type RevItem = { kind: 'del' | 'fmt'; from: number; to: number; text: string }

/**
 * Tracked revisions that live in balloons, in document order, merged over
 * contiguous positions. A deleted paragraph mark bridges [end-1, end+1], so a
 * fully deleted paragraph chains with the next paragraph's leading deletion
 * into one balloon, matching Word's "XXX¶XXX¶" grouping.
 */
function revGroupsOf(view: EditorView): RevItem[] {
  const items: RevItem[] = []
  view.state.doc.descendants((node, pos) => {
    if (node.isTextblock) {
      if (node.attrs?.pPrChange) {
        items.push({ kind: 'fmt', from: pos + 1, to: pos + 1, text: t('editorFormatRevision') })
      }
      if (node.attrs?.paraMarkDel) {
        const end = pos + node.nodeSize
        items.push({ kind: 'del', from: end - 1, to: end + 1, text: '¶' })
      }
      return true
    }
    if (!node.isText) return true
    if (node.marks.some((m) => m.type.name === 'del')) {
      items.push({ kind: 'del', from: pos, to: pos + node.nodeSize, text: node.text ?? '' })
      return true
    }
    const rpr = node.marks.find((m) => m.type.name === 'rprChange')
    if (rpr) {
      items.push({
        kind: 'fmt',
        from: pos,
        to: pos + node.nodeSize,
        text: fmtDesc((rpr.attrs.old ?? {}) as Record<string, unknown>, node.marks),
      })
    }
    return true
  })
  items.sort((a, b) => a.from - b.from || a.to - b.to)
  const groups: RevItem[] = []
  for (const it of items) {
    const last = groups[groups.length - 1]
    if (
      last &&
      last.kind === it.kind &&
      it.from <= last.to &&
      (it.kind === 'del' || it.text === last.text)
    ) {
      last.to = Math.max(last.to, it.to)
      if (it.kind === 'del') last.text += it.text
    } else {
      groups.push({ ...it })
    }
  }
  return groups
}

/** Nearest on-screen box before a hidden position: previous visible sibling, walking up. */
function visibleRectNear(el: HTMLElement | null, pm: HTMLElement): DOMRect | null {
  let cur: HTMLElement | null = el
  while (cur && cur !== pm) {
    for (
      let sib = cur.previousElementSibling as HTMLElement | null;
      sib;
      sib = sib.previousElementSibling as HTMLElement | null
    ) {
      const rects = sib.getClientRects()
      if (rects.length > 0) return rects[rects.length - 1]
    }
    cur = cur.parentElement
  }
  return null
}

function anchorPointFor(
  view: EditorView,
  pm: HTMLElement,
  pos: number,
): { top: number; bottom: number; left: number } | null {
  const clamped = Math.max(0, Math.min(pos, view.state.doc.content.size))
  try {
    const c = view.coordsAtPos(clamped)
    if (c && (c.top !== 0 || c.left !== 0 || c.bottom !== 0)) {
      return { top: c.top, bottom: c.bottom, left: c.left }
    }
  } catch {
    /* position renders inside hidden (balloon-collapsed) content */
  }
  try {
    const { node } = view.domAtPos(clamped)
    const el = (node instanceof HTMLElement ? node : node.parentElement) as HTMLElement | null
    const rect = visibleRectNear(el, pm)
    if (rect) return { top: rect.bottom, bottom: rect.bottom, left: rect.right }
  } catch {
    /* fall through */
  }
  return null
}

/** Page bands (wrap coords) from the top-level page-gap widgets; one whole-doc band when absent. */
function pageBandsOf(pm: HTMLElement, wrapRect: DOMRect, wrapH: number, f: number): Seg[] {
  const gaps = [...pm.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-inline)')]
    .map((g) => g.getBoundingClientRect())
    .filter((r) => r.height > 0)
    .sort((a, b) => a.top - b.top)
  if (gaps.length === 0) return [{ top: 0, bottom: wrapH }]
  const bands: Seg[] = []
  let top = 0
  for (const g of gaps) {
    bands.push({ top, bottom: (g.top - wrapRect.top) / f })
    top = (g.bottom - wrapRect.top) / f
  }
  bands.push({ top, bottom: wrapH })
  return bands
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
  view?: EditorView,
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
  // one bubble per open top-level thread that has an anchor range in the body
  const anchorOf = new Map<string, HTMLElement>()
  for (const span of pm.querySelectorAll<HTMLElement>('.doc-comment')) {
    for (const id of (span.dataset.commentIds ?? '').split(' ')) {
      if (id && !anchorOf.has(id)) anchorOf.set(id, span)
    }
  }
  type Placed = {
    top: number
    y: number
    x: number
    /** comments always render; only revision balloons overflow-drop (Word's overflow pane) */
    sticky: boolean
    make: () => HTMLElement
  }
  const placed: Placed[] = []
  const localThread = (c: CommentInfo, rect: DOMRect, blockAnchor: HTMLElement | null): Placed => ({
    sticky: true,
    top: (rect.top - wrapRect.top) / f,
    // leader start: end of the marked range, or start of the anchor paragraph's
    // first line for range-less comments (bare w:commentReference on an empty run)
    y: blockAnchor
      ? (rect.top - wrapRect.top) / f + Math.min(rect.height / f, 16)
      : (rect.bottom - wrapRect.top) / f - 1,
    x: ((blockAnchor ? rect.left : rect.right) - wrapRect.left) / f,
    make: () =>
      makeBubble(
        c,
        comments.filter((r) => r.parentId === c.id),
        // block-anchored threads have no .doc-comment span: flash the block itself
        blockAnchor ? () => flashEls([blockAnchor]) : () => flashAnchors(pm, c.id),
      ),
  })
  for (const c of comments) {
    if (c.parentId || c.done) continue
    const el = anchorOf.get(c.id)
    const rect = el ? [...el.getClientRects()].find((r) => r.height > 0) : undefined
    if (rect) {
      placed.push(localThread(c, rect, null))
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
    if (blockEl && blockRect) placed.push(localThread(c, blockRect, blockEl))
  }

  // revision balloons: only in balloon mode (print view, All Markup), where the
  // deleted text / format chips have left the flow
  if (view && wrap.closest('.rev-balloon')) {
    for (const g of revGroupsOf(view)) {
      const p = anchorPointFor(view, pm, g.from)
      if (!p) continue
      // hidden deletions have no rects for the bar pass above: mark their anchor
      // line (the sibling-fallback anchor is a point — give it one line of bar)
      if (g.kind === 'del') {
        const top = (p.top - wrapRect.top) / f
        segs.push({ top, bottom: Math.max((p.bottom - wrapRect.top) / f, top + 12) })
      }
      placed.push({
        sticky: false,
        top: (p.top - wrapRect.top) / f,
        y: (p.bottom - wrapRect.top) / f - 1,
        x: (p.left - wrapRect.left) / f,
        make: () => makeRevBubble(g.kind, g.text),
      })
    }
  }
  placed.sort((a, b) => a.top - b.top)
  const bars = mergeSegs(segs)

  if (bars.length === 0 && placed.length === 0) {
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

  wrap.classList.toggle('has-markup-area', placed.length > 0)
  if (placed.length > 0) wrap.style.setProperty('--markup-w', `${MARKUP_AREA_W}px`)

  const barX = (pmRect.left - wrapRect.left) / f + CHANGE_BAR_X
  for (const b of bars) {
    const el = document.createElement('div')
    el.className = 'change-bar'
    el.style.left = `${barX}px`
    el.style.top = `${b.top}px`
    el.style.height = `${b.bottom - b.top}px`
    layer.appendChild(el)
  }

  if (placed.length === 0) return
  const paperRight = (pmRect.right - wrapRect.left) / f
  const bubbleLeft = paperRight + BUBBLE_ENTRY_X
  const svg = document.createElementNS(SVG_NS, 'svg')
  svg.setAttribute('class', 'comment-leaders')
  layer.appendChild(svg)
  // stack per page band; a bubble that cannot start within its band is dropped
  // (Word routes those to the overflow pane)
  const bands = pageBandsOf(pm, wrapRect, wrapRect.height / f, f)
  let bandIdx = 0
  let prevBottom = -Infinity
  for (const p of placed) {
    while (bandIdx < bands.length - 1 && p.top >= bands[bandIdx + 1].top) {
      bandIdx++
      prevBottom = -Infinity
    }
    const bubble = p.make()
    bubble.style.left = `${bubbleLeft}px`
    bubble.style.width = `${BUBBLE_W}px`
    const top = Math.max(p.top, prevBottom + BUBBLE_STACK_GAP, bands[bandIdx].top)
    bubble.style.top = `${top}px`
    layer.appendChild(bubble)
    if (
      !p.sticky &&
      top + bubble.offsetHeight > bands[bandIdx].bottom &&
      top > bands[bandIdx].top
    ) {
      bubble.remove()
      continue
    }
    prevBottom = top + bubble.offsetHeight
    const path = document.createElementNS(SVG_NS, 'path')
    path.setAttribute(
      'd',
      `M ${p.x} ${p.y} L ${paperRight + BUBBLE_ENTRY_X / 2} ${p.y} L ${bubbleLeft} ${top + 9}`,
    )
    svg.appendChild(path)
  }
}
