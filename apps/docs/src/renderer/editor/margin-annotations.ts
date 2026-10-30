/**
 * Word-style margin annotations, drawn as an absolute overlay on the page wrap
 * (same pattern as .page-cut-overlays: zero layout height, rebuilt after each
 * pagination remeasure, cleared wholesale).
 *
 * - Comment bubbles: one bubble per open thread in a markup column right of the
 *   paper, top-aligned to its anchor line, stacked downward on overlap, with a
 *   dashed leader from the anchor. Hidden while the Comments panel is open
 *   (root `.comments-pane-open`): the pane replaces the margin cards, as in Word. The wrap gains `has-markup-area` +
 *   `--markup-w` so the column takes real width (screenshots/regression capture
 *   the wrap box); percentage-centered wrap overlays are re-centered in CSS.
 * - Revision bubbles: "Deleted:"/"Formatted:" balloons per Show Markup →
 *   Balloons mode (root data-balloons): revisions = both, deletions leave the
 *   text flow; comments (Word's default) = format changes only; inline = none
 *   and no markup column. Bubbles stack per page and overflow past a page
 *   bottom is dropped, like Word's overflow pane.
 * - Change bars: vertical segments in the left page margin covering every line
 *   that carries a tracked revision (Word's changed-line marks). In Simple
 *   Markup they are the only trace of a change: deletions are hidden, so their
 *   anchor line gets a bar too, and clicking a bar hands the position back
 *   (App switches to All Markup and scrolls there, like Word).
 * - Comment balloons (Simple Markup): one small icon per open thread inside
 *   the right page margin instead of the bubble column; click focuses the
 *   thread in the Comments panel.
 */
import type { CommentInfo } from '@genoffice/docx-engine'
import type { EditorView } from '@tiptap/pm/view'
import type { Mark as PmMark, Node as PmNode } from '@tiptap/pm/model'
import { getLang, t } from '../i18n/locale'
import { collectRevisions } from './revisions'
import { absoluteDate, authorInitials, authorInk } from './comment-card'
import { balloonRevKinds, currentBalloonMode } from './revision-view'
import { trackChangesOptionsState } from './track-changes-options'

export const MARKUP_AREA_W = 200
/** Word's comment column: a separate strip right of the markup area (or of the page) */
export const COMMENT_AREA_W = 260
const COMMENT_CARD_PAD = 28
const COMMENT_ENTRY_X = 14
/** column padding around a bubble: markup width − bubble width */
const BUBBLE_PAD = 32
const BUBBLE_ENTRY_X = 12
const BUBBLE_STACK_GAP = 6
const CHANGE_BAR_X = 24
const REV_TEXT_MAX = 220
const BALLOON_ICON_W = 18
const BALLOON_ICON_X = 48
/** Word draws the markup-area divider a little inside the page edge */
const STRIP_INSET = 20
const BALLOON_STACK_GAP = 2

const REV_SELECTOR =
  '.doc-ins, .doc-del, .has-move-from, .has-move-to, .has-rpr-change, .has-ppr-change, ' +
  'tr.row-rev-ins, tr.row-rev-del, td.cell-rev-ins, th.cell-rev-ins, td.cell-rev-del, th.cell-rev-del'

const SVG_NS = 'http://www.w3.org/2000/svg'

export type ChangeLine = { top: number; bottom: number; pos?: number }

/** touching / overlapping changed lines collapse into one bar that keeps its first document position */
export function mergeChangeLines(segs: ChangeLine[]): ChangeLine[] {
  segs.sort((a, b) => a.top - b.top)
  const out: ChangeLine[] = []
  for (const s of segs) {
    const last = out[out.length - 1]
    if (last && s.top <= last.bottom + 3) {
      last.bottom = Math.max(last.bottom, s.bottom)
      if (s.pos != null && (last.pos == null || s.pos < last.pos)) last.pos = s.pos
    } else out.push({ ...s })
  }
  return out
}

export interface MarginHandlers {
  onChangeBar?: (pos: number) => void
  /** anchor icon clicked */
  onCommentBalloon?: (id: string) => void
  /** margin card clicked: select that thread */
  onCommentSelect?: (id: string) => void
  /** Reply typed into a margin bubble's inline box */
  onCommentReply?: (id: string, text: string) => void
  onCommentResolve?: (id: string, done: boolean) => void
  onCommentEdit?: (id: string, text: string) => void
  onCommentDelete?: (id: string) => void
}

type BubbleHandlers = Pick<
  MarginHandlers,
  'onCommentReply' | 'onCommentResolve' | 'onCommentEdit' | 'onCommentDelete'
>

function makeBalloonIcon(c: CommentInfo, on: boolean, onJump: () => void): HTMLElement {
  const btn = document.createElement('button')
  btn.type = 'button'
  btn.className = `comment-balloon${on ? ' on' : ''}`
  btn.setAttribute('aria-pressed', on ? 'true' : 'false')
  btn.title = `${c.author}\n${c.text.split('\n')[0] ?? ''}`
  btn.dataset.commentId = c.id
  btn.innerHTML =
    '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M 2.99 3.91 h 10.01 v 6.83 h -5.46 L 4.81 13.46 v -2.73 h -1.82 z"/><path d="M 5.27 6.18 h 5.46 M 5.27 8.46 h 3.64"/></svg>'
  btn.addEventListener('click', onJump)
  return btn
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

/** avatar + author + relative time, the same markup the Comments panel renders (shared classes in styles.css) */
function makeCommentHead(c: CommentInfo): HTMLElement {
  const head = document.createElement('div')
  head.className = 'comment-card-head'
  const name = c.author || t('appUnknownAuthor')
  const avatar = document.createElement('span')
  avatar.className = 'comment-avatar'
  avatar.style.background = authorInk(name)
  avatar.textContent = authorInitials(c)
  const meta = document.createElement('span')
  meta.className = 'comment-meta'
  const author = document.createElement('span')
  author.className = 'comment-author'
  author.textContent = name
  meta.appendChild(author)
  if (c.date) {
    const when = document.createElement('span')
    when.className = 'comment-when'
    when.title = absoluteDate(c.date)
    when.textContent = revWhen(c.date)
    meta.appendChild(when)
  }
  head.append(avatar, meta)
  return head
}

const LIKE_SVG =
  '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round" aria-hidden="true"><path d="M5 7.5v6H2.5v-6zM5 7.5l2.6-5.2c1 0 1.7.8 1.5 1.8L8.7 6.5h3.6c.9 0 1.5.8 1.3 1.7l-1 4.3c-.2.6-.7 1-1.3 1H5"/></svg>'
/** 👍 reactions live in the session only: the docx comment model has no slot for them */
const likedThreads = new Set<string>()

const CHECK_SVG =
  '<svg viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.8 8.6 6.2 12l7-7.5"/></svg>'

function iconButton(
  cls: string,
  label: string,
  html: string,
  onClick: () => void,
): HTMLButtonElement {
  const btn = document.createElement('button')
  btn.type = 'button'
  btn.className = cls
  btn.setAttribute('aria-label', label)
  btn.dataset.tip = label
  btn.innerHTML = html
  btn.addEventListener('click', (e) => {
    e.stopPropagation()
    onClick()
  })
  return btn
}

function menuItem(cls: string, label: string, onClick: () => void): HTMLButtonElement {
  const btn = document.createElement('button')
  btn.type = 'button'
  btn.className = cls
  btn.setAttribute('role', 'menuitem')
  btn.textContent = label
  btn.addEventListener('click', (e) => {
    e.stopPropagation()
    onClick()
  })
  return btn
}

/** Word's card head: ✓ Resolve on the root, ⋯ → Edit / Delete on every entry */
function makeCardActions(
  c: CommentInfo,
  root: boolean,
  replyCount: number,
  h: BubbleHandlers,
  startEdit: () => void,
  flushEdits: () => void,
): HTMLElement | null {
  const actions = document.createElement('span')
  actions.className = 'comment-card-actions'
  actions.addEventListener('click', (e) => e.stopPropagation())
  if (root && h.onCommentResolve) {
    actions.appendChild(
      iconButton('comment-resolve', t('appResolve'), CHECK_SVG, () => {
        // resolving rebuilds the card away: commit an open edit first, as the panel does
        flushEdits()
        h.onCommentResolve!(c.id, true)
      }),
    )
  }
  if (h.onCommentEdit || h.onCommentDelete) {
    const more = iconButton('comment-more', t('appMoreActions'), '⋯', () => toggle())
    more.setAttribute('aria-haspopup', 'menu')
    const menu = document.createElement('div')
    menu.className = 'comment-menu'
    menu.setAttribute('role', 'menu')
    menu.hidden = true
    if (h.onCommentEdit)
      menu.appendChild(
        menuItem('comment-card-edit', t(root ? 'appEditComment' : 'appEditReply'), () => {
          close()
          startEdit()
        }),
      )
    if (h.onCommentDelete) {
      const label = root
        ? replyCount > 0
          ? t('appDeleteCommentWithReplies')
          : t('appDeleteComment')
        : t('appDeleteReply')
      menu.appendChild(menuItem('comment-card-del', label, () => h.onCommentDelete!(c.id)))
    }
    const onDocDown = (e: PointerEvent) => {
      if (!actions.contains(e.target as Node)) close()
    }
    const close = () => {
      menu.hidden = true
      more.setAttribute('aria-expanded', 'false')
      document.removeEventListener('pointerdown', onDocDown, true)
    }
    const toggle = () => {
      if (!menu.hidden) return close()
      menu.hidden = false
      more.setAttribute('aria-expanded', 'true')
      document.addEventListener('pointerdown', onDocDown, true)
    }
    actions.append(more, menu)
  }
  if (root) {
    const like = iconButton('comment-like', t('appLike'), LIKE_SVG, () => {
      if (likedThreads.has(c.id)) likedThreads.delete(c.id)
      else likedThreads.add(c.id)
      like.classList.toggle('on', likedThreads.has(c.id))
      like.setAttribute('aria-pressed', String(likedThreads.has(c.id)))
    })
    like.classList.toggle('on', likedThreads.has(c.id))
    like.setAttribute('aria-pressed', String(likedThreads.has(c.id)))
    actions.appendChild(like)
  }
  return actions.childElementCount ? actions : null
}

function makeBubble(
  root: CommentInfo,
  replies: CommentInfo[],
  onJump: () => void,
  h: BubbleHandlers,
): HTMLElement {
  const bubble = document.createElement('div')
  bubble.className = 'comment-bubble'
  bubble.dataset.commentId = root.id
  const openEdits = new Set<() => void>()
  const flushEdits = () => {
    for (const save of [...openEdits]) save()
  }
  const addEntry = (c: CommentInfo, isRoot: boolean) => {
    const entry = document.createElement('div')
    entry.className = isRoot ? 'comment-entry' : 'comment-reply'
    entry.dataset.entryId = c.id
    const text = document.createElement('div')
    text.className = 'comment-text'
    text.textContent = c.text
    const head = makeCommentHead(c)
    const startEdit = () => {
      if (!h.onCommentEdit || entry.querySelector('textarea')) return
      const box = document.createElement('div')
      box.className = 'comment-reply-compose'
      const ta = document.createElement('textarea')
      ta.className = 'comment-edit-input'
      ta.value = c.text
      const row = document.createElement('div')
      row.className = 'comment-compose-actions'
      const cancel = () => {
        openEdits.delete(save)
        box.replaceWith(text)
      }
      const save = () => {
        const v = ta.value.trim()
        if (!v) return
        if (v !== c.text) h.onCommentEdit!(c.id, v)
        cancel()
      }
      openEdits.add(save)
      const cancelBtn = document.createElement('button')
      cancelBtn.type = 'button'
      cancelBtn.textContent = t('appCancel')
      cancelBtn.addEventListener('click', cancel)
      const saveBtn = document.createElement('button')
      saveBtn.type = 'button'
      saveBtn.className = 'primary'
      saveBtn.textContent = t('appSave')
      saveBtn.addEventListener('click', save)
      ta.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) save()
        if (e.key === 'Escape') cancel()
        e.stopPropagation()
      })
      box.addEventListener('click', (e) => e.stopPropagation())
      row.append(cancelBtn, saveBtn)
      box.append(ta, row)
      text.replaceWith(box)
      ta.focus()
      ta.setSelectionRange(ta.value.length, ta.value.length)
    }
    const actions = makeCardActions(c, isRoot, replies.length, h, startEdit, flushEdits)
    if (actions) head.appendChild(actions)
    entry.append(head, text)
    bubble.appendChild(entry)
  }
  addEntry(root, true)
  for (const r of replies) addEntry(r, false)
  const onReply = h.onCommentReply
  if (onReply) {
    // Word's margin card: the Reply box is part of the selected card (.active, CSS), posts on Enter
    const input = document.createElement('input')
    input.type = 'text'
    input.className = 'comment-reply-input'
    input.placeholder = t('appReplyPlaceholder')
    input.setAttribute('aria-label', t('appReply'))
    input.addEventListener('click', (e) => e.stopPropagation())
    input.addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
        e.preventDefault()
        const text = input.value.trim()
        if (!text) return
        input.value = ''
        onReply(root.id, text)
      }
      if (e.key === 'Escape') input.blur()
    })
    bubble.appendChild(input)
  }
  bubble.addEventListener('click', onJump)
  return bubble
}

/** per-thread display flags the app root carries for the margin layer */
function commentFlags(wrap: HTMLElement) {
  const root = wrap.closest('.app')
  return {
    cards: !!root?.hasAttribute('data-comment-cards'),
    focus: root?.getAttribute('data-comment-focus') ?? null,
    popup: root?.getAttribute('data-comment-popup') ?? null,
  }
}

/** Word's balloon clock: relative within a week, else the locale date */
export function revWhen(date: string | null | undefined): string {
  if (!date) return ''
  const ms = Date.parse(date)
  if (Number.isNaN(ms)) return ''
  const lang = getLang()
  const diffMin = Math.round((ms - Date.now()) / 60000)
  const abs = Math.abs(diffMin)
  const rtf = new Intl.RelativeTimeFormat(lang, { numeric: 'auto' })
  if (abs < 1) return rtf.format(0, 'minute')
  if (abs < 60) return rtf.format(diffMin, 'minute')
  if (abs < 60 * 24) return rtf.format(Math.round(diffMin / 60), 'hour')
  if (abs < 60 * 24 * 7) return rtf.format(Math.round(diffMin / (60 * 24)), 'day')
  return new Date(ms).toLocaleDateString(lang)
}

function makeRevBubble(item: RevItem): HTMLElement {
  const { kind, text } = item
  const bubble = document.createElement('div')
  bubble.className = `comment-bubble rev-bubble rev-bubble-${kind}`
  bubble.dataset.pos = String(item.from)
  bubble.dataset.end = String(item.to)
  const head = document.createElement('div')
  head.className = 'rev-bubble-head'
  const author = document.createElement('span')
  author.className = 'rev-bubble-author'
  author.textContent = item.author || t('editorUnknownAuthor')
  const when = document.createElement('span')
  when.className = 'rev-bubble-when'
  when.textContent = revWhen(item.date)
  head.append(author, when)
  const line = document.createElement('div')
  const label = document.createElement('span')
  label.className = 'rev-bubble-label'
  label.textContent = `${t(item.moved ? 'editorRevMoved' : kind === 'del' ? 'editorRevDeleted' : 'editorRevFormatted')}: `
  const body = document.createElement('span')
  body.className = 'rev-bubble-text'
  body.textContent = text.length > REV_TEXT_MAX ? `${text.slice(0, REV_TEXT_MAX)}…` : text
  line.append(label, body)
  bubble.append(head, line)
  return bubble
}

/** Word names its standard palette in balloons ("Font color: Red"); anything else stays hex */
const COLOR_NAME_KEYS: Record<string, Parameters<typeof t>[0]> = {
  '000000': 'editorColorBlack',
  FFFFFF: 'editorColorWhite',
  C00000: 'editorColorDarkRed',
  FF0000: 'editorColorRed',
  FFC000: 'editorColorOrange',
  FFFF00: 'editorColorYellow',
  '92D050': 'editorColorLightGreen',
  '00B050': 'editorColorGreen',
  '00B0F0': 'editorColorLightBlue',
  '0070C0': 'editorColorBlue',
  '002060': 'editorColorDarkBlue',
  '7030A0': 'editorColorPurple',
  '0000FF': 'editorColorBlue',
  '008000': 'editorColorGreen',
  '808080': 'editorColorGray',
}

function colorName(hex: string): string {
  const key = COLOR_NAME_KEYS[hex]
  return key ? t(key) : `#${hex}`
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
  const underline = marks.some((m) => m.type.name === 'underline')
  if (underline !== !!old.underline)
    parts.push(t(underline ? 'editorRevUnderline' : 'editorRevNotUnderline'))
  const strike = marks.some((m) => m.type.name === 'strike')
  if (strike !== !!old.strike) parts.push(t(strike ? 'editorRevStrike' : 'editorRevNotStrike'))
  const colorKey = (c: unknown) => (typeof c === 'string' && c !== 'auto' ? c.toUpperCase() : null)
  const newColor = colorKey(ts?.attrs.color)
  if (newColor !== colorKey(old.color)) {
    parts.push(
      `${t('editorRevFontColor')}: ${newColor ? colorName(newColor) : t('editorRevColorAuto')}`,
    )
  }
  const newHl = (ts?.attrs.highlight as string | null | undefined) ?? null
  if (newHl !== (((old.highlight as string | undefined) ?? null) as string | null)) {
    parts.push(t('editorRevHighlight'))
  }
  return parts.join(', ') || t('editorFormatRevision')
}

export type RevItem = {
  kind: 'del' | 'fmt'
  from: number
  to: number
  text: string
  author: string
  date: string | null
  /** a deletion that is the source of a move (balloon says Moved) */
  moved?: boolean
}

const revMeta = (a: { author?: unknown; date?: unknown } | null | undefined) => ({
  author: typeof a?.author === 'string' ? a.author : '',
  date: typeof a?.date === 'string' ? a.date : null,
})

const parseMeta = (json: unknown) => {
  if (typeof json !== 'string') return revMeta(null)
  try {
    return revMeta(JSON.parse(json))
  } catch {
    return revMeta(null)
  }
}

/**
 * Tracked revisions that live in balloons, in document order, merged over
 * contiguous positions. A deleted paragraph mark bridges [end-1, end+1], so a
 * fully deleted paragraph chains with the next paragraph's leading deletion
 * into one balloon, matching Word's "XXX¶XXX¶" grouping.
 */
export function revGroupsOf(doc: PmNode): RevItem[] {
  const items: RevItem[] = []
  const propItem = (record: unknown, pos: number, text: string) => {
    const r = record as { author?: string; date?: string; resolved?: string } | null
    if (!r || typeof r !== 'object' || r.resolved) return
    items.push({ kind: 'fmt', from: pos, to: pos, text, ...revMeta(r) })
  }
  const pendingJson = (raw: unknown) => typeof raw === 'string' && !raw.includes('"resolved"')
  doc.descendants((node, pos) => {
    const name = node.type.name
    if (name === 'docTable') {
      // the balloon anchors at the first text position inside the table
      const at = pos + 4
      propItem(node.attrs?.tblPrChange, at, t('editorRevFormattedTable'))
      propItem(node.attrs?.tblGridChange, at, t('editorRevFormattedTable'))
    } else if (name === 'docTableRow') {
      propItem(node.attrs?.trPrChange, pos + 3, t('editorRevFormattedTable'))
      propItem(node.attrs?.tblPrExChange, pos + 3, t('editorRevFormattedTable'))
    } else if (name === 'docTableCell' || name === 'docTableHeader') {
      propItem(node.attrs?.tcPrChange, pos + 2, t('editorRevFormattedTable'))
      propItem(node.attrs?.cellMerge, pos + 2, t('editorRevMergedCells'))
    }
    if (pendingJson(node.attrs?.sectPrChange)) {
      items.push({
        kind: 'fmt',
        from: pos + 1,
        to: pos + 1,
        text: t('editorRevFormattedSection'),
        ...parseMeta(node.attrs!.sectPrChange),
      })
    }
    if (node.isTextblock) {
      if (node.attrs?.numberingChange) {
        items.push({
          kind: 'fmt',
          from: pos + 1,
          to: pos + 1,
          text: t('editorRevNumberingChange'),
          ...parseMeta(node.attrs.numberingChange),
        })
      }
      if (node.attrs?.pPrChange) {
        items.push({
          kind: 'fmt',
          from: pos + 1,
          to: pos + 1,
          text: t('editorFormatRevision'),
          ...parseMeta(node.attrs.pPrChange),
        })
      }
      if (node.attrs?.paraMarkDel) {
        const end = pos + node.nodeSize
        items.push({
          kind: 'del',
          from: end - 1,
          to: end + 1,
          text: '¶',
          ...parseMeta(node.attrs.paraMarkDel),
        })
      }
      return true
    }
    if (!node.isText) return true
    const del = node.marks.find((m) => m.type.name === 'del')
    if (del) {
      items.push({
        kind: 'del',
        from: pos,
        to: pos + node.nodeSize,
        text: node.text ?? '',
        ...revMeta(del.attrs),
        ...(del.attrs.move ? { moved: true } : {}),
      })
      return true
    }
    const rpr = node.marks.find((m) => m.type.name === 'rprChange')
    if (rpr) {
      items.push({
        kind: 'fmt',
        from: pos,
        to: pos + node.nodeSize,
        text: fmtDesc((rpr.attrs.old ?? {}) as Record<string, unknown>, node.marks),
        ...revMeta(rpr.attrs),
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
      !!last.moved === !!it.moved &&
      last.author === it.author &&
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

export type AnchorPoint = { top: number; bottom: number; left: number }

/**
 * On-screen point for a hidden element: the end of the previous visible
 * sibling, else the first line of the nearest visible ancestor (a deletion
 * opening a paragraph anchors at that paragraph's start), walking up.
 */
export function visiblePointNear(el: HTMLElement | null, pm: HTMLElement): AnchorPoint | null {
  let cur: HTMLElement | null = el
  while (cur && cur !== pm) {
    for (
      let sib = cur.previousElementSibling as HTMLElement | null;
      sib;
      sib = sib.previousElementSibling as HTMLElement | null
    ) {
      const rects = sib.getClientRects()
      if (rects.length > 0) {
        const r = rects[rects.length - 1]
        return { top: r.top, bottom: r.bottom, left: r.right }
      }
    }
    const parent = cur.parentElement
    if (!parent || parent === pm) break
    const own = parent.getClientRects()
    if (own.length > 0) {
      const r = own[0]
      return { top: r.top, bottom: r.top + Math.min(r.height, 16), left: r.left }
    }
    cur = parent
  }
  return null
}

export function anchorPointFor(view: EditorView, pm: HTMLElement, pos: number): AnchorPoint | null {
  const clamped = Math.max(0, Math.min(pos, view.state.doc.content.size))
  try {
    const c = view.coordsAtPos(clamped, -1)
    if (c && (c.top !== 0 || c.left !== 0 || c.bottom !== 0)) {
      return { top: c.top, bottom: c.bottom, left: c.left }
    }
  } catch {
    /* position renders inside hidden (balloon-collapsed) content */
  }
  try {
    const { node, offset } = view.domAtPos(clamped)
    const child = node.childNodes[offset] ?? node.childNodes[offset - 1] ?? node
    const el = (child instanceof HTMLElement ? child : child.parentElement) as HTMLElement | null
    return visiblePointNear(el, pm)
  } catch {
    return null
  }
}

/**
 * Bar segment for hidden (deleted) content anchored at a visible point: a line
 * anchor gets one line of bar; a whole-row fallback rect would bar the untouched
 * neighbor row, so it marks the seam where the deleted content was instead.
 */
function hiddenLineSeg(p: AnchorPoint, wrapRect: DOMRect, f: number, pos: number): ChangeLine {
  const top = (p.top - wrapRect.top) / f
  const bottom = (p.bottom - wrapRect.top) / f
  if (bottom - top > 24) return { top: bottom - 8, bottom: bottom + 4, pos }
  return { top, bottom: Math.max(bottom, top + 12), pos }
}

/** Page bands (wrap coords) from the top-level page-gap widgets; one whole-doc band when absent. */
function pageBandsOf(pm: HTMLElement, wrapRect: DOMRect, wrapH: number, f: number): ChangeLine[] {
  const gaps = [...pm.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-inline)')]
    .map((g) => g.getBoundingClientRect())
    .filter((r) => r.height > 0)
    .sort((a, b) => a.top - b.top)
  if (gaps.length === 0) return [{ top: 0, bottom: wrapH }]
  const bands: ChangeLine[] = []
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
  type?: string
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
  handlers?: MarginHandlers,
): void {
  const f = zoomFactor
  const wrapRect = wrap.getBoundingClientRect()
  const pmRect = pm.getBoundingClientRect()
  const simple = !!wrap.closest('.rev-display-simple')
  const opts = trackChangesOptionsState.current
  const balloons = currentBalloonMode(wrap)
  // Show All Revisions Inline: no margin column at all (comments carry inline tags).
  // Anchor icons always sit in the page margin; cards need Show Comments → Contextual
  // (data-comment-cards), else only the popped-up thread (data-comment-popup) gets one
  const inline = balloons === 'inline'
  const flags = commentFlags(wrap)
  const cardsOn = flags.cards && !simple
  const revKinds = view && balloons ? balloonRevKinds(balloons) : new Set<string>()
  const delInBalloons = revKinds.has('del')

  // change bars in All / Simple Markup (Word hides them in No Markup / Original)
  const segs: ChangeLine[] = []
  if (!wrap.closest('.rev-display-none, .rev-display-original')) {
    for (const el of pm.querySelectorAll<HTMLElement>(REV_SELECTOR)) {
      let pos: number | undefined
      if (simple && view) {
        try {
          pos = view.posAtDOM(el, 0)
        } catch {
          /* detached decoration node */
        }
      }
      for (const r of el.getClientRects()) {
        if (r.height > 0)
          segs.push({ top: (r.top - wrapRect.top) / f, bottom: (r.bottom - wrapRect.top) / f, pos })
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
  const icons: Placed[] = []
  const bubbleHandlers: BubbleHandlers = {
    onCommentReply: handlers?.onCommentReply,
    onCommentResolve: handlers?.onCommentResolve,
    onCommentEdit: handlers?.onCommentEdit,
    onCommentDelete: handlers?.onCommentDelete,
  }
  const localThread = (c: CommentInfo, rect: DOMRect, blockAnchor: HTMLElement | null) => {
    // block-anchored threads have no .doc-comment span: flash the block itself
    const flash = blockAnchor ? () => flashEls([blockAnchor]) : () => flashAnchors(pm, c.id)
    const top = (rect.top - wrapRect.top) / f
    const selected = flags.focus === c.id || flags.popup === c.id
    icons.push({
      sticky: true,
      top,
      y: 0,
      x: 0,
      make: () => makeBalloonIcon(c, selected, () => handlers?.onCommentBalloon?.(c.id)),
    })
    const popup = !cardsOn && flags.popup === c.id
    if (!cardsOn && !popup) return
    placed.push({
      sticky: true,
      top,
      // leader start: end of the marked range, or start of the anchor paragraph's
      // first line for range-less comments (bare w:commentReference on an empty run)
      y: blockAnchor ? top + Math.min(rect.height / f, 16) : (rect.bottom - wrapRect.top) / f - 1,
      x: ((blockAnchor ? rect.left : rect.right) - wrapRect.left) / f,
      make: () => {
        const bubble = makeBubble(
          c,
          comments.filter((r) => r.parentId === c.id),
          () => {
            flash()
            handlers?.onCommentSelect?.(c.id)
          },
          bubbleHandlers,
        )
        if (flags.focus === c.id) bubble.classList.add('active')
        if (popup) bubble.classList.add('comment-popup')
        return bubble
      },
    })
  }
  for (const c of comments) {
    if (inline || c.parentId || c.done) continue
    const el = anchorOf.get(c.id)
    const rect = el ? [...el.getClientRects()].find((r) => r.height > 0) : undefined
    if (rect) {
      localThread(c, rect, null)
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
    if (blockEl && blockRect) localThread(c, blockRect, blockEl)
  }

  // hidden deletions (Simple Markup, revisions-in-balloons) have no rects for
  // the bar pass above: mark their anchor line (the sibling-fallback anchor is
  // a point — give it one line of bar); balloon modes add their bubbles
  if (view && simple) {
    // deleted rows / cells leave the flow too: bar the line where they were
    for (const r of collectRevisions(view.state.doc)) {
      if (r.kind !== 'rowDel' && r.kind !== 'cellDel') continue
      const p = anchorPointFor(view, pm, r.from)
      if (p) segs.push(hiddenLineSeg(p, wrapRect, f, r.from))
    }
  }
  if (view && (simple || revKinds.size > 0)) {
    for (const g of revGroupsOf(view.state.doc)) {
      const hiddenDel = g.kind === 'del' && (simple || delInBalloons)
      if (!hiddenDel && !revKinds.has(g.kind)) continue
      const p = anchorPointFor(view, pm, g.from)
      if (!p) continue
      if (hiddenDel) segs.push(hiddenLineSeg(p, wrapRect, f, g.from))
      if (simple || !revKinds.has(g.kind)) continue
      placed.push({
        sticky: false,
        top: (p.top - wrapRect.top) / f,
        y: (p.bottom - wrapRect.top) / f - 1,
        x: (p.left - wrapRect.left) / f,
        make: () => makeRevBubble(g),
      })
    }
  }
  placed.sort((a, b) => a.top - b.top)
  icons.sort((a, b) => a.top - b.top)
  const bars = opts.changedLines === 'none' ? [] : mergeChangeLines(segs)

  if (bars.length === 0 && placed.length === 0 && icons.length === 0) {
    clearMarginAnnotations(wrap)
    return
  }

  let layer = wrap.querySelector(':scope > .page-margin-annotations') as HTMLElement | null
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-margin-annotations'
    wrap.appendChild(layer)
  }
  // selecting a card rebuilds this layer: carry a half-typed reply over to the new card
  const typing = layer.querySelector<HTMLInputElement>('.comment-reply-input:focus')
  const draft = typing && {
    id: typing.closest<HTMLElement>('.comment-bubble')?.dataset.commentId,
    value: typing.value,
    caret: typing.selectionStart ?? typing.value.length,
  }
  layer.textContent = ''

  // Word: revision balloons live in a markup strip glued to the page; comment cards
  // sit in their own column outside it (a popup floats there without reserving width)
  const revColumn = !simple && placed.some((p) => !p.sticky)
  const commentColumn = cardsOn && placed.some((p) => p.sticky)
  const revW = revColumn ? opts.balloonWidth : 0
  const commentW = commentColumn ? COMMENT_AREA_W : 0
  const column = revColumn || commentColumn
  const leftSide = opts.balloonMargin === 'left'
  wrap.classList.toggle('has-markup-area', column)
  wrap.classList.toggle('markup-left', column && leftSide)
  wrap.classList.toggle('markup-no-lines', !opts.balloonLines)
  if (column) wrap.style.setProperty('--markup-w', `${revW + commentW}px`)

  // a left column pads the wrap on the paper's side: the paper just moved, so
  // re-measure it and carry the anchors (measured above) along
  const paperNow = pm.getBoundingClientRect()
  const shiftX = (paperNow.left - pmRect.left) / f
  const paperLeft = (paperNow.left - wrapRect.left) / f
  const paperRight = (paperNow.right - wrapRect.left) / f
  const bands = pageBandsOf(pm, wrapRect, wrapRect.height / f, f)
  if (revColumn) {
    for (const b of bands) {
      const strip = document.createElement('div')
      strip.className = 'markup-strip'
      strip.style.left = `${leftSide ? paperLeft - revW : paperRight - STRIP_INSET}px`
      strip.style.width = `${revW + STRIP_INSET}px`
      strip.style.top = `${b.top}px`
      strip.style.height = `${b.bottom - b.top}px`
      layer.appendChild(strip)
    }
  }
  const barX =
    opts.changedLines === 'right' ? paperRight - CHANGE_BAR_X - 2 : paperLeft + CHANGE_BAR_X
  for (const b of bars) {
    const el = document.createElement('div')
    el.className = 'change-bar'
    el.style.left = `${barX}px`
    el.style.top = `${b.top}px`
    el.style.height = `${b.bottom - b.top}px`
    if (simple && b.pos != null && handlers?.onChangeBar) {
      const pos = b.pos
      el.addEventListener('click', () => handlers.onChangeBar?.(pos))
    }
    layer.appendChild(el)
  }

  // Word's anchor glyphs: one per commented line, stacked inside the right page margin
  let prevIconBottom = -Infinity
  for (const p of icons) {
    const icon = p.make()
    const top = Math.max(p.top, prevIconBottom + BALLOON_STACK_GAP)
    icon.style.left = `${paperRight - BALLOON_ICON_X}px`
    icon.style.top = `${top}px`
    layer.appendChild(icon)
    prevIconBottom = top + BALLOON_ICON_W
  }
  if (placed.length === 0) return
  const svg = document.createElementNS(SVG_NS, 'svg')
  svg.setAttribute('class', 'comment-leaders')
  layer.appendChild(svg)
  // two independent stacks per page band; a revision balloon that cannot start within
  // its band is dropped (Word routes those to the overflow pane), comments always render
  const revBubbleW = opts.balloonWidth - BUBBLE_PAD
  const cardW = COMMENT_AREA_W - COMMENT_CARD_PAD
  const revLeft = leftSide ? paperLeft - BUBBLE_ENTRY_X - revBubbleW : paperRight + BUBBLE_ENTRY_X
  const cardLeft = leftSide
    ? paperLeft - revW - COMMENT_ENTRY_X - cardW
    : paperRight + revW + COMMENT_ENTRY_X
  const entryX = leftSide ? paperLeft - BUBBLE_ENTRY_X / 2 : paperRight + BUBBLE_ENTRY_X / 2
  const stacks = {
    rev: { bandIdx: 0, prevBottom: -Infinity },
    card: { bandIdx: 0, prevBottom: -Infinity },
  }
  for (const p of placed) {
    const isCard = p.sticky
    const st = isCard ? stacks.card : stacks.rev
    while (st.bandIdx < bands.length - 1 && p.top >= bands[st.bandIdx + 1].top) {
      st.bandIdx++
      st.prevBottom = -Infinity
    }
    const bubble = p.make()
    const left = isCard ? cardLeft : revLeft
    const width = isCard ? cardW : revBubbleW
    bubble.style.left = `${left}px`
    bubble.style.width = `${width}px`
    const top = Math.max(p.top, st.prevBottom + BUBBLE_STACK_GAP, bands[st.bandIdx].top)
    bubble.style.top = `${top}px`
    layer.appendChild(bubble)
    if (
      !isCard &&
      top + bubble.offsetHeight > bands[st.bandIdx].bottom &&
      top > bands[st.bandIdx].top
    ) {
      bubble.remove()
      continue
    }
    st.prevBottom = top + bubble.offsetHeight
    const edge = leftSide ? left + width : left
    const path = document.createElementNS(SVG_NS, 'path')
    // a card's leader runs from its anchor glyph, not across the markup strip
    const d = isCard
      ? `M ${leftSide ? paperLeft + BALLOON_ICON_X : paperRight - BALLOON_ICON_X + BALLOON_ICON_W} ${p.top + BALLOON_ICON_W / 2} L ${edge} ${top + 9}`
      : `M ${p.x + shiftX} ${p.y} L ${entryX} ${p.y} L ${edge} ${top + 9}`
    path.setAttribute('d', d)
    if (isCard) path.setAttribute('class', 'comment-leader')
    svg.appendChild(path)
    if (draft && draft.id === bubble.dataset.commentId) {
      const input = bubble.querySelector<HTMLInputElement>('.comment-reply-input')
      if (input) {
        input.value = draft.value
        input.focus()
        input.setSelectionRange(draft.caret, draft.caret)
      }
    }
  }
}
