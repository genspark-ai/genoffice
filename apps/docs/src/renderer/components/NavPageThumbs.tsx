import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { useI18n } from '../i18n/locale'

/** page paper edges in unzoomed canvas px, measured from the live page gaps */
export interface ThumbPage {
  top: number
  height: number
}

interface ThumbBlock {
  el: HTMLElement
  top: number
  left: number
  width: number
  height: number
}

export interface ThumbSnapshot {
  pageW: number
  pages: ThumbPage[]
  blocks: ThumbBlock[]
  /** the canvas' page-geometry custom properties, so clones lay out like the page */
  vars: string
}

const FALLBACK_PAGE_H = 1056
const THUMB_GUTTER = 12
const THUMB_MAX_W = 200

/**
 * Word's Pages tab reads the printed pages; ours reads the canvas: every non-carry
 * page gap ends one paper and starts the next, so the thumbnails show exactly what
 * the canvas shows (in-table cuts and mid-paragraph breaks included).
 */
export function snapshotPages(pm: HTMLElement, zoom: number): ThumbSnapshot | null {
  const wrap = pm.closest<HTMLElement>('.page-wrap')
  const docZoom = wrap?.parentElement
  if (!wrap || !docZoom || zoom <= 0) return null
  const wrapRect = wrap.getBoundingClientRect()
  if (wrapRect.width === 0) return null
  const pageW = wrapRect.width / zoom
  const pageH =
    parseFloat(getComputedStyle(docZoom).getPropertyValue('--page-h')) || FALLBACK_PAGE_H
  const starts = [0]
  const ends: number[] = []
  for (const gap of pm.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-carry)')) {
    const r = gap.getBoundingClientRect()
    if (r.height === 0) continue
    const cs = getComputedStyle(gap)
    const mb = parseFloat(cs.getPropertyValue('--gap-mb')) || 0
    const mt = parseFloat(cs.getPropertyValue('--gap-mt')) || 0
    ends.push((r.top - wrapRect.top) / zoom + mb)
    starts.push((r.bottom - wrapRect.top) / zoom - mt)
  }
  ends.push(starts[starts.length - 1] + pageH)
  const pages = starts.map((top, i) => ({ top, height: Math.max(1, ends[i] - top) }))
  const blocks: ThumbBlock[] = []
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    const r = el.getBoundingClientRect()
    if (r.height === 0 || r.width === 0) continue
    blocks.push({
      el,
      top: (r.top - wrapRect.top) / zoom,
      left: (r.left - wrapRect.left) / zoom,
      width: r.width / zoom,
      height: r.height / zoom,
    })
  }
  const vars = (docZoom.getAttribute('style') ?? '')
    .split(';')
    .filter((d) => d.trim().startsWith('--'))
    .join(';')
  return { pageW, pages, blocks, vars }
}

/** index of the page whose paper contains canvas y */
export function pageIndexAtY(pages: ThumbPage[], y: number): number {
  let idx = 0
  for (let i = 1; i < pages.length; i++) if (pages[i].top <= y) idx = i
  return idx
}

function Thumb({
  snap,
  index,
  scale,
  active,
  hit,
  onPick,
}: {
  snap: ThumbSnapshot
  index: number
  scale: number
  active: boolean
  hit: boolean
  onPick: (index: number) => void
}) {
  const { t } = useI18n()
  const hostRef = useRef<HTMLButtonElement>(null)
  const pageRef = useRef<HTMLDivElement>(null)
  const [visible, setVisible] = useState(false)
  const page = snap.pages[index]

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const io = new IntersectionObserver(
      (entries) => setVisible(entries.some((e) => e.isIntersecting)),
      { root: host.closest('.nav-body'), rootMargin: '200px 0px' },
    )
    io.observe(host)
    return () => io.disconnect()
  }, [])

  useLayoutEffect(() => {
    const root = pageRef.current
    if (!root || !visible) return
    root.replaceChildren()
    const bottom = page.top + page.height
    for (const b of snap.blocks) {
      if (b.top >= bottom || b.top + b.height <= page.top) continue
      const slot = document.createElement('div')
      slot.className = 'nav-thumb-slot'
      slot.style.top = `${b.top - page.top}px`
      slot.style.left = `${b.left}px`
      slot.style.width = `${b.width}px`
      const clone = b.el.cloneNode(true) as HTMLElement
      // the slot already sits where the block's margins and transforms put it
      clone.style.margin = '0'
      clone.style.transform = 'none'
      clone.removeAttribute('contenteditable')
      slot.appendChild(clone)
      root.appendChild(slot)
    }
    return () => root.replaceChildren()
  }, [snap, page, visible])

  useEffect(() => {
    if (active) hostRef.current?.scrollIntoView({ block: 'nearest' })
  }, [active])

  return (
    <button
      ref={hostRef}
      type="button"
      className={`nav-thumb${active ? ' on' : ''}${hit ? ' hit' : ''}`}
      style={{ width: snap.pageW * scale, height: page.height * scale }}
      aria-label={t('appNavPage', { n: index + 1 })}
      aria-current={active ? 'page' : undefined}
      onClick={() => onPick(index)}
    >
      {visible && (
        <div
          ref={pageRef}
          className="nav-thumb-page doc-page ProseMirror"
          style={{
            width: snap.pageW,
            height: page.height,
            transform: `scale(${scale})`,
            ...styleVars(snap.vars),
          }}
          aria-hidden
        />
      )}
      <span className="nav-thumb-num">{index + 1}</span>
    </button>
  )
}

function styleVars(vars: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const decl of vars.split(';')) {
    const i = decl.indexOf(':')
    if (i > 0) out[decl.slice(0, i).trim()] = decl.slice(i + 1).trim()
  }
  return out
}

export function NavPageThumbs({
  editor,
  zoom,
  width,
  layoutVersion,
  currentPage,
  hitPositions,
}: {
  editor: Editor
  zoom: number
  width: number
  /** bumps when the document or its pagination changed */
  layoutVersion: number
  /** 1-based visible page under the viewport middle (status bar page) */
  currentPage: number
  /** search hit positions: when non-null the tab lists only pages with a hit */
  hitPositions: number[] | null
}) {
  const { t } = useI18n()
  const [snap, setSnap] = useState<ThumbSnapshot | null>(null)
  // pagination lands after the document change: the page-wrap grows or shrinks with its gaps
  const [wrapHeight, setWrapHeight] = useState(0)
  useEffect(() => {
    const wrap = editor.view.dom.closest<HTMLElement>('.page-wrap')
    if (!wrap) return
    const ro = new ResizeObserver(() => setWrapHeight(wrap.offsetHeight))
    ro.observe(wrap)
    return () => ro.disconnect()
  }, [editor])

  useEffect(() => {
    let raf = 0
    const timer = window.setTimeout(() => {
      raf = requestAnimationFrame(() => setSnap(snapshotPages(editor.view.dom, zoom)))
    }, 250)
    return () => {
      window.clearTimeout(timer)
      cancelAnimationFrame(raf)
    }
  }, [editor, zoom, layoutVersion, wrapHeight])

  const hitPages = useMemo(() => {
    if (!snap || !hitPositions) return null
    const wrap = editor.view.dom.closest<HTMLElement>('.page-wrap')
    if (!wrap) return null
    const wrapTop = wrap.getBoundingClientRect().top
    const set = new Set<number>()
    for (const pos of hitPositions) {
      try {
        const y = (editor.view.coordsAtPos(pos).top - wrapTop) / zoom
        set.add(pageIndexAtY(snap.pages, y))
      } catch {
        /* position no longer in the document */
      }
    }
    return set
  }, [snap, hitPositions, editor, zoom])

  const pick = (index: number) => {
    if (!snap) return
    const view = editor.view
    const wrap = view.dom.closest<HTMLElement>('.page-wrap')
    const scroller = view.dom.closest<HTMLElement>('.editor-scroll')
    if (!wrap || !scroller) return
    const page = snap.pages[index]
    const wrapTop =
      wrap.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop
    scroller.scrollTop = wrapTop + page.top * zoom - 8
    const bottom = page.top + page.height
    const onPage = (b: ThumbBlock) => b.el.isConnected && b.top + b.height > page.top + 1
    // a block cut by the page top (table, long paragraph) starts on the previous page
    const first =
      snap.blocks.find((b) => onPage(b) && b.top >= page.top - 1 && b.top < bottom) ??
      snap.blocks.find(onPage)
    let pos: number | null = null
    if (first) {
      try {
        pos = view.posAtDOM(first.el, 0)
      } catch {
        pos = null
      }
    }
    if (pos === null) return
    const sel = TextSelection.near(view.state.doc.resolve(pos), 1)
    view.dispatch(view.state.tr.setSelection(sel))
    view.focus()
  }

  if (!snap) return <div className="nav-empty">{t('appNavPagesPending')}</div>
  const scale = Math.max(0.05, Math.min(width - THUMB_GUTTER * 2, THUMB_MAX_W) / snap.pageW)
  const indices = snap.pages.map((_, i) => i).filter((i) => !hitPages || hitPages.has(i))
  if (indices.length === 0) return <div className="nav-empty">{t('appNavNoResults')}</div>
  return (
    <div className="nav-pages">
      {indices.map((i) => (
        <Thumb
          key={i}
          snap={snap}
          index={i}
          scale={scale}
          active={currentPage === i + 1}
          hit={!!hitPages?.has(i)}
          onPick={pick}
        />
      ))}
    </div>
  )
}
