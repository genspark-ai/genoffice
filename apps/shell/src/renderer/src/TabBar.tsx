import { useEffect, useRef, useState } from 'react'
import type { CSSProperties, ReactElement } from 'react'
import type { TabsApi, TabSummary } from '../../shared/tabs-api'
import type { ChromeApi } from '../../shared/chrome-api'
import { useI18n } from './locale'
import { useToast } from './toast'
import { onShortcut } from './shortcuts'
import type { ShortcutId } from './shortcuts'
import { useShortcutHint } from './use-shortcut-hint'
import iconDocx from './assets/file-docx.svg'
import iconXlsx from './assets/file-xlsx.svg'
import iconPptx from './assets/file-pptx.svg'
import iconPdf from './assets/file-pdf.svg'
import iconMd from './assets/file-md.svg'
import iconJson from './assets/file-json.svg'
import iconXml from './assets/file-xml.svg'
import iconYaml from './assets/file-yaml.svg'
import iconSql from './assets/file-sql.svg'
import iconCsv from './assets/file-csv.svg'
import iconTsv from './assets/file-tsv.svg'
import iconTxt from './assets/file-txt.svg'

declare global {
  interface Window {
    aiOfficeTabs: TabsApi
    aiOfficeChrome?: ChromeApi
  }
}

function DocIcon() {
  return (
    <img
      src={iconDocx}
      width={16}
      height={16}
      alt=""
      aria-hidden="true"
      style={{ objectFit: 'contain', display: 'block' }}
    />
  )
}

function SheetIcon() {
  return (
    <img
      src={iconXlsx}
      width={16}
      height={16}
      alt=""
      aria-hidden="true"
      style={{ objectFit: 'contain', display: 'block' }}
    />
  )
}

function PdfIcon() {
  return (
    <img
      src={iconPdf}
      width={16}
      height={16}
      alt=""
      aria-hidden="true"
      style={{ objectFit: 'contain', display: 'block' }}
    />
  )
}

function HomeIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M9.06163 4.82633L3.23911 9.92134C2.7398 10.3583 3.07458 11.1343 3.76238 11.1343C4.18259 11.1343 4.52324 11.4489 4.52324 11.8371V15.0806C4.52324 17.871 4.52324 19.2662 5.46176 20.1331C6.40029 21 7.91082 21 10.9319 21H13.0681C16.0892 21 17.5997 21 18.5382 20.1331C19.4768 19.2662 19.4768 17.871 19.4768 15.0806V11.8371C19.4768 11.4489 19.8174 11.1343 20.2376 11.1343C20.9254 11.1343 21.2602 10.3583 20.7609 9.92134L14.9383 4.82633C13.5469 3.60878 12.8512 3 12 3C11.1488 3 10.4531 3.60878 9.06163 4.82633Z"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M12 16.0011H12.0105"
        stroke="currentColor"
        strokeWidth="2.57143"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function SlideIcon() {
  return (
    <img
      src={iconPptx}
      width={16}
      height={16}
      alt=""
      aria-hidden="true"
      style={{ objectFit: 'contain', display: 'block' }}
    />
  )
}

function MarkdownIcon() {
  return (
    <img
      src={iconMd}
      width={16}
      height={16}
      alt=""
      aria-hidden="true"
      style={{ objectFit: 'contain', display: 'block' }}
    />
  )
}

const KIND_ICON: Record<TabSummary['kind'], ReactElement> = {
  home: <HomeIcon />,
  docs: <DocIcon />,
  sheets: <SheetIcon />,
  slides: <SlideIcon />,
  pdf: <PdfIcon />,
  markdown: <MarkdownIcon />,
}

function tabIcon(tab: TabSummary): ReactElement {
  if (tab.kind === 'home') return <HomeIcon />
  const lower = tab.title.toLowerCase()
  if (lower.endsWith('.csv')) {
    return (
      <img
        src={iconCsv}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.tsv')) {
    return (
      <img
        src={iconTsv}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.json')) {
    return (
      <img
        src={iconJson}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.xml')) {
    return (
      <img
        src={iconXml}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.yaml') || lower.endsWith('.yml')) {
    return (
      <img
        src={iconYaml}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.sql')) {
    return (
      <img
        src={iconSql}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  if (lower.endsWith('.txt')) {
    return (
      <img
        src={iconTxt}
        width={16}
        height={16}
        alt=""
        aria-hidden="true"
        style={{ objectFit: 'contain', display: 'block' }}
      />
    )
  }
  return KIND_ICON[tab.kind] ?? <DocIcon />
}

/** Windows: the application menu lives in the tab strip (native menu row is hidden) */
function WindowMenus() {
  const [labels, setLabels] = useState<string[]>([])
  useEffect(() => {
    const api = window.aiOfficeChrome
    if (!api) return
    void api.menuLabels().then(setLabels)
    return api.onMenuChanged(setLabels)
  }, [])
  if (labels.length === 0) return null
  return (
    <div className="tab-menus" role="menubar">
      {labels.map((label, index) => (
        <button
          key={`${index}-${label}`}
          type="button"
          role="menuitem"
          aria-haspopup="menu"
          className="tab-menu-btn"
          onClick={(event) => {
            const rect = event.currentTarget.getBoundingClientRect()
            void window.aiOfficeChrome?.popupMenu(index, rect.left, rect.bottom)
          }}
        >
          {label}
        </button>
      ))}
    </div>
  )
}

/** most recent closed tabs kept for Reopen Closed Tab */
const CLOSED_HISTORY = 20

export function TabBar() {
  const { t } = useI18n()
  const toast = useToast()
  const newTabHint = useShortcutHint('newTab')
  const [tabs, setTabs] = useState<TabSummary[]>([])
  const stripRef = useRef<HTMLDivElement>(null)
  const newBtnRef = useRef<HTMLButtonElement>(null)
  const tabsRef = useRef<TabSummary[]>([])
  tabsRef.current = tabs
  const closedRef = useRef<Pick<TabSummary, 'kind' | 'title'>[]>([])

  // Chrome-style drag-to-reorder: the grabbed tab tracks the pointer 1:1 while
  // its neighbours slide aside live; the final order is committed on release.
  interface DragInfo {
    pointerId: number
    id: string
    from: number
    startX: number
    /** viewport-x left edge + width of every tab, sampled at drag start */
    lefts: number[]
    widths: number[]
    target: number
    started: boolean
  }
  const dragRef = useRef<DragInfo | null>(null)
  const [dragVisual, setDragVisual] = useState<{
    id: string
    dx: number
    from: number
    target: number
    width: number
  } | null>(null)

  const finishDrag = (pointerId: number, commit: boolean) => {
    const drag = dragRef.current
    if (!drag || pointerId !== drag.pointerId) return
    dragRef.current = null
    if (!drag.started) {
      // plain click: the in-view scroll was suppressed while the press was
      // held (dragRef was set), so honor it now that the press is over
      stripRef.current
        ?.querySelector('.tab-item.active')
        ?.scrollIntoView({ inline: 'nearest', block: 'nearest' })
      return
    }
    setDragVisual(null)
    if (commit && drag.target !== drag.from) {
      // optimistic local reorder so clearing the transforms causes no flash;
      // the main-process broadcast arrives with the identical order
      setTabs((prev) => {
        // look the tab up by id — the list may have changed mid-drag (e.g.
        // Cmd+W), which would make the indices captured at pointer-down stale
        const fromIdx = prev.findIndex((tb) => tb.id === drag.id)
        if (fromIdx < 0) return prev
        const next = [...prev]
        const [moved] = next.splice(fromIdx, 1)
        next.splice(Math.min(Math.max(drag.target, 1), next.length), 0, moved)
        return next
      })
      void window.aiOfficeTabs.reorder(drag.id, drag.target)
    }
  }

  useEffect(() => {
    let prev: TabSummary[] = []
    // remember tabs that disappeared between two snapshots so they can be reopened
    const apply = (next: TabSummary[]): void => {
      for (const tab of prev) {
        if (tab.kind !== 'home' && !next.some((n) => n.id === tab.id)) {
          closedRef.current = [...closedRef.current, { kind: tab.kind, title: tab.title }].slice(
            -CLOSED_HISTORY,
          )
        }
      }
      prev = next
      setTabs(next)
    }
    void window.aiOfficeTabs.list().then(apply)
    return window.aiOfficeTabs.onChanged(apply)
  }, [])

  // ── shortcuts owned by the tab strip: new tab, reopen closed tab, previous / next tab ──
  const reopenClosed = async (): Promise<void> => {
    const entry = closedRef.current.pop()
    if (!entry) {
      toast.show({ message: t('toastNoClosedTab') })
      return
    }
    try {
      // saved files are found again by name in the recents; the tab title is the file name
      const title = entry.title.toLowerCase()
      const page = await window.aiOffice.recents({ query: entry.title, limit: 25 })
      const hit = page.entries.find((e) => {
        const name = e.name.toLowerCase()
        return name === title || name.replace(/\.[^.]+$/, '') === title
      })
      if (hit) {
        await window.aiOffice.openPath(hit.path)
        return
      }
      // an unsaved/untitled tab has no file to return to: start a fresh one of the same kind
      if (entry.kind === 'docs') await window.aiOffice.newDoc()
      else if (entry.kind === 'sheets') await window.aiOffice.newSheet()
      else if (entry.kind === 'slides') await window.aiOffice.newSlide()
      else if (entry.kind === 'markdown') await window.aiOffice.newMarkdown()
      else toast.show({ message: t('toastReopenFailed', { name: entry.title }) })
    } catch {
      toast.show({ message: t('toastReopenFailed', { name: entry.title }) })
    }
  }
  const tabShortcutRef = useRef<(id: ShortcutId) => boolean>(() => false)
  tabShortcutRef.current = (id) => {
    if (id === 'newTab') {
      const rect = newBtnRef.current?.getBoundingClientRect()
      void window.aiOfficeTabs.showNewMenu(Math.round(rect?.left ?? 120), Math.round(rect?.bottom ?? 40))
      return true
    }
    if (id === 'reopenClosedTab') {
      void reopenClosed()
      return true
    }
    if (id === 'previousTab' || id === 'nextTab') {
      const list = tabsRef.current
      if (list.length < 2) return true
      const at = Math.max(
        0,
        list.findIndex((tab) => tab.active),
      )
      const to = (at + (id === 'nextTab' ? 1 : -1) + list.length) % list.length
      void window.aiOfficeTabs.activate(list[to]!.id)
      return true
    }
    return false
  }
  useEffect(() => onShortcut((id) => tabShortcutRef.current(id)), [])

  // document tabs are sibling WebContentsViews: they see neither this press
  // nor a focus change, so relay it for them to dismiss open popovers
  useEffect(() => {
    const notify = (): void => window.aiOfficeTabs.notifyChromePressed?.()
    document.addEventListener('pointerdown', notify, true)
    return () => document.removeEventListener('pointerdown', notify, true)
  }, [])

  // if the dragged tab is closed mid-drag (e.g. Cmd+W) its element unmounts
  // and pointerup/pointercancel never fire — clear the drag state ourselves
  useEffect(() => {
    const drag = dragRef.current
    if (drag && !tabs.some((t) => t.id === drag.id)) {
      dragRef.current = null
      setDragVisual(null)
    }
  }, [tabs])

  // Trackpads scroll the strip natively; map a mouse's vertical wheel to
  // horizontal scrolling. Native listener because React registers wheel as
  // passive, which forbids preventDefault.
  useEffect(() => {
    const strip = stripRef.current
    if (!strip) return
    const onWheel = (event: WheelEvent) => {
      if (strip.scrollWidth <= strip.clientWidth) return
      if (Math.abs(event.deltaY) <= Math.abs(event.deltaX)) return
      event.preventDefault()
      strip.scrollLeft += event.deltaY
    }
    strip.addEventListener('wheel', onWheel, { passive: false })
    return () => strip.removeEventListener('wheel', onWheel)
  }, [])

  // keep the active tab in view — new tabs open at the far end of the strip
  const activeId = tabs.find((tab) => tab.active)?.id
  useEffect(() => {
    // pointer-down activation runs while the user is pressing that tab — it is
    // already visible, and scrolling the strip mid-press would invalidate the
    // drag geometry sampled at pointer-down
    if (dragRef.current) return
    stripRef.current
      ?.querySelector('.tab-item.active')
      ?.scrollIntoView({ inline: 'nearest', block: 'nearest' })
  }, [activeId])

  return (
    <div className="tab-bar">
      <div className="tab-bar-drag-spacer" />
      <WindowMenus />
      <div className={dragVisual ? 'tab-strip dragging' : 'tab-strip'} ref={stripRef}>
        {tabs.map((tab, index) => {
          // live transforms: the grabbed tab tracks the pointer; tabs between
          // the origin and the current target slide aside by the grabbed width
          let dragStyle: CSSProperties | undefined
          if (dragVisual) {
            if (dragVisual.id === tab.id) {
              dragStyle = { transform: `translateX(${dragVisual.dx}px)` }
            } else if (dragVisual.target <= index && index < dragVisual.from) {
              dragStyle = { transform: `translateX(${dragVisual.width}px)` }
            } else if (dragVisual.from < index && index <= dragVisual.target) {
              dragStyle = { transform: `translateX(-${dragVisual.width}px)` }
            }
          }
          return (
            <div
              key={tab.id}
              className={`tab-item ${tab.kind === 'home' ? 'tab-home' : ''} ${tab.active ? 'active' : ''} ${dragVisual?.id === tab.id ? 'drag-source' : ''}`}
              style={dragStyle}
              onContextMenu={(event) => {
                event.preventDefault()
                void window.aiOfficeTabs.showTabMenu(tab.id, event.clientX, event.clientY)
              }}
              onPointerDown={(event) => {
                if (event.button !== 0) return
                if ((event.target as HTMLElement).closest('.tab-close')) return
                // Chrome-style: pressing a tab activates it immediately, so
                // activation never depends on the click that a drag would eat
                if (!tab.active) void window.aiOfficeTabs.activate(tab.id)
                if (tab.id === 'home') return
                const strip = stripRef.current
                if (!strip) return
                const rects = Array.from(strip.querySelectorAll<HTMLElement>('.tab-item'), (el) =>
                  el.getBoundingClientRect(),
                )
                dragRef.current = {
                  pointerId: event.pointerId,
                  id: tab.id,
                  from: index,
                  startX: event.clientX,
                  lefts: rects.map((r) => r.left),
                  widths: rects.map((r) => r.width),
                  target: index,
                  started: false,
                }
                event.currentTarget.setPointerCapture(event.pointerId)
              }}
              onPointerMove={(event) => {
                const drag = dragRef.current
                if (!drag || event.pointerId !== drag.pointerId) return
                let dx = event.clientX - drag.startX
                // 4px dead zone so plain clicks never wiggle the tab
                if (!drag.started) {
                  if (Math.abs(dx) < 4) return
                  // re-sample geometry the moment the drag really starts — the
                  // pointer-down activation re-renders and could have moved tabs
                  const strip = stripRef.current
                  if (strip) {
                    const rects = Array.from(
                      strip.querySelectorAll<HTMLElement>('.tab-item'),
                      (el) => el.getBoundingClientRect(),
                    )
                    drag.lefts = rects.map((r) => r.left)
                    drag.widths = rects.map((r) => r.width)
                  }
                  drag.started = true
                }
                // keep the tab inside the strip; slot 0 (Home) is off limits
                const last = drag.lefts.length - 1
                const minDx = drag.lefts[1] - drag.lefts[drag.from]
                const maxDx =
                  drag.lefts[last] +
                  drag.widths[last] -
                  drag.widths[drag.from] -
                  drag.lefts[drag.from]
                dx = Math.min(Math.max(dx, minDx), Math.max(minDx, maxDx))
                // Chrome's rule: swap once the grabbed tab's leading edge crosses
                // a neighbour's midpoint (the clamped center can only ever *touch*
                // the first slot's midpoint, so edge-based tests have no dead spot)
                const draggedLeft = drag.lefts[drag.from] + dx
                const draggedRight = draggedLeft + drag.widths[drag.from]
                let target = drag.from
                for (let i = 1; i < drag.from; i++) {
                  if (draggedLeft < drag.lefts[i] + drag.widths[i] / 2) {
                    target = i
                    break
                  }
                }
                for (let i = last; i > drag.from; i--) {
                  if (draggedRight > drag.lefts[i] + drag.widths[i] / 2) {
                    target = i
                    break
                  }
                }
                drag.target = target
                setDragVisual({
                  id: drag.id,
                  dx,
                  from: drag.from,
                  target,
                  width: drag.widths[drag.from],
                })
              }}
              onPointerUp={(event) => finishDrag(event.pointerId, true)}
              onPointerCancel={(event) => finishDrag(event.pointerId, false)}
              onLostPointerCapture={(event) => finishDrag(event.pointerId, false)}
            >
              {/* highlight plate behind the content — hover capsule / active white body */}
              <span className="tab-plate" aria-hidden="true" />
              <span className="tab-icon">{tabIcon(tab)}</span>
              <span className="tab-title">{tab.title}</span>
              {tab.closable && (
                <button
                  className="tab-close"
                  title={t('closeTab')}
                  aria-label={t('closeTab')}
                  onClick={(event) => {
                    event.stopPropagation()
                    void window.aiOfficeTabs.close(tab.id)
                  }}
                >
                  ×
                </button>
              )}
            </div>
          )
        })}
        <button
          ref={newBtnRef}
          className="tab-new-btn"
          title={newTabHint ? `${t('newTab')} (${newTabHint})` : t('newTab')}
          aria-label={t('newTab')}
          onClick={(event) => {
            const rect = event.currentTarget.getBoundingClientRect()
            void window.aiOfficeTabs.showNewMenu(Math.round(rect.left), Math.round(rect.bottom))
          }}
        >
          <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true">
            <path
              d="M12 4.286v15.429M4.286 12h15.429"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </button>
      </div>
      <button
        className="tab-overflow-btn"
        title={t('tabList')}
        aria-label={t('tabList')}
        onClick={(event) => {
          const rect = event.currentTarget.getBoundingClientRect()
          void window.aiOfficeTabs.showMenu(Math.round(rect.left), Math.round(rect.bottom))
        }}
      >
        {/* window-with-tab-bar glyph: slanted tab cells above a full-width
            header divider (from design asset tab.svg) */}
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path
            d="M21 4H3C2.44772 4 2 4.44772 2 5V19C2 19.5523 2.44772 20 3 20H21C21.5523 20 22 19.5523 22 19V5C22 4.44772 21.5523 4 21 4Z"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinejoin="round"
          />
          <path
            d="M11.5 9.5H22M11.5 9.5L9.5 4M17.5 9.5L15.5 4M2 19V8.5M22 19V8.5M4.5 20H19.5"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
    </div>
  )
}
