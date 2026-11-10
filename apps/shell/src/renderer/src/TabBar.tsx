import { useEffect, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import type { TabsApi, TabSummary } from '../../shared/tabs-api'
import type { ChromeApi } from '../../shared/chrome-api'
import { TEAR_OFF_SLACK, insertionIndexForX, isBeyondBand } from '../../shared/tab-drag-geometry'
import { notifyFilesChanged } from './file-events'
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

/**
 * Formats hover tooltip for document tab:
 * - Tabs with a saved filePath show the full title on line 1, and the full absolute path on line 2.
 * - Untitled or Home/Help tabs without a filePath show the title alone.
 */
export function formatTabTooltip(tab: TabSummary): string {
  return tab.filePath ? `${tab.title}\n${tab.filePath}` : tab.title
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

function HelpIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#7C5CE0" />
      <path
        d="M120 176c-6.6 0-12-5.2-12-11.6 0-6.5 5.4-11.7 12-11.7s12 5.2 12 11.7c0 6.4-5.4 11.6-12 11.6zm-9.6-34.9c0-16.2 19.2-19.1 19.2-30.4 0-5.6-4.4-9.3-10.6-9.3-6.5 0-11 3.4-13.3 8.7-1.5 3.5-5.1 5.4-8.9 4.6l-8.7-1.8c-4.7-1-7.6-5.7-6.2-10.2C86.5 88.7 99.6 79 119.6 79c20.7 0 34.3 12.3 34.3 29.6 0 23.7-24.4 26.6-24.4 39.4 0 2.5-2.1 4.4-4.7 4.4h-10.1c-2.6 0-4.7-1.9-4.7-4.4z"
        fill="#fff"
      />
    </svg>
  )
}

const IS_MAC = navigator.platform.toLowerCase().includes('mac')

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

function HtmlIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#0FA3A3" />
      <path
        d="M92 72L44 120L92 168M148 72L196 120L148 168"
        stroke="#fff"
        strokeWidth="20"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

const KIND_ICON: Record<TabSummary['kind'], ReactElement> = {
  home: <HomeIcon />,
  docs: <DocIcon />,
  sheets: <SheetIcon />,
  slides: <SlideIcon />,
  pdf: <PdfIcon />,
  markdown: <MarkdownIcon />,
  html: <HtmlIcon />,
  help: <HelpIcon />,
}

/**
 * Extension of a path's last segment, without the dot ('' when it has none).
 * Read off the basename: a dot in a directory name is not an extension, and
 * taking 'v2\Notes' from C:\Users\me.v2\Notes built a rename target that
 * renameFile could not resolve.
 */
export function fileExtension(filePath: string): string {
  const name = filePath.slice(Math.max(filePath.lastIndexOf('\\'), filePath.lastIndexOf('/')) + 1)
  const dot = name.lastIndexOf('.')
  return dot > -1 ? name.slice(dot + 1) : ''
}

// a trailing dot counts: "note." must reach the shell's own rejection, not become "note..md"
function hasExtension(name: string): boolean {
  return name.lastIndexOf('.') > -1
}

/** An extensionless name keeps the current extension; a typed one is taken verbatim (note.md -> note.txt). */
export function renamedFileName(currentPath: string, typedName: string): string {
  if (hasExtension(typedName)) return typedName
  const ext = fileExtension(currentPath)
  return ext ? `${typedName}.${ext}` : typedName
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

  // Double-click a file tab to rename the underlying file inline (Home's row
  // rename, one tab over): the input prefills the base name, Enter/blur commits
  // through the same renameFile IPC (title syncs via tabManager.renameTabFile),
  // Escape cancels. Home tabs and untitled documents cannot be renamed.
  const [renaming, setRenaming] = useState<{ id: string; value: string } | null>(null)
  const renamingRef = useRef(renaming)
  renamingRef.current = renaming
  /** open the inline rename box on a file tab (double-click, or "Rename" in the tab menu) */
  const beginRename = (tab: TabSummary) => {
    if (tab.id === 'home' || !tab.filePath) return
    // whole name, extension included: only the shell may change an extension
    setRenaming({ id: tab.id, value: tab.title })
  }
  const beginRenameRef = useRef(beginRename)
  beginRenameRef.current = beginRename
  useEffect(
    () =>
      window.aiOfficeTabs.onStartRename?.((id) => {
        const tab = tabsRef.current.find((candidate) => candidate.id === id)
        if (tab) beginRenameRef.current(tab)
      }),
    [],
  )
  const commitRename = () => {
    const r = renamingRef.current
    renamingRef.current = null
    setRenaming(null)
    if (!r) return
    const tab = tabsRef.current.find((tb) => tb.id === r.id)
    const value = r.value.trim()
    if (!tab?.filePath || !value) return
    const newName = renamedFileName(tab.filePath, value)
    if (newName === tab.title) return
    void window.aiOffice.renameFile(tab.filePath, newName).then((result) => {
      if (!result.ok) window.alert(result.error ?? t('renameFailed'))
      // Home shares this renderer and only re-pulls on window focus, which the
      // rename input already holds: tell it the recents / folder rows moved.
      else notifyFilesChanged()
    })
  }

  // Chrome-style drag-to-reorder: the grabbed tab tracks the pointer 1:1 while
  // its neighbours slide aside live; the final order is committed on release.
  // Leaving the strip vertically tears the tab off into its own window, which
  // then follows the still-held pointer; coming back over the strip docks it.
  interface DragInfo {
    pointerId: number
    id: string
    from: number
    startX: number
    /** viewport-x left edge + width of every tab, sampled at drag start */
    lefts: number[]
    widths: number[]
    /** the strip's vertical band, sampled at drag start (tear-off threshold) */
    stripTop: number
    stripBottom: number
    target: number
    started: boolean
    /** tear-off requested, main has not answered yet */
    tearing: boolean
    /** the tab is a window now; moves steer it, release lets it go */
    torn: boolean
  }
  const dragRef = useRef<DragInfo | null>(null)
  const [dragVisual, setDragVisual] = useState<{
    id: string
    dx: number
    from: number
    target: number
    width: number
  } | null>(null)
  // a detached window is being dragged over the strip: insertion indicator x
  // (strip-relative) — the main process asks, this strip answers with the slot
  const [dockX, setDockX] = useState<number | null>(null)

  const tabRects = () =>
    Array.from(stripRef.current?.querySelectorAll<HTMLElement>('.tab-item') ?? [], (el) => {
      const r = el.getBoundingClientRect()
      return { left: r.left, width: r.width }
    })

  const finishDrag = (pointerId: number, commit: boolean) => {
    const drag = dragRef.current
    if (!drag || pointerId !== drag.pointerId) return
    dragRef.current = null
    if (drag.torn) {
      // release: the window under the pointer stays where it is and takes focus
      void window.aiOfficeTabs.endTornDrag()
      return
    }
    // released while main is still creating the window: the tearOff callback
    // sees no drag and ends the drag itself
    if (drag.tearing) return
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
    // a torn-off tab leaves the list by design and its drag goes on
    if (drag && !drag.torn && !drag.tearing && !tabs.some((t) => t.id === drag.id)) {
      dragRef.current = null
      setDragVisual(null)
    }
  }, [tabs])

  // A detached window dragged over this strip: place the insertion indicator
  // at the slot under the pointer and tell main which slot that is.
  useEffect(() => {
    return window.aiOfficeTabs.onDockPreview?.((preview) => {
      const strip = stripRef.current
      if (!preview || !strip) {
        setDockX(null)
        return
      }
      const rects = tabRects()
      const index = insertionIndexForX(rects, preview.x)
      window.aiOfficeTabs.reportDockIndex(index)
      const stripLeft = strip.getBoundingClientRect().left - strip.scrollLeft
      const slot = rects[index]
      const last = rects[rects.length - 1]
      const edge = slot ? slot.left : last ? last.left + last.width : preview.x
      setDockX(Math.round(edge - stripLeft))
    })
  }, [])

  /** pointer moves during a drag, delivered to the strip (it holds the capture) */
  const handleDragMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const drag = dragRef.current
    if (!drag || event.pointerId !== drag.pointerId) return
    if (drag.torn) {
      window.aiOfficeTabs.dragTornWindow(event.screenX, event.screenY)
      // back over the strip: the window folds back into a tab at the pointer's slot
      if (!isBeyondBand(event.clientY, drag.stripTop, drag.stripBottom, 0)) {
        dragRef.current = null
        void window.aiOfficeTabs.dockTornWindow(insertionIndexForX(tabRects(), event.clientX))
      }
      return
    }
    if (drag.tearing) return
    let dx = event.clientX - drag.startX
    // 4px dead zone so plain clicks never wiggle the tab
    if (!drag.started) {
      if (Math.abs(dx) < 4 && Math.abs(event.clientY - (drag.stripTop + drag.stripBottom) / 2) < 4)
        return
      // re-sample geometry the moment the drag really starts — the
      // pointer-down activation re-renders and could have moved tabs
      const rects = tabRects()
      drag.lefts = rects.map((r) => r.left)
      drag.widths = rects.map((r) => r.width)
      drag.started = true
    }
    // far enough above or below the strip: tear the tab off into its own window
    if (isBeyondBand(event.clientY, drag.stripTop, drag.stripBottom, TEAR_OFF_SLACK)) {
      drag.tearing = true
      const { screenX, screenY, pointerId } = event
      void window.aiOfficeTabs.tearOff(drag.id, screenX, screenY).then((ok) => {
        const current = dragRef.current
        if (!current || current.pointerId !== pointerId) {
          // released meanwhile: let the window go (it exists only when ok)
          if (ok) void window.aiOfficeTabs.endTornDrag()
          return
        }
        current.tearing = false
        if (!ok) return
        current.torn = true
        setDragVisual(null)
      })
      return
    }
    // keep the tab inside the strip; slot 0 (Home) is off limits
    const last = drag.lefts.length - 1
    const minDx = drag.lefts[1] - drag.lefts[drag.from]
    const maxDx =
      drag.lefts[last] + drag.widths[last] - drag.widths[drag.from] - drag.lefts[drag.from]
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
  }

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
      {!IS_MAC && (
        <button
          className="tab-app-menu-btn"
          title={t('appMenu')}
          aria-label={t('appMenu')}
          onClick={(event) => {
            const rect = event.currentTarget.getBoundingClientRect()
            void window.aiOfficeTabs.showAppMenu(Math.round(rect.left), Math.round(rect.bottom))
          }}
        >
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path
              d="M4 7h16M4 12h16M4 17h16"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
            />
          </svg>
        </button>
      )}
      <div
        className={`tab-strip ${dragVisual ? 'dragging' : ''} ${dockX !== null ? 'dock-target' : ''}`}
        ref={stripRef}
        // the strip, not the grabbed tab, owns the pointer capture: a torn-off
        // tab's element unmounts mid-gesture, and the moves must keep coming
        onPointerMove={handleDragMove}
        onPointerUp={(event) => finishDrag(event.pointerId, true)}
        onPointerCancel={(event) => finishDrag(event.pointerId, false)}
        onLostPointerCapture={(event) => finishDrag(event.pointerId, false)}
      >
        {dockX !== null && (
          <div className="tab-dock-indicator" style={{ left: dockX }} aria-hidden="true" />
        )}
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
              className={`tab-item ${tab.kind === 'home' ? 'tab-home' : ''} ${tab.active ? 'active' : ''} ${tab.dormant ? 'tab-dormant' : ''} ${dragVisual?.id === tab.id ? 'drag-source' : ''}`}
              // long file names ellipsize in the strip — hover reveals the
              // full title and where the file is (the close button's own tooltip still wins there)
              title={
                tab.dormant
                  ? `${formatTabTooltip(tab)} (${t('tabSleeping')})`
                  : formatTabTooltip(tab)
              }
              style={dragStyle}
              onContextMenu={(event) => {
                event.preventDefault()
                if (tab.id === 'home') return
                void window.aiOfficeTabs.showTabMenu(
                  tab.id,
                  Math.round(event.clientX),
                  Math.round(event.clientY),
                )
              }}
              onAuxClick={(event) => {
                if (event.button !== 1 || !tab.closable) return
                event.preventDefault()
                void window.aiOfficeTabs.close(tab.id)
              }}
              onMouseDown={(event) => {
                // middle button: stop Chromium's autoscroll on the scrollable
                // strip so the tab is still under the pointer at auxclick
                if (event.button === 1) {
                  event.preventDefault()
                  return
                }
                // the strip captures the pointer for dragging, which retargets
                // the derived dblclick to the strip; pointer events carry no
                // click count either, so read it off the compat mousedown
                if (event.button !== 0 || event.detail !== 2) return
                if (tab.id === 'home' || !tab.filePath) return
                if ((event.target as HTMLElement).closest('.tab-close')) return
                if ((event.target as HTMLElement).closest('.tab-rename-input')) return
                // the input mounts and autofocuses inside this dispatch; the
                // default mousedown focus step would blur it straight away
                event.preventDefault()
                beginRename(tab)
              }}
              onPointerDown={(event) => {
                if (event.button !== 0) return
                if ((event.target as HTMLElement).closest('.tab-close')) return
                if ((event.target as HTMLElement).closest('.tab-rename-input')) return
                // Chrome-style: pressing a tab activates it immediately, so
                // activation never depends on the click that a drag would eat
                if (!tab.active) void window.aiOfficeTabs.activate(tab.id)
                if (tab.id === 'home') return
                const strip = stripRef.current
                if (!strip) return
                const rects = tabRects()
                const band = strip.getBoundingClientRect()
                dragRef.current = {
                  pointerId: event.pointerId,
                  id: tab.id,
                  from: index,
                  startX: event.clientX,
                  lefts: rects.map((r) => r.left),
                  widths: rects.map((r) => r.width),
                  stripTop: band.top,
                  stripBottom: band.bottom,
                  target: index,
                  started: false,
                  tearing: false,
                  torn: false,
                }
                strip.setPointerCapture(event.pointerId)
              }}
            >
              {/* highlight plate behind the content — hover capsule / active white body */}
              <span className="tab-plate" aria-hidden="true" />
              <span className="tab-icon">{tabIcon(tab)}</span>
              {renaming?.id === tab.id ? (
                <input
                  className="tab-rename-input"
                  autoFocus
                  value={renaming.value}
                  aria-label={t('rename')}
                  spellCheck={false}
                  onClick={(event) => event.stopPropagation()}
                  onDoubleClick={(event) => event.stopPropagation()}
                  onPointerDown={(event) => event.stopPropagation()}
                  onChange={(event) => setRenaming({ id: tab.id, value: event.target.value })}
                  onKeyDown={(event) => {
                    // Enter that confirms an IME candidate is not a commit (Home's rename does the same)
                    if (event.nativeEvent.isComposing) return
                    if (event.key === 'Enter') commitRename()
                    else if (event.key === 'Escape') {
                      renamingRef.current = null
                      setRenaming(null)
                    }
                  }}
                  onBlur={commitRename}
                />
              ) : (
                <span className="tab-title">{tab.title}</span>
              )}
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
      <div className="tab-bar-caption-spacer" />
    </div>
  )
}
