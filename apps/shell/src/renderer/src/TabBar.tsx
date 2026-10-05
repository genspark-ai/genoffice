import { useEffect, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import type { TabsApi, TabSummary } from '../../shared/tabs-api'
import { TEAR_OFF_SLACK, insertionIndexForX, isBeyondBand } from '../../shared/tab-drag-geometry'
import { notifyFilesChanged } from './file-events'
import { useI18n } from './locale'

declare global {
  interface Window {
    aiOfficeTabs: TabsApi
  }
}

function DocIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#3276CD" />
      <path
        d="M183.373 72H172.927C168.749 72 165.267 74.4 164.57 78C149.946 136.8 150.642 135 149.946 139.8C149.946 139.2 149.946 138.6 149.249 137.4C148.553 134.4 149.946 137.4 133.232 78C131.839 74.4 129.053 72 124.875 72H115.822C111.643 72 108.161 74.4 107.465 78C90.7509 137.4 90.7509 135.6 90.0544 139.8V137.4C89.358 134.4 80.3047 93.6 76.8227 78C75.4299 74.4 72.6442 72 68.4658 72H56.6268C51.0556 72 46.8771 76.8 48.2699 81C53.8412 100.8 67.073 147 71.2514 162.6C72.6442 166.2 76.1263 168 79.6083 168H97.0185C101.197 168 104.679 166.2 105.375 162.6L117.911 120L120 114L122.089 120C122.089 120 130.446 150 134.625 162.6C135.321 165.6 138.803 168 142.285 168H159.695C163.177 168 166.659 166.2 168.052 162.6C181.98 113.4 188.944 91.2 191.73 81C193.123 76.8 188.944 72 183.373 72Z"
        fill="#fff"
      />
    </svg>
  )
}

function SheetIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#4FA16B" />
      <path
        d="M96.3863 61.8047C98.5543 61.8052 100.377 63.1591 101.332 65.1562L101.367 65.2148C114.487 90.1591 110.279 82.597 118.664 100.758L120.445 104.637L121.98 100.641C125.532 91.3552 127.882 85.7897 138.621 65.2031L137.031 64.373L138.632 65.2031L138.656 65.1562C139.61 63.1599 141.434 61.8064 143.601 61.8047H160.933C164.595 61.8047 167.33 65.8991 165.75 69.5977L165.375 70.3242L165.339 70.3828C160.568 79.1142 153.7 91.1365 148.019 101.062C145.182 106.02 142.633 110.459 140.8 113.695C139.885 115.311 139.148 116.643 138.632 117.586C138.378 118.052 138.169 118.436 138.023 118.723C137.951 118.864 137.877 119.002 137.824 119.121C137.799 119.178 137.772 119.259 137.742 119.344C137.727 119.386 137.704 119.452 137.683 119.531C137.669 119.588 137.625 119.77 137.625 120V120.469L137.859 120.879L165.351 169.629L165.375 169.676C167.682 173.542 164.837 178.195 160.933 178.195H143.601C141.455 178.194 139.917 177.443 139.007 176.086L138.656 175.477L138.621 175.418L131.425 161.473C126.778 152.195 128.26 153.91 121.945 139.289L120.316 135.504L118.652 139.277C115.953 145.392 114.95 148.635 109.089 160.395L101.367 175.406L101.332 175.477C100.498 177.217 98.8376 178.195 96.3863 178.195H79.6519C75.0237 178.195 72.3782 173.434 74.6246 169.676L74.6363 169.629L102.14 120.879L102.632 120L102.14 119.121L74.6832 70.4648C73.747 68.4964 73.8839 66.2974 74.8003 64.6172C75.6999 62.9712 77.3618 61.8048 79.6519 61.8047H96.3863Z"
        fill="#fff"
      />
    </svg>
  )
}

function PdfIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#EF4444" />
      <path
        d="M102.719 63.0153C105.738 52.1477 126.265 50.9398 128.68 66.6374C131.699 75.6938 127.472 90.7878 125.661 100.448C130.491 113.127 137.133 122.183 147.397 128.22C158.264 127.013 179.395 125.202 186.641 132.447C192.678 138.485 191.471 155.389 175.774 155.389C166.717 155.389 153.434 151.767 141.963 145.126C129.284 147.541 114.19 152.974 100.907 157.804C70.7196 209.727 53.2104 186.181 55.0216 176.521C57.4366 164.446 73.7385 154.786 85.8136 148.749C91.8511 137.277 100.907 117.957 106.944 103.466C102.718 86.5617 100.304 72.6754 102.719 63.0153ZM85.2149 158.437C81.5921 161.456 70.1214 171.117 67.1026 179.569C67.1026 179.569 73.7436 176.55 85.2149 158.437ZM116.605 113.718C112.378 124.586 107.548 136.662 101.511 146.925C111.171 142.699 122.039 137.869 134.718 134.85C127.473 130.02 121.435 122.775 116.605 113.718ZM180.613 143.932C183.028 142.121 179.406 137.291 158.275 139.102C177.595 147.555 180.613 143.932 180.613 143.932ZM116.013 64.2419C114.805 64.2436 114.806 80.5436 117.221 88.9958C120.239 83.5616 120.843 64.2421 116.013 64.2419Z"
        fill="#fff"
      />
    </svg>
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
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#D33922" />
      <path
        d="M130.5 72C152.5 72 167 87.75 167 109.898C167 154.195 122.5 147.797 111 147.797V175.852C111 179.297 108.5 181.758 105 181.758H90C86.5 181.758 84 179.297 84 175.852V77.9062C84 74.4609 86.5 72 90 72H130.5ZM111 124.664H124.5C129 124.664 132.5 123.188 135 120.727C140 114.82 140 104.484 135.5 99.0703C133 96.1172 129.5 95.1328 125 95.1328H111V124.664Z"
        fill="#fff"
      />
    </svg>
  )
}

/* same artwork as the home screen's file-md.svg asset */
function MarkdownIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#8B5CF6" />
      <path
        d="M36.4103 164L44.1723 164C47.2768 164 49.8641 161.8 50.382 158.5C61.2484 104.6 60.7313 106.25 61.2484 101.85C61.2484 102.4 62 98.9167 63.5 101.85C64.5 105.792 61.2484 92.1333 73.6679 146.583C74.7029 149.883 76.7731 152.083 79.8776 152.083L86.6045 152.083C89.7097 152.083 92.297 149.883 92.8142 146.583C103.693 98.8889 103.329 99.9343 104.415 100.62C104.569 100.717 104.752 100.807 105 100.75L105.752 104.05C106.269 106.8 112.996 144.2 115.583 158.5C116.618 161.8 118.688 164 121.793 164L130.59 164C134.729 164 137.834 159.6 136.799 155.75C132.66 137.6 122.828 95.25 119.723 80.95C118.688 77.65 116.101 76 113.513 76L100.576 76C97.4717 76 94.8843 77.65 94.3672 80.95L85.0522 120L83.5 125.5L81.9477 120C81.9477 120 75.738 92.5 72.6328 80.95C72.1156 78.2 69.5283 76 66.941 76L54.0044 76C51.4171 76 48.8298 77.65 47.7947 80.95C37.4454 126.05 32.2708 146.4 30.2006 155.75C29.1655 159.6 32.2708 164 36.4103 164Z"
        fill="#fff"
      />
      <path
        d="M168 82C168 78.6863 170.686 76 174 76H184C187.314 76 190 78.6863 190 82V142H168V82Z"
        fill="#fff"
      />
      <path
        d="M175.248 162.741C177.239 165.001 180.761 165.001 182.752 162.741L208.684 133.305C211.528 130.076 209.236 125 204.932 125H153.068C148.765 125 146.472 130.076 149.316 133.305L175.248 162.741Z"
        fill="#fff"
      />
    </svg>
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

/**
 * True when the typed name already spells an extension.
 *
 * A trailing dot counts: "note." is a name the shell rejects outright (Windows
 * strips trailing dots, so it would land somewhere else), and this must not
 * paper over that by turning it into "note..md". A leading dot counts too, to
 * agree with fileExtension about where a name begins.
 */
function hasExtension(name: string): boolean {
  return name.lastIndexOf('.') > -1
}

/**
 * The file name a tab rename actually commits.
 *
 * The input holds the whole name, extension included, because the shell owns
 * which extensions an app can open and only the shell may change one. Someone
 * who types a bare "Meeting" is not asking to drop the extension — they just
 * are not repeating it — so an extensionless name keeps the file's current one.
 * A name that does carry an extension is taken verbatim, which is what lets
 * note.md become note.txt and back.
 */
export function renamedFileName(currentPath: string, typedName: string): string {
  if (hasExtension(typedName)) return typedName
  const ext = fileExtension(currentPath)
  return ext ? `${typedName}.${ext}` : typedName
}

export function TabBar() {
  const { t } = useI18n()
  const [tabs, setTabs] = useState<TabSummary[]>([])
  const stripRef = useRef<HTMLDivElement>(null)

  // Double-click a file tab to rename the underlying file inline (Home's row
  // rename, one tab over): the input prefills the base name, Enter/blur commits
  // through the same renameFile IPC (title syncs via tabManager.renameTabFile),
  // Escape cancels. Home tabs and untitled documents cannot be renamed.
  const [renaming, setRenaming] = useState<{ id: string; value: string } | null>(null)
  const renamingRef = useRef(renaming)
  renamingRef.current = renaming
  const tabsRef = useRef(tabs)
  tabsRef.current = tabs
  const commitRename = () => {
    const r = renamingRef.current
    renamingRef.current = null
    setRenaming(null)
    if (!r) return
    const tab = tabsRef.current.find((tb) => tb.id === r.id)
    const value = r.value.trim()
    if (!tab?.filePath || !value) return
    // a name typed without an extension keeps the file's current one; see
    // renamedFileName. Compared against the resolved name so retyping the base
    // half of the current title is a no-op rather than a pointless IPC.
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
    void window.aiOfficeTabs.list().then(setTabs)
    return window.aiOfficeTabs.onChanged(setTabs)
  }, [])

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
              className={`tab-item ${tab.kind === 'home' ? 'tab-home' : ''} ${tab.active ? 'active' : ''} ${dragVisual?.id === tab.id ? 'drag-source' : ''}`}
              // long file names ellipsize in the strip — hover reveals the
              // full title (the close button's own tooltip still wins there)
              title={tab.title}
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
              onMouseDown={(event) => {
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
                // the whole name, extension included: the shell owns which
                // extensions an app can open, so only it may change one. A
                // name typed without an extension keeps the current one.
                setRenaming({ id: tab.id, value: tab.title })
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
              <span className="tab-icon">{KIND_ICON[tab.kind]}</span>
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
          className="tab-new-btn"
          title={t('newTab')}
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
