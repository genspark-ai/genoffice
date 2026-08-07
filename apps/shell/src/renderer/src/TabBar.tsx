import { useEffect, useRef, useState } from 'react'
import type { CSSProperties, ReactElement } from 'react'
import type { TabsApi, TabSummary } from '../../shared/tabs-api'
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

function MarkdownIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 240 240" fill="none" aria-hidden="true">
      <rect width="240" height="240" rx="48" fill="#7C3AED" />
      <path d="M52 168V96H76L96 122L116 96H140V168H118V128L96 154L74 128V168H52Z" fill="#fff" />
      <path d="M168 92H184V130H198L176 170L154 130H168V92Z" fill="#fff" />
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
}

export function TabBar() {
  const { t } = useI18n()
  const [tabs, setTabs] = useState<TabSummary[]>([])
  const stripRef = useRef<HTMLDivElement>(null)

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
    void window.aiOfficeTabs.list().then(setTabs)
    return window.aiOfficeTabs.onChanged(setTabs)
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
              <span className="tab-icon">{KIND_ICON[tab.kind]}</span>
              <span className="tab-title">{tab.title}</span>
              {tab.closable && (
                <button
                  className="tab-close"
                  title={t('closeTab')}
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
