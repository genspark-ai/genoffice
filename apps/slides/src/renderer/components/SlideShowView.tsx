/**
 * Slide show — the full-screen playback view.
 *
 * - A black show layer covering the whole window, attempting system full screen (restored on exit)
 * - The playback sequence skips hidden slides (starting from a hidden slide still plays it)
 * - Moving forward plays the target page's transition effect (CSS approximations of fade/push/wipe/split/circle/random)
 * - In-page shape animations (Animations tab): moving forward plays animations step by step, turning the page only when done;
 *   going back/jumping shows the all-animations-finished state
 * - →/space/enter/PgDn/click next step/page; ←/PgUp/right-click previous page; Home/End first/last page;
 *   Esc exits; advancing past the last page shows the "end of show" black screen
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import type { AnimationItem, LinkTargetOp, ShapeKey, TransitionKind } from '../../shared/ipc'
import { AnimatedSlideStage, useAnimPlayer } from './AnimatedSlide'
import { useI18n } from '../i18n/locale'
import { MorphStage } from './MorphStage'
import {
  computePlayOrder,
  finishRehearse,
  formatClock,
  startRehearse,
  switchRehearsePage,
  type RehearseTiming,
} from '../slideshow-utils'

const ANIMATED = [
  'fade',
  'push',
  'wipe',
  'split',
  'circle',
  'cover',
  'pull',
  'dissolve',
  'zoom',
] as const

export function SlideShowView({
  slides,
  images,
  startAt,
  onExit,
  customOrder,
  rehearseMode,
  onRehearseDone,
}: {
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  /** Start page (original index) */
  startAt: number
  /** Exit the show; lastIndex is the original index of the page dwelt on (for locating back in the edit view) */
  onExit: (lastIndex: number) => void
  /** Custom show: a specified playback sequence (original indexes), replacing the default full order when non-empty */
  customOrder?: number[]
  /** Rehearsal timing mode: shows a timer bar at the top and records each page's dwell time */
  rehearseMode?: boolean
  /** Rehearsal-end callback (called before onExit on exit); perPageSec is by original page index, unvisited pages are 0 */
  onRehearseDone?: (perPageSec: number[]) => void
}) {
  const { t } = useI18n()
  // Playback sequence (original indexes): hidden pages skipped (except the start page); custom shows use the given order
  const order = useMemo(
    () => computePlayOrder(slides, startAt, customOrder),
    [slides, startAt, customOrder],
  )
  const [pos, setPos] = useState(() => Math.max(0, order.indexOf(startAt)))
  const [ended, setEnded] = useState(false)
  /** Current transition animation: kind + replay nonce (key change re-triggers the CSS animation) */
  const [anim, setAnim] = useState<{ kind: TransitionKind; nonce: number }>({
    kind: 'none',
    nonce: 0,
  })
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  /** Per-page transition effects (prefetched once when the show starts, zero IPC on page turns) */
  const transRef = useRef<TransitionKind[]>([])
  /** Per-page animation lists (also prefetched once) */
  const [allAnims, setAllAnims] = useState<AnimationItem[][] | null>(null)
  /** Per-page element Morph pairing keys (also prefetched once) */
  const keysRef = useRef<ShapeKey[][]>([])
  /** Per-page element hyperlinks (also prefetched once): sourceId → target; clicks hit-test against these */
  const linksRef = useRef<Array<Map<string, LinkTargetOp>>>([])
  /** Per-page run hyperlinks: "sourceId:para:run" → target; hit-tested against layout glyph runs */
  const runLinksRef = useRef<Array<Map<string, LinkTargetOp>>>([])
  /** Morph tween in progress: previous/target page original indexes + replay nonce */
  const [morph, setMorph] = useState<{ fromIdx: number; toIdx: number; nonce: number } | null>(null)
  /** How the current page was entered: forward = initial state playing step by step, others = all-finished state */
  const navModeRef = useRef<'fresh' | 'all'>('fresh')

  useEffect(() => {
    let cancelled = false
    void Promise.all(slides.map((_, i) => window.slidesApi.getTransition(i))).then((kinds) => {
      if (!cancelled) transRef.current = kinds
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getAnimations(i))).then((lists) => {
      if (!cancelled) setAllAnims(lists)
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getShapeKeys(i))).then((keys) => {
      if (!cancelled) keysRef.current = keys
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getSlideLinks(i))).then((lists) => {
      if (!cancelled)
        linksRef.current = lists.map(
          (list) => new Map(list.map(({ sourceId, target }) => [sourceId, target])),
        )
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getRunLinks(i))).then((lists) => {
      if (!cancelled)
        runLinksRef.current = lists.map(
          (list) =>
            new Map(
              list.map((l) => [`${l.sourceId}:${l.paraIndex}:${l.runIndex}`, l.target] as const),
            ),
        )
    })
    return () => {
      cancelled = true
    }
  }, [slides])

  const slide = slides[order[pos]!]
  const player = useAnimPlayer(slide?.heightPx ?? 540, slide?.widthPx ?? 960)

  // ── Rehearsal timing: start timing the first page on entry; accumulate the previous page's dwell on turn; redraw the timer bar every 500ms ──
  const rehearseRef = useRef<RehearseTiming | null>(null)
  const [, setRehearseTick] = useState(0)
  useEffect(() => {
    if (!rehearseMode) return
    rehearseRef.current = startRehearse(slides.length, order[pos] ?? startAt, Date.now())
    const h = window.setInterval(() => setRehearseTick((n) => n + 1), 500)
    return () => window.clearInterval(h)
    // Initialize only once on entering the show (slides/order don't change during the show)
  }, [rehearseMode]) // eslint-disable-line react-hooks/exhaustive-deps
  const curIdx = order[pos]
  useEffect(() => {
    const t = rehearseRef.current
    if (t && curIdx != null && curIdx !== t.currentIndex) {
      rehearseRef.current = switchRehearsePage(t, curIdx, Date.now())
    }
  }, [curIdx])

  // Load the page's animations when the page changes/prefetch completes (forward = initial state, back/jump = finished state)
  useEffect(() => {
    player.load(allAnims?.[order[pos]!] ?? [], navModeRef.current)
  }, [allAnims, pos, order, player.load]) // eslint-disable-line react-hooks/exhaustive-deps

  const exitRef = useRef(() => {})
  exitRef.current = () => {
    // Rehearsal mode: report each page's dwell seconds before exit (ref nulled to prevent duplicate fullscreenchange triggers)
    const t = rehearseRef.current
    if (rehearseMode && onRehearseDone && t) {
      rehearseRef.current = null
      onRehearseDone(finishRehearse(t, Date.now()))
    }
    onExit(order[Math.min(pos, order.length - 1)] ?? startAt)
  }

  // System full screen: requested on entering the show; exiting/user leaving full screen ends the show.
  // Entry is detected from the fullscreenchange event, not the requestFullscreen promise: inside the
  // shell's WebContentsView the promise can stay pending/reject even though fullscreen engaged, which
  // left the show mounted after Esc (fullscreen gone, show still covering the window).
  useEffect(() => {
    // already-fullscreen mounts (presenter view handing off to the normal show
    // keeps fullscreen) never get a fullscreenchange, so seed from current state
    let entered = !!document.fullscreenElement
    let exitTimer = 0
    void document.documentElement.requestFullscreen?.().catch(() => {})
    const onFsChange = () => {
      if (document.fullscreenElement) {
        entered = true
        window.clearTimeout(exitTimer)
        return
      }
      if (!entered) return
      // Grace window before ending the show: a strict-mode remount briefly drops
      // fullscreen (previous cleanup's exitFullscreen) and re-enters right away —
      // only a loss that sticks means the user actually left fullscreen.
      window.clearTimeout(exitTimer)
      exitTimer = window.setTimeout(() => {
        if (!document.fullscreenElement) exitRef.current()
      }, 150)
    }
    document.addEventListener('fullscreenchange', onFsChange)
    return () => {
      window.clearTimeout(exitTimer)
      document.removeEventListener('fullscreenchange', onFsChange)
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => {})
    }
  }, [])

  useEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight })
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const goTo = useCallback(
    (nextPos: number, animate: boolean) => {
      const target = order[nextPos]
      if (target == null) return
      navModeRef.current = animate ? 'fresh' : 'all'
      const current = order[pos]
      let kind: TransitionKind = 'none'
      if (animate) {
        kind = transRef.current[target] ?? 'none'
        if (kind === 'random') kind = ANIMATED[Math.floor(Math.random() * ANIMATED.length)]!
      }
      if (kind === 'morph' && current != null && current !== target) {
        // Morph: skips the CSS page transition; MorphStage tweens elements from the previous page to the target
        setMorph((m) => ({ fromIdx: current, toIdx: target, nonce: (m?.nonce ?? 0) + 1 }))
        setAnim((a) => ({ kind: 'none', nonce: a.nonce + 1 }))
      } else {
        // Morphs that can't tween (start page/same page) degrade to fade-in
        if (kind === 'morph') kind = 'fade'
        setMorph(null)
        setAnim((a) => ({ kind, nonce: a.nonce + 1 }))
      }
      setPos(nextPos)
    },
    [order, pos],
  )

  const next = useCallback(() => {
    if (ended) {
      exitRef.current()
      return
    }
    // Advance in-page animations first; turn the page only when this page's animations are done
    if (player.advance()) return
    if (pos >= order.length - 1) setEnded(true)
    else goTo(pos + 1, true)
  }, [ended, pos, order.length, goTo, player.advance]) // eslint-disable-line react-hooks/exhaustive-deps

  const prev = useCallback(() => {
    if (ended) {
      setEnded(false)
      return
    }
    if (pos > 0) goTo(pos - 1, false)
  }, [ended, pos, goTo])

  // Element hyperlinks during the show (PowerPoint behavior): a click on a linked element follows
  // the link instead of advancing — slide links (Zoom/jump) go to that page, URLs open in the browser
  const followLink = useCallback(
    (target: LinkTargetOp) => {
      if (target.kind === 'slide') {
        const p = order.indexOf(target.slideIndex)
        // Hidden pages aren't in the play order; ignore jumps to them (matching the skip semantics)
        if (p >= 0) {
          setEnded(false)
          goTo(p, true)
        }
        return
      }
      // Electron routes window.open to the system browser (setWindowOpenHandler denies in-app windows)
      window.open(target.kind === 'url' ? target.url : '', '_blank', 'noreferrer')
    },
    [order, goTo],
  )
  /** Click/hover position → slide-model px → topmost linked element's target (null = no link there) */
  const linkAt = useCallback(
    (e: React.MouseEvent<HTMLElement>): LinkTargetOp | null => {
      const cur = order[pos]
      const links = (cur != null ? linksRef.current[cur] : undefined) ?? new Map()
      const runLinks = (cur != null ? runLinksRef.current[cur] : undefined) ?? new Map()
      if ((!links.size && !runLinks.size) || !slide) return null
      const rect = e.currentTarget.getBoundingClientRect()
      const kx = slide.widthPx / rect.width
      return hitLink(
        slide.nodes,
        (e.clientX - rect.left) * kx,
        (e.clientY - rect.top) * kx,
        links,
        runLinks,
      )
    },
    [order, pos, slide],
  )

  // Keyboard navigation (capture beats the editor's generic shortcuts)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        exitRef.current()
      } else if (
        e.key === 'ArrowRight' ||
        e.key === 'ArrowDown' ||
        e.key === ' ' ||
        e.key === 'Enter' ||
        e.key === 'PageDown'
      ) {
        e.preventDefault()
        next()
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault()
        prev()
      } else if (e.key === 'Home') {
        e.preventDefault()
        setEnded(false)
        goTo(0, false)
      } else if (e.key === 'End') {
        e.preventDefault()
        setEnded(false)
        goTo(order.length - 1, false)
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [next, prev, goTo, order.length])

  if (!slide) return null
  const fitW = Math.round(Math.min(size.w, (size.h * slide.widthPx) / slide.heightPx))
  // Rehearsal timer bar values: current page dwell (including earlier revisit accumulation) + total elapsed
  const rehearse = rehearseRef.current
  const sinceEntered = rehearse ? Date.now() - rehearse.enteredAt : 0
  const rehearseCurMs = rehearse
    ? (rehearse.perPageMs[rehearse.currentIndex] ?? 0) + sinceEntered
    : 0
  const rehearseTotalMs = rehearse
    ? rehearse.perPageMs.reduce((a, b) => a + b, 0) + sinceEntered
    : 0

  return (
    <div
      className="slideshow"
      onClick={next}
      onContextMenu={(e) => {
        e.preventDefault()
        prev()
      }}
    >
      {ended ? (
        <div className="ss-end">{t('paneShowEndedClick')}</div>
      ) : (
        <>
          {morph && slides[morph.fromIdx] && slides[morph.toIdx] ? (
            <div key={`morph-${morph.nonce}`} className="ss-frame">
              <MorphStage
                from={slides[morph.fromIdx]!}
                to={slides[morph.toIdx]!}
                fromKeys={keysRef.current[morph.fromIdx] ?? []}
                toKeys={keysRef.current[morph.toIdx] ?? []}
                images={images}
                width={fitW}
                onDone={() => setMorph(null)}
              />
            </div>
          ) : (
            <div
              key={anim.nonce}
              className={`ss-frame${anim.kind !== 'none' ? ` ss-anim-${anim.kind}` : ''}`}
            >
              <div
                style={{ position: 'relative', width: fitW, margin: '0 auto' }}
                onClick={(e) => {
                  const target = linkAt(e)
                  if (!target) return // Bubbles to the root onClick → next page
                  e.stopPropagation()
                  followLink(target)
                }}
                onMouseMove={(e) => {
                  e.currentTarget.style.cursor = linkAt(e) ? 'pointer' : ''
                }}
              >
                <AnimatedSlideStage
                  slide={slide}
                  images={images}
                  width={fitW}
                  states={player.states}
                />
                <ShowMediaLayer slide={slide} slideIndex={order[pos]!} width={fitW} />
              </div>
            </div>
          )}
          {rehearseMode && rehearse && (
            <div className="ss-rehearse" title={t('paneShowRehearseTip')}>
              <span className="ss-rehearse-cur">⏱ {formatClock(rehearseCurMs)}</span>
              <span className="ss-rehearse-total">
                {t('paneShowRehearseTotal', { time: formatClock(rehearseTotalMs) })}
              </span>
            </div>
          )}
          <div className="ss-counter">
            {pos + 1} / {order.length}
          </div>
        </>
      )}
    </div>
  )
}

/**
 * Topmost linked element at a point (slide-model px). Scans nodes back-to-front (z-order),
 * descends into groups (children boxes are group-local); unlinked overlapping nodes don't block
 * links underneath (lenient, matches the show's forgiving click behavior).
 */
function hitLink(
  nodes: RenderNode[],
  x: number,
  y: number,
  links: Map<string, LinkTargetOp>,
  runLinks: Map<string, LinkTargetOp>,
): LinkTargetOp | null {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const n = nodes[i]!
    const { box } = n
    if (x < box.x || y < box.y || x > box.x + box.w || y > box.y + box.h) continue
    if (n.type === 'group') {
      const inner = hitLink(n.children, x - box.x, y - box.y, links, runLinks)
      if (inner) return inner
    }
    // Run-level link on the glyph under the pointer beats the whole-element link (more specific)
    const text = (n as ShapeRenderNode).text
    if (text && runLinks.size) {
      let para = -1
      for (const ln of text.lines) {
        if (ln.paraStart !== false) para++
        const ly = box.y + text.insets.t + ln.top
        if (y < ly || y > ly + ln.height) continue
        for (const r of ln.runs) {
          if (r.srcRunIdx == null) continue
          const rx = box.x + text.insets.l + r.x
          if (x < rx || x > rx + r.widthPx) continue
          const target = runLinks.get(`${n.sourceId}:${para}:${r.srcRunIdx}`)
          if (target) return target
        }
      }
    }
    const target = links.get(n.sourceId)
    if (target) return target
  }
  return null
}

/**
 * Audio/video layer during the show: DOM players stacked over video/audio nodes (click to
 * play/pause, auto-stop on page turn). In the editor media only has poster frames; a show must
 * be able to play, or decks with video are crippled.
 */
function ShowMediaLayer({
  slide,
  slideIndex,
  width,
}: {
  slide: RenderSlide
  slideIndex: number
  width: number
}) {
  const k = width / slide.widthPx
  const nodes = slide.nodes.filter(
    (n): n is import('@genoffice/pptx-render').PictureRenderNode =>
      n.type === 'picture' && !!(n as import('@genoffice/pptx-render').PictureRenderNode).media,
  )
  const [urls, setUrls] = useState<Record<string, { kind: 'video' | 'audio'; dataUrl: string }>>({})
  const [playing, setPlaying] = useState<Record<string, boolean>>({})
  useEffect(() => {
    let cancelled = false
    setUrls({})
    setPlaying({})
    for (const n of nodes) {
      void window.slidesApi.getMediaData(slideIndex, n.sourceId).then((d) => {
        if (!cancelled && d) setUrls((u) => ({ ...u, [n.sourceId]: d }))
      })
    }
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [slideIndex])
  if (!nodes.length) return null
  return (
    <>
      {nodes.map((n) => {
        const media = urls[n.sourceId]
        if (!media) return null
        const style: React.CSSProperties = {
          position: 'absolute',
          left: n.box.x * k,
          top: n.box.y * k,
          width: n.box.w * k,
          height: n.box.h * k,
          cursor: 'pointer',
        }
        const isPlaying = !!playing[n.sourceId]
        const toggle = (el: HTMLMediaElement | null) => {
          if (!el) return
          if (el.paused) void el.play()
          else el.pause()
        }
        if (media.kind === 'video') {
          return (
            <video
              key={n.sourceId}
              src={media.dataUrl}
              style={style}
              playsInline
              onClick={(e) => {
                e.stopPropagation()
                toggle(e.currentTarget)
              }}
              onPlay={() => setPlaying((p) => ({ ...p, [n.sourceId]: true }))}
              onPause={() => setPlaying((p) => ({ ...p, [n.sourceId]: false }))}
            />
          )
        }
        // Audio: the poster frame is drawn by the canvas; overlay a transparent click layer + play badge
        return (
          <div
            key={n.sourceId}
            style={{ ...style, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
            onClick={(e) => {
              e.stopPropagation()
              toggle(document.getElementById(`ss-audio-${n.sourceId}`) as HTMLMediaElement | null)
            }}
          >
            <audio
              id={`ss-audio-${n.sourceId}`}
              src={media.dataUrl}
              onPlay={() => setPlaying((p) => ({ ...p, [n.sourceId]: true }))}
              onPause={() => setPlaying((p) => ({ ...p, [n.sourceId]: false }))}
            />
            <span
              style={{
                fontSize: Math.min(n.box.w, n.box.h) * k * 0.4,
                lineHeight: 1,
                opacity: 0.85,
              }}
            >
              {isPlaying ? '⏸' : '▶'}
            </span>
          </div>
        )
      })}
    </>
  )
}
