/**
 * Presenter view for slide shows.
 *
 * Multi-screen: on entry the main process detects displays and opens a full-screen audience
 * show window (AudienceView) on the external screen. This view is the only driver — page
 * number/animation cursor/blackout/ink are all broadcast over IPC (absolute state, the audience
 * side seeks idempotently), and audience clicks/keys are sent back here for arbitration.
 * Single screen: this window only.
 *
 * Layout: top bar (end show/swap displays/switch to normal show),
 * left column timer (pause/reset) + clock → big frame → toolbar (pen/laser/clear ink/blackout)
 * + page progress navigation, right column next-slide preview + notes (adjustable font size),
 * bottom full-width thumbnail strip with click-to-jump.
 * Keyboard: →/space/enter/PgDn next step; ←/PgUp previous page; Home/End first/last; B blackout; Esc exit.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem, ShowSyncState } from '../../shared/ipc'
import { AnimatedSlideStage, useAnimPlayer } from './AnimatedSlide'
import { useI18n } from '../i18n/locale'
import { SlideThumb } from '../SlideThumb'
import { InkLayer, type InkStroke } from './ShowInk'

/** Layout constants (aligned with styles.css) */
const SIDE_W = 340
const TOP_H = 44
const TIMER_H = 40
const BOTTOM_H = 56
const FILM_H = 118

const PEN_COLOR = '#e53935'

/** Elapsed show seconds → mm:ss (past 1 hour minutes keep accumulating) */
function fmtElapsed(sec: number): string {
  const m = Math.floor(sec / 60)
  const s = sec % 60
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
}

function fmtWallClock(): string {
  const d = new Date()
  return `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`
}

export function PresenterView({
  slides,
  images,
  startAt,
  onExit,
  onUseSlideShow,
}: {
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  /** Start page (original index) */
  startAt: number
  /** Exit presenter view; lastIndex is the original index of the page dwelt on (for locating back in the edit view) */
  onExit: (lastIndex: number) => void
  /** "Switch to normal show": exit this view and start a normal show from lastIndex */
  onUseSlideShow?: (lastIndex: number) => void
}) {
  const { t } = useI18n()
  // Playback sequence (original indexes): hidden pages skipped; the start page kept even if hidden (consistent with SlideShowView)
  const order = useMemo(() => {
    const o = slides.map((_, i) => i).filter((i) => !slides[i]!.hidden || i === startAt)
    return o.length > 0 ? o : [startAt]
  }, [slides, startAt])
  const [pos, setPos] = useState(() => Math.max(0, order.indexOf(startAt)))
  const [ended, setEnded] = useState(false)
  const [black, setBlack] = useState(false)
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  /** Per-page animation lists + notes (prefetched once on entry, zero IPC on page turns) */
  const [allAnims, setAllAnims] = useState<AnimationItem[][] | null>(null)
  const [allNotes, setAllNotes] = useState<string[]>([])
  /** How the current page was entered: forward = initial state playing step by step, others = all-finished state */
  const navModeRef = useRef<'fresh' | 'all'>('fresh')
  /** Whether an external audience window was opened (swap-displays button availability) */
  const [hasAudience, setHasAudience] = useState(false)

  // ── Timer (pausable/resettable) and clock ──────────────────────────────────
  const [elapsed, setElapsed] = useState(0)
  const [paused, setPaused] = useState(false)
  const [clock, setClock] = useState(fmtWallClock)
  useEffect(() => {
    if (paused) return
    const t = window.setInterval(() => setElapsed((s) => s + 1), 1000)
    return () => window.clearInterval(t)
  }, [paused])
  useEffect(() => {
    const t = window.setInterval(() => setClock(fmtWallClock()), 10_000)
    return () => window.clearInterval(t)
  }, [])

  // ── Notes font size ──────────────────────────────────────────────────
  const [noteSize, setNoteSize] = useState(15)

  // ── Ink/laser pointer ───────────────────────────────────────────────
  const [tool, setTool] = useState<'none' | 'pen' | 'laser'>('none')
  const [strokes, setStrokes] = useState<InkStroke[]>([])
  const [laser, setLaser] = useState<{ x: number; y: number } | null>(null)
  const stageboxRef = useRef<HTMLDivElement>(null)
  const drawingRef = useRef(false)
  const laserSentAtRef = useRef(0)

  useEffect(() => {
    let cancelled = false
    void Promise.all(slides.map((_, i) => window.slidesApi.getAnimations(i))).then((lists) => {
      if (!cancelled) setAllAnims(lists)
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getNotes(i))).then((notes) => {
      if (!cancelled) setAllNotes(notes)
    })
    return () => {
      cancelled = true
    }
  }, [slides])

  const slide = slides[order[pos]!]
  const player = useAnimPlayer(slide?.heightPx ?? 540, slide?.widthPx ?? 960)

  // Load the page's animations when the page changes/prefetch completes (forward = initial state, back/jump = finished state);
  // loadedRef ticks with player.epoch; the broadcast effect uses it to read the "in place" page number
  const loadedRef = useRef<{ idx: number; fresh: boolean }>({ idx: order[pos]!, fresh: true })
  useEffect(() => {
    loadedRef.current = { idx: order[pos]!, fresh: navModeRef.current === 'fresh' }
    player.load(allAnims?.[order[pos]!] ?? [], navModeRef.current)
  }, [allAnims, pos, order, player.load]) // eslint-disable-line react-hooks/exhaustive-deps

  // ── Multi-screen: open the audience window on entry, close on exit ─────────────
  useEffect(() => {
    let disposed = false
    void window.slidesApi.presenterStart().then((r) => {
      if (!disposed) setHasAudience(r.audience)
    })
    return () => {
      disposed = true
      void window.slidesApi.presenterEnd()
    }
  }, [])

  // Broadcast show state: epoch change = the new page's animations are loaded and loadedRef matches the player cursor
  useEffect(() => {
    const state: ShowSyncState = {
      idx: loadedRef.current.idx,
      fresh: loadedRef.current.fresh,
      played: player.played,
      playing: player.playing,
      ended,
      black,
    }
    window.slidesApi.presenterSync(state)
  }, [player.epoch, player.played, player.playing, ended, black])

  // Clear ink on page turn (the audience side clears in sync)
  useEffect(() => {
    setStrokes([])
    setLaser(null)
    window.slidesApi.presenterInk({ type: 'clear' })
  }, [pos])

  const exitRef = useRef(() => {})
  exitRef.current = () => onExit(order[Math.min(pos, order.length - 1)] ?? startAt)

  // System full screen: requested on entry; kept when switching to normal show (SlideShowView takes over seamlessly)
  const keepFsRef = useRef(false)
  useEffect(() => {
    void document.documentElement.requestFullscreen?.().catch(() => {})
    return () => {
      if (!keepFsRef.current && document.fullscreenElement)
        void document.exitFullscreen().catch(() => {})
    }
  }, [])

  useEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight })
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const goTo = useCallback((nextPos: number, fresh: boolean) => {
    navModeRef.current = fresh ? 'fresh' : 'all'
    setPos(nextPos)
  }, [])

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

  // Navigation sent back from the audience window (clicks/keys)
  const nextRef = useRef(next)
  nextRef.current = next
  const prevRef = useRef(prev)
  prevRef.current = prev
  useEffect(
    () =>
      window.slidesApi.onAudienceNav((action) => {
        if (action === 'next') nextRef.current()
        else if (action === 'prev') prevRef.current()
        else exitRef.current()
      }),
    [],
  )

  // Keyboard navigation (capture beats the editor's generic shortcuts; keys match SlideShowView + B blackout)
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
      } else if (e.key === 'b' || e.key === 'B') {
        e.preventDefault()
        setBlack((v) => !v)
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [next, prev, goTo, order.length])

  // Scroll the current thumbnail into view
  const filmRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    filmRef.current
      ?.querySelector('.pv-film-cur')
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' })
  }, [pos])

  const clearInk = useCallback(() => {
    setStrokes([])
    setLaser(null)
    window.slidesApi.presenterInk({ type: 'clear' })
  }, [])

  // ── Ink pointer events (normalized coordinates, each side restores its own) ────────────
  const normPoint = useCallback((e: React.PointerEvent): { x: number; y: number } | null => {
    const r = stageboxRef.current?.getBoundingClientRect()
    if (!r || r.width === 0 || r.height === 0) return null
    return {
      x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)),
      y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)),
    }
  }, [])

  const onStagePointerDown = useCallback(
    (e: React.PointerEvent) => {
      if (tool !== 'pen') return
      const p = normPoint(e)
      if (!p) return
      e.currentTarget.setPointerCapture(e.pointerId)
      drawingRef.current = true
      setStrokes((ss) => [...ss, { color: PEN_COLOR, points: [p.x, p.y] }])
      window.slidesApi.presenterInk({ type: 'stroke-start', x: p.x, y: p.y, color: PEN_COLOR })
    },
    [tool, normPoint],
  )

  const onStagePointerMove = useCallback(
    (e: React.PointerEvent) => {
      if (tool === 'pen' && drawingRef.current) {
        const p = normPoint(e)
        if (!p) return
        setStrokes((ss) => {
          const last = ss[ss.length - 1]
          if (!last) return ss
          return [...ss.slice(0, -1), { ...last, points: [...last.points, p.x, p.y] }]
        })
        window.slidesApi.presenterInk({ type: 'stroke-move', x: p.x, y: p.y })
      } else if (tool === 'laser') {
        const p = normPoint(e)
        if (!p) return
        setLaser(p)
        const now = performance.now()
        if (now - laserSentAtRef.current > 30) {
          laserSentAtRef.current = now
          window.slidesApi.presenterInk({ type: 'laser', x: p.x, y: p.y })
        }
      }
    },
    [tool, normPoint],
  )

  const onStagePointerUp = useCallback(() => {
    drawingRef.current = false
  }, [])

  const onStagePointerLeave = useCallback(() => {
    if (tool === 'laser') {
      setLaser(null)
      window.slidesApi.presenterInk({ type: 'laser', x: -1, y: -1 })
    }
  }, [tool])

  if (!slide) return null

  // Left main area auto-fit: subtract the right column/top bar/timer row/toolbar/thumbnail strip, then take the max fit by aspect ratio
  const ar = slide.widthPx / slide.heightPx
  const mainW = size.w - SIDE_W
  const mainH = size.h - TOP_H - TIMER_H - BOTTOM_H - FILM_H
  const fitW = Math.max(80, Math.round(Math.min(mainW - 48, (mainH - 16) * ar)))
  const fitH = Math.round(fitW / ar)

  const nextSlide = ended ? null : (slides[order[pos + 1] ?? -1] ?? null)
  const notes = allNotes[order[pos]!] ?? ''

  const useShow = () => {
    keepFsRef.current = true
    onUseSlideShow?.(order[Math.min(pos, order.length - 1)] ?? startAt)
  }

  return (
    <div className="presenter">
      <div className="pv-top">
        <button
          className="pv-top-btn pv-top-exit"
          onClick={() => exitRef.current()}
          data-tip={t('panePresenterEndTip')}
        >
          ⊗ {t('panePresenterEndShow')}
        </button>
        <button
          className="pv-top-btn"
          disabled={!hasAudience}
          onClick={() => void window.slidesApi.presenterSwap()}
          data-tip={hasAudience ? t('panePresenterSwapTip') : t('panePresenterNoSecond')}
        >
          ⇄ {t('panePresenterSwap')}
        </button>
        {onUseSlideShow && (
          <button className="pv-top-btn" onClick={useShow} data-tip={t('panePresenterUseShowTip')}>
            ▤ {t('panePresenterUseShow')}
          </button>
        )}
        <div className="pv-top-spacer" />
        {!hasAudience && <span className="pv-top-hint">{t('panePresenterSingleHint')}</span>}
      </div>
      <div className="pv-body">
        <div className="pv-left">
          <div className="pv-timer-row">
            <span className="pv-timer" data-tip={t('panePresenterElapsed')}>
              {fmtElapsed(elapsed)}
            </span>
            <button
              className="pv-mini-btn"
              onClick={() => setPaused((v) => !v)}
              data-tip={paused ? t('panePresenterResume') : t('panePresenterPause')}
              aria-label={paused ? t('panePresenterResume') : t('panePresenterPause')}
            >
              {paused ? '▶' : '❚❚'}
            </button>
            <button
              className="pv-mini-btn"
              onClick={() => {
                setElapsed(0)
                setPaused(false)
              }}
              data-tip={t('panePresenterRestart')}
              aria-label={t('panePresenterRestart')}
            >
              ↻
            </button>
            <div className="pv-top-spacer" />
            <span className="pv-clock" data-tip={t('panePresenterClock')}>
              {clock}
            </span>
          </div>
          <div className="pv-stage" onClick={tool === 'none' ? next : undefined}>
            {ended ? (
              <div className="pv-end">{t('panePresenterEnded')}</div>
            ) : (
              <div
                ref={stageboxRef}
                className={`pv-stagebox${tool !== 'none' ? ` pv-tool-${tool}` : ''}`}
                style={{ width: fitW, height: fitH }}
                onPointerDown={onStagePointerDown}
                onPointerMove={onStagePointerMove}
                onPointerUp={onStagePointerUp}
                onPointerLeave={onStagePointerLeave}
              >
                <AnimatedSlideStage
                  slide={slide}
                  images={images}
                  width={fitW}
                  states={player.states}
                />
                <InkLayer strokes={strokes} laser={laser} width={fitW} height={fitH} />
                {black && <div className="pv-black" data-tip={t('panePresenterBlackOn')} />}
              </div>
            )}
          </div>
          <div className="pv-bottomrow">
            <div className="pv-tools">
              <button
                className={`pv-tool-btn${tool === 'pen' ? ' pv-tool-on' : ''}`}
                onClick={() => setTool((cur) => (cur === 'pen' ? 'none' : 'pen'))}
                data-tip={t('panePresenterPen')}
                aria-label={t('panePresenterPen')}
              >
                ✎
              </button>
              <button
                className={`pv-tool-btn${tool === 'laser' ? ' pv-tool-on' : ''}`}
                onClick={() => setTool((cur) => (cur === 'laser' ? 'none' : 'laser'))}
                data-tip={t('panePresenterLaser')}
                aria-label={t('panePresenterLaser')}
              >
                ◉
              </button>
              <button
                className="pv-tool-btn"
                disabled={strokes.length === 0}
                onClick={clearInk}
                data-tip={t('panePresenterEraseInk')}
                aria-label={t('panePresenterEraseInk')}
              >
                ⌫
              </button>
              <button
                className={`pv-tool-btn${black ? ' pv-tool-on' : ''}`}
                onClick={() => setBlack((v) => !v)}
                data-tip={t('panePresenterBlackTip')}
                aria-label={t('panePresenterBlackTip')}
              >
                ▮
              </button>
            </div>
            <div className="pv-nav">
              <button
                className="pv-round"
                disabled={!ended && pos === 0}
                onClick={prev}
                data-tip={t('panePresenterPrevTip')}
                aria-label={t('panePresenterPrevTip')}
              >
                ‹
              </button>
              <div className="pv-nav-mid">
                <div className="pv-nav-label">
                  {t('panePresenterSlideOf', { cur: pos + 1, total: order.length })}
                </div>
                <div className="pv-progress">
                  <div style={{ width: `${Math.round(((pos + 1) / order.length) * 100)}%` }} />
                </div>
              </div>
              <button
                className="pv-round"
                onClick={next}
                data-tip={t('panePresenterNextTip')}
                aria-label={t('panePresenterNextTip')}
              >
                ›
              </button>
            </div>
            <div className="pv-tools" style={{ visibility: 'hidden' }} aria-hidden />
          </div>
        </div>
        <div className="pv-side">
          <div className="pv-section-label">{t('panePresenterNextSlide')}</div>
          <div className="pv-next">
            {nextSlide ? (
              <SlideThumb slide={nextSlide} images={images} width={SIDE_W - 28} />
            ) : (
              <div className="pv-next-none">{ended ? '—' : t('panePresenterLastSlide')}</div>
            )}
          </div>
          <div className="pv-section-label">{t('panePresenterNotes')}</div>
          <div className="pv-notes" style={{ fontSize: noteSize }}>
            {notes.trim() ? notes : t('panePresenterNoNotes')}
          </div>
          <div className="pv-notes-size">
            <button
              className="pv-mini-btn"
              onClick={() => setNoteSize((s) => Math.min(28, s + 2))}
              data-tip={t('panePresenterNotesBigger')}
            >
              A⁺
            </button>
            <button
              className="pv-mini-btn"
              onClick={() => setNoteSize((s) => Math.max(11, s - 2))}
              data-tip={t('panePresenterNotesSmaller')}
            >
              A⁻
            </button>
          </div>
        </div>
      </div>
      <div className="pv-film" ref={filmRef}>
        {order.map((idx, i) => (
          <div
            key={idx}
            className={`pv-film-item${i === pos ? ' pv-film-cur' : ''}`}
            onClick={() => {
              setEnded(false)
              goTo(i, false)
            }}
            data-tip={t('panePresenterFilmSlideN', { n: i + 1 })}
          >
            <SlideThumb slide={slides[idx]!} images={images} width={150} />
          </div>
        ))}
      </div>
    </div>
  )
}
