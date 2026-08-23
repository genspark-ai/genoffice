/**
 * Audience show window — the presenter view's external-screen full-screen playback end
 * (?mode=audience entry).
 *
 * When the main process creates this window it has already shared the presenter's document
 * session with this webContents, so read-only IPC pulls RenderSlide/animations/transitions and
 * sees the in-memory document. State is fully passive: the presenter broadcasts ShowSyncState
 * (absolute state) and this side seeks idempotently; forward page turns play the transition
 * effect (including morph tweens), ink/laser mirror the presenter.
 * Navigation input (clicks/→ etc.) is sent back for presenter arbitration (audienceNav); Esc ends
 * the whole show.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { RenderFill, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem, ShapeKey, ShowSyncState, TransitionKind } from '../../shared/ipc'
import { AnimatedSlideStage, useAnimPlayer } from './AnimatedSlide'
import { useI18n } from '../i18n/locale'
import { MorphStage } from './MorphStage'
import { InkLayer, type InkStroke } from './ShowInk'

const ANIMATED = ['fade', 'push', 'wipe', 'split', 'circle'] as const

/** dataUrl → HTMLImageElement (picture elements/picture fills/background images, recursing into groups and table cells) */
function useSlideImages(slides: RenderSlide[] | null): Map<string, HTMLImageElement> {
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  useEffect(() => {
    if (!slides || slides.length === 0) return
    const urls = new Set<string>()
    const addFillUrl = (fill: RenderFill | undefined) => {
      if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
    }
    const walk = (nodes: readonly RenderNode[]) => {
      for (const n of nodes) {
        if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
        if ((n.type === 'shape' || n.type === 'text') && n.fill) addFillUrl(n.fill)
        if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
        if (n.type === 'table' && Array.isArray(n.cells)) {
          for (const c of n.cells) if (c.fill) addFillUrl(c.fill)
        }
      }
    }
    for (const s of slides) {
      addFillUrl(s.background)
      walk(s.nodes)
    }
    let cancelled = false
    const loaded = new Map<string, HTMLImageElement>()
    let pending = 0
    const done = () => {
      if (--pending === 0 && !cancelled && loaded.size > 0) {
        setImages((prev) => {
          const m = new Map(prev)
          for (const [k, v] of loaded) m.set(k, v)
          return m
        })
      }
    }
    urls.forEach((u) => {
      pending++
      const img = new Image()
      img.onload = () => {
        loaded.set(u, img)
        done()
      }
      img.onerror = done
      img.src = u
    })
    return () => {
      cancelled = true
    }
  }, [slides])
  return images
}

export function AudienceView() {
  const { t } = useI18n()
  const [slides, setSlides] = useState<RenderSlide[] | null>(null)
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  const images = useSlideImages(slides)
  const [sync, setSync] = useState<ShowSyncState | null>(null)
  const [allAnims, setAllAnims] = useState<AnimationItem[][] | null>(null)
  const transRef = useRef<TransitionKind[]>([])
  const keysRef = useRef<ShapeKey[][]>([])
  const [anim, setAnim] = useState<{ kind: TransitionKind; nonce: number }>({
    kind: 'none',
    nonce: 0,
  })
  const [morph, setMorph] = useState<{ fromIdx: number; toIdx: number; nonce: number } | null>(null)
  const [strokes, setStrokes] = useState<InkStroke[]>([])
  const [laser, setLaser] = useState<{ x: number; y: number } | null>(null)

  // Document load: session already shared; retry covers the occasional timing where the window beats the sharing
  useEffect(() => {
    let stop = false
    const tryLoad = (attempt: number) => {
      void window.slidesApi.getRenderSlides().then((r) => {
        if (stop) return
        if (r && r.length > 0) setSlides(r)
        else if (attempt < 20) window.setTimeout(() => tryLoad(attempt + 1), 250)
      })
    }
    tryLoad(0)
    return () => {
      stop = true
    }
  }, [])

  // Mid-show edits from the presenter window arrive as shared-session broadcasts.
  // The sync cursor doesn't change on a content edit, so clear it: once the
  // reloaded animation lists land, the seek effect re-aligns the player against
  // the fresh deck instead of keeping stale animation state.
  useEffect(
    () =>
      window.slidesApi.onDeckChanged?.(({ slides: all }) => {
        if (all.length === 0) return
        cursorRef.current = ''
        setSlides(all)
      }),
    [],
  )

  useEffect(() => {
    if (!slides) return
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
    return () => {
      cancelled = true
    }
  }, [slides])

  useEffect(() => {
    const offSync = window.slidesApi.onShowSync(setSync)
    const offInk = window.slidesApi.onShowInk((ev) => {
      if (ev.type === 'clear') {
        setStrokes([])
        setLaser(null)
      } else if (ev.type === 'laser') {
        setLaser(ev.x < 0 ? null : { x: ev.x, y: ev.y })
      } else if (ev.type === 'stroke-start') {
        setStrokes((ss) => [...ss, { color: ev.color, points: [ev.x, ev.y] }])
      } else if (ev.type === 'stroke-move') {
        setStrokes((ss) => {
          const last = ss[ss.length - 1]
          if (!last) return ss
          return [...ss.slice(0, -1), { ...last, points: [...last.points, ev.x, ev.y] }]
        })
      }
    })
    // When mounted after the presenter's first broadcast, snapshot the current state
    void window.slidesApi.audienceReady().then((s) => {
      if (s) setSync((prev) => prev ?? s)
    })
    return () => {
      offSync()
      offInk()
    }
  }, [])

  useEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight })
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  // Navigation input sent back to the presenter (this side doesn't change state); Esc ends the whole show
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        window.slidesApi.audienceNav('exit')
      } else if (
        e.key === 'ArrowRight' ||
        e.key === 'ArrowDown' ||
        e.key === ' ' ||
        e.key === 'Enter' ||
        e.key === 'PageDown'
      ) {
        e.preventDefault()
        window.slidesApi.audienceNav('next')
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault()
        window.slidesApi.audienceNav('prev')
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [])

  const slide = sync && slides ? slides[sync.idx] : null
  const player = useAnimPlayer(slide?.heightPx ?? 540, slide?.widthPx ?? 960)

  // Apply synced state: animation cursor seek alignment; play the transition effect on page turns (forward only)
  const shownRef = useRef(-1)
  const cursorRef = useRef('')
  useEffect(() => {
    if (!sync || !slides || !allAnims) return
    const cursor = `${sync.idx}:${sync.played}:${sync.playing}`
    if (cursor !== cursorRef.current) {
      cursorRef.current = cursor
      player.seek(allAnims[sync.idx] ?? [], sync.played, sync.playing)
    }
    if (shownRef.current !== sync.idx) {
      const from = shownRef.current
      shownRef.current = sync.idx
      let kind: TransitionKind = 'none'
      if (sync.fresh && from >= 0) {
        kind = transRef.current[sync.idx] ?? 'none'
        if (kind === 'random') kind = ANIMATED[Math.floor(Math.random() * ANIMATED.length)]!
      }
      if (kind === 'morph' && from >= 0 && from !== sync.idx) {
        setMorph((m) => ({ fromIdx: from, toIdx: sync.idx, nonce: (m?.nonce ?? 0) + 1 }))
        setAnim((a) => ({ kind: 'none', nonce: a.nonce + 1 }))
      } else {
        if (kind === 'morph') kind = 'fade'
        setMorph(null)
        setAnim((a) => ({ kind, nonce: a.nonce + 1 }))
      }
      setStrokes([])
      setLaser(null)
    }
    // player.seek has a stable reference
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sync, slides, allAnims])

  if (!slide || !sync) {
    return <div className="slideshow" />
  }

  const fitW = Math.round(Math.min(size.w, (size.h * slide.widthPx) / slide.heightPx))
  const fitH = Math.round((fitW * slide.heightPx) / slide.widthPx)

  return (
    <div
      className="slideshow"
      onClick={() => window.slidesApi.audienceNav('next')}
      onContextMenu={(e) => {
        e.preventDefault()
        window.slidesApi.audienceNav('prev')
      }}
    >
      {sync.ended ? (
        <div className="ss-end">{t('paneShowEnded')}</div>
      ) : (
        <div className="ss-stagebox" style={{ width: fitW, height: fitH }}>
          {morph && slides![morph.fromIdx] && slides![morph.toIdx] ? (
            <div key={`morph-${morph.nonce}`} className="ss-frame">
              <MorphStage
                from={slides![morph.fromIdx]!}
                to={slides![morph.toIdx]!}
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
              <AnimatedSlideStage
                slide={slide}
                images={images}
                width={fitW}
                states={player.states}
              />
            </div>
          )}
          <InkLayer strokes={strokes} laser={laser} width={fitW} height={fitH} />
        </div>
      )}
      {sync.black && !sync.ended && <div className="ss-black" />}
    </div>
  )
}
