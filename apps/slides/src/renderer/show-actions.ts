import { mergeSlideUpdate } from '../shared/slide-update'
/**
 * Slide-show tab actions extracted from App.tsx: starting shows,
 * presenter view, custom shows, rehearsal timings, and hiding slides.
 * Functions take the ActionCtx built fresh per call.
 */
import type { ActionCtx } from './action-context'
import { recordedSlides } from './narration/encode'
import type { NarrationTake } from './narration/use-recording'
import type { CustomShow } from './slideshow-utils'
import { t } from './i18n/locale'

/** Instant black curtain under the upcoming show (body::after overlay): painted on the
 *  very next frame after the click, it hides the React mount + window-snap latency.
 *  SlideShowView lifts it once the show is revealed (and on unmount, for early exits). */
export function dropShowCurtain(): void {
  document.body.classList.add('show-curtain')
}
export function liftShowCurtain(): void {
  document.body.classList.remove('show-curtain')
}

export function startSlideShow(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  // From start: jump to the first unhidden slide (if all hidden, still start at slide 1)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitSlideShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setSlideShow(null)
  ctx.setCurrent(lastIndex)
}

/** Update the custom show list overwrite-style and persist */
export function updateCustomShows(ctx: ActionCtx, shows: CustomShow[]): void {
  ctx.setCustomShows(shows)
  if (ctx.path) {
    try {
      localStorage.setItem(`ai-slides-custom-shows:${ctx.path}`, JSON.stringify(shows))
    } catch {
      // Degrade to memory-only when localStorage is full/unavailable
    }
  }
}

/** Play a custom show: only its included slides (filtering out-of-range indexes stale after deletions) */
export function playCustomShow(ctx: ActionCtx, show: CustomShow): void {
  if (ctx.slideShow || ctx.presenter) return
  const order = show.slideIndices.filter((i) => i >= 0 && i < ctx.slides.length)
  if (order.length === 0) {
    ctx.setStatus(t('appStatusCustomShowEmpty'))
    return
  }
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  ctx.setCustomShowDlgOpen(false)
  ctx.setSlideShow({ startAt: order[0]!, customOrder: order })
}

/** Rehearsal timing: show from the start and record each slide's dwell time */
export function startRehearseShow(ctx: ActionCtx): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: Math.max(0, first), rehearse: true })
}

/** Rehearsal ended: stash per-slide seconds; after exiting the show, prompt "save?" */
export function onRehearseDone(ctx: ActionCtx, perPageSec: number[]): void {
  if (perPageSec.some((s) => s > 0)) ctx.setPendingRehearse(perPageSec)
}

/** Save rehearsal timings: write each slide's dwell seconds as auto-advance times (<p:transition advTm>, milliseconds) */
export async function saveRehearseTimings(ctx: ActionCtx): Promise<void> {
  if (!ctx.pendingRehearse) return
  const times = ctx.pendingRehearse
    .map((sec, i) => ({ slideIndex: i, ms: sec * 1000 }))
    .filter((t) => t.ms > 0)
  ctx.setPendingRehearse(null)
  const ok = await window.slidesApi.setAdvanceTimes({ times })
  if (ok) {
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusRehearseSaved', { count: times.length }))
  }
}

export function startRecordShow(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: fromStart ? Math.max(0, first) : ctx.current, record: true })
}

/** Recording ended: once the microphone stops, prompt "keep the recording?" */
export function onRecordDone(ctx: ActionCtx, take: Promise<NarrationTake | null>): void {
  void take.then((recorded) => {
    if (recorded && recorded.endMs > 0) ctx.setPendingRecording(recorded)
  })
}

/** Keep the recording: each recorded slide gets its narration and its time as advance time */
export async function saveRecording(ctx: ActionCtx): Promise<void> {
  const take = ctx.pendingRecording
  if (!take) return
  ctx.setPendingRecording(null)
  try {
    const slides = await recordedSlides(take.audio, take.marks, take.endMs)
    const all = slides.length > 0 ? await window.slidesApi.saveRecording({ slides }) : null
    if (!all) throw new Error('the recording was not saved')
    ctx.setSlides(all)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusRecordSaved', { count: slides.length }))
  } catch {
    ctx.setStatus(t('appStatusRecordFailed'))
  }
}

export async function clearNarration(ctx: ActionCtx, all: boolean): Promise<void> {
  const indexes = all ? ctx.slides.map((_, i) => i) : [ctx.current]
  const updated = await window.slidesApi.clearNarration(indexes)
  if (!updated) return
  ctx.setSlides(updated)
  ctx.setDirty(true)
  ctx.setStatus(t('appStatusNarrationCleared'))
}

/** Presenter view (single-window version, entry aligned with the show) */
export function startPresenterView(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setPresenter({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitPresenterView(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
}

/** "Switch to normal show" inside presenter view: seamlessly turns into a full-screen show in this window */
export function switchPresenterToShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
  ctx.setSlideShow({ startAt: lastIndex })
}

/** Hide/unhide a thumbnail selection as one undo step; the selection stays put */
export async function setSlidesHidden(
  ctx: ActionCtx,
  indexes: number[],
  hidden: boolean,
): Promise<void> {
  const sel = [...new Set(indexes)].filter((i) => ctx.slides[i])
  if (!sel.length) return
  const r = await window.slidesApi.setSlidesHidden({ slideIndexes: sel, hidden })
  if (!r) return
  ctx.setSlides((current) => mergeSlideUpdate(current, r))
  ctx.setDirty(true)
  if (sel.length === 1) {
    const page = sel[0]! + 1
    ctx.setStatus(
      hidden ? t('appStatusSlideHidden', { page }) : t('appStatusSlideUnhidden', { page }),
    )
  } else {
    const count = sel.length
    ctx.setStatus(
      hidden ? t('appStatusSlidesHidden', { count }) : t('appStatusSlidesUnhidden', { count }),
    )
  }
}
