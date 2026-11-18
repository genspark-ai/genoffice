import { createZoomWheelClassifier, notchStep } from '@genoffice/ui'

import { SHEET_ZOOM_MAX as MAX_PERCENT, SHEET_ZOOM_MIN as MIN_PERCENT } from './zoom-range'

/// Same range as the status-bar slider and +/- buttons, as a ratio.
export const SHEET_ZOOM_MIN = MIN_PERCENT / 100
export const SHEET_ZOOM_MAX = MAX_PERCENT / 100

export interface WheelZoomTarget {
  readonly getZoom: () => number | undefined
  readonly setZoom: (zoom: number) => void
  readonly isCellEditing: () => boolean
}

type WheelLike = Pick<
  WheelEvent,
  | 'deltaY'
  | 'deltaMode'
  | 'ctrlKey'
  | 'metaKey'
  | 'timeStamp'
  | 'preventDefault'
  | 'stopPropagation'
>

/// A trackpad pinch arrives as a stream of small wheel deltas; accumulate them
/// into the same 10-point step a mouse notch produces.
const PINCH_STEP_DELTA = 40

/// Runs ahead of Univer's own Ctrl+wheel zoom (which scales by pixel delta, so
/// a Windows mouse notch jumps 14 points): a notch steps 10 points here and is
/// stopped before Univer sees it. Outside the cell editor a trackpad pinch
/// still falls through to Univer's continuous zoom; while a cell is being
/// edited that path is refused by `SetZoomRatioCommand`, so the pinch is
/// accumulated into the same step — Excel keeps zooming in both cases and the
/// two input paths must agree.
export function createWheelZoomStepper(target: WheelZoomTarget): (event: WheelLike) => void {
  const classifier = createZoomWheelClassifier()
  let pinchDelta = 0

  const step = (direction: 1 | -1) => {
    const zoom = target.getZoom()
    if (zoom === undefined) return
    const next = notchStep(zoom, direction, SHEET_ZOOM_MIN, SHEET_ZOOM_MAX)
    if (next !== zoom) target.setZoom(next)
  }

  return (event) => {
    if (!event.ctrlKey && !event.metaKey) return
    const intent = classifier.feed(event, event.timeStamp)
    if (intent === 'pinch') {
      if (!target.isCellEditing()) return
      event.preventDefault()
      event.stopPropagation()
      pinchDelta += event.deltaY
      if (Math.abs(pinchDelta) < PINCH_STEP_DELTA) return
      step(pinchDelta > 0 ? -1 : 1)
      pinchDelta = 0
      return
    }
    event.preventDefault()
    event.stopPropagation()
    if (!intent) return
    step(intent === 'zoom-in' ? 1 : -1)
  }
}

export function installWheelZoomSteps(host: HTMLElement, target: WheelZoomTarget): () => void {
  const step = createWheelZoomStepper(target)
  const onWheel = (event: WheelEvent) => {
    if ((event.target as HTMLElement | null)?.closest?.('.chart-editor')) return
    step(event)
  }
  host.addEventListener('wheel', onWheel, { capture: true, passive: false })
  return () => host.removeEventListener('wheel', onWheel, { capture: true })
}
