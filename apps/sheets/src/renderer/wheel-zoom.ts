import { createZoomWheelClassifier, notchStep } from '@genoffice/ui'

/// Same range as the status-bar slider and +/- buttons.
export const SHEET_ZOOM_MIN = 0.5
export const SHEET_ZOOM_MAX = 4

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

/// Runs ahead of Univer's own Ctrl+wheel zoom (which scales by pixel delta, so
/// a Windows mouse notch jumps 14 points): a notch steps 10 points here and is
/// stopped before Univer sees it; a trackpad pinch falls through to Univer's
/// continuous zoom untouched.
export function createWheelZoomStepper(target: WheelZoomTarget): (event: WheelLike) => void {
  const classifier = createZoomWheelClassifier()
  return (event) => {
    if (!event.ctrlKey && !event.metaKey) return
    if (target.isCellEditing()) return
    const intent = classifier.feed(event, event.timeStamp)
    if (intent === 'pinch') return
    event.preventDefault()
    event.stopPropagation()
    if (!intent) return
    const zoom = target.getZoom()
    if (zoom === undefined) return
    const next = notchStep(zoom, intent === 'zoom-in' ? 1 : -1, SHEET_ZOOM_MIN, SHEET_ZOOM_MAX)
    if (next !== zoom) target.setZoom(next)
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
