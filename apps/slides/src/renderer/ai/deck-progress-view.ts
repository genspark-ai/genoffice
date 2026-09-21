import type { PageProgressItem } from './slides-skill'
import type { TFunc } from '../i18n/locale'

type StageStatus = 'running' | 'done' | 'error'

/** Generation progress snapshot in the chat stream (same card updated in real time) */
export interface DeckProgressSnapshot {
  style?: { label: string; status: StageStatus; summary: string }
  plan?: { label: string; done: number; total: number; status: StageStatus; summary: string }
  images?: { label: string; done: number; total: number; status: StageStatus; summary: string }
  pages?: {
    label: string
    done: number
    total: number
    status: StageStatus
    summary: string
    items: PageProgressItem[]
  }
  finalTotal?: number // Total page count from the done event
  isDone?: boolean
  doneSummary?: string
  doneOutcome?: 'failed' | 'cancelled'
}

export type StepStatus = 'done' | 'error' | 'running'

export interface DeckProgressView {
  head: { text: string; tone: 'running' | 'done' | 'error' }
  steps: Array<{ key: string; label: string; stepStatus: StepStatus }>
}

/**
 * Head line + step list of the progress card. Terminal outcome comes from the done event itself:
 * a run that landed zero pages is a failure even though it is "done", and a stopped run keeps its
 * own wording instead of the success line.
 */
export function deriveDeckProgressView(progress: DeckProgressSnapshot, t: TFunc): DeckProgressView {
  const { style, plan, images, pages, isDone, finalTotal, doneSummary, doneOutcome } = progress
  const failed = doneOutcome === 'failed'

  // In progress/failed read summary; success freezes to the label
  const stepView = (
    status: StageStatus,
    label: string,
    summary: string,
  ): { label: string; stepStatus: StepStatus } => ({
    label: status === 'done' ? label : summary,
    stepStatus: status === 'done' ? 'done' : status === 'error' ? 'error' : 'running',
  })

  const steps: DeckProgressView['steps'] = []
  if (style) steps.push({ key: 'style', ...stepView(style.status, style.label, style.summary) })
  if (plan) steps.push({ key: 'plan', ...stepView(plan.status, plan.label, plan.summary) })
  if (images) {
    steps.push({ key: 'images', ...stepView(images.status, images.label, images.summary) })
  }
  if (pages) {
    const allDone = isDone || pages.status === 'done'
    const hasError = failed || pages.items.some((p) => p.status === 'error')
    const settledLabel = `${pages.label}${pages.total > 0 ? t('aiPagesSuffix', { n: pages.total }) : ''}`
    steps.push({
      key: 'pages',
      label:
        failed && doneSummary ? doneSummary : allDone ? settledLabel : pages.summary || pages.label,
      stepStatus: allDone ? (hasError ? 'error' : 'done') : 'running',
    })
  }

  const hasStepError = steps.some((s) => s.stepStatus === 'error')
  let head: DeckProgressView['head']
  if (failed) {
    head = { text: t('aiProgressFailed'), tone: 'error' }
  } else if (doneOutcome === 'cancelled' && doneSummary) {
    head = { text: doneSummary, tone: 'done' }
  } else if (isDone && finalTotal != null) {
    head = { text: t('aiProgressDone', { n: finalTotal }), tone: 'done' }
  } else if (hasStepError) {
    head = { text: t('aiProgressFailed'), tone: 'error' }
  } else {
    head = { text: t('aiProgressTitle'), tone: 'running' }
  }
  return { head, steps }
}
