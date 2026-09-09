import { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { ClarifyQuestion } from '../ai/tools'

/** Questionnaire card shown in the composer slot while ask_clarification waits (same card as slides). */
export function ClarifyCard({
  questions,
  onSubmit,
  onSkip,
}: {
  questions: ClarifyQuestion[]
  onSubmit: (answers: string, qa: Array<{ q: string; a: string }>) => void
  onSkip: () => void
}) {
  const { t } = useI18n()
  const decideAnswer = t('aiClarifyDecideAnswer')
  // Per question: set of selected options + the "Other" free text
  const [picked, setPicked] = useState<Record<string, Set<string>>>({})
  const [other, setOther] = useState<Record<string, string>>({})
  // Latest selections for the delayed auto-submit (the timer closure would otherwise
  // read the pre-click state and drop the final pick)
  const pickedRef = useRef(picked)
  pickedRef.current = picked
  const otherRef = useRef(other)
  otherRef.current = other
  // Pager view-state (one question at a time): back-nav limited to visited range,
  // single-select auto-advances after a beat, typing cancels the pending advance
  const [qIdx, setQIdx] = useState(0)
  const [furthest, setFurthest] = useState(0)
  const [slideDir, setSlideDir] = useState<'next' | 'prev'>('next')
  const advanceTimerRef = useRef<number | null>(null)

  const cancelAdvance = () => {
    if (advanceTimerRef.current) {
      window.clearTimeout(advanceTimerRef.current)
      advanceTimerRef.current = null
    }
  }

  const goTo = (i: number) => {
    cancelAdvance()
    const clamped = Math.max(0, Math.min(i, questions.length - 1))
    setSlideDir(clamped >= qIdx ? 'next' : 'prev')
    setQIdx(clamped)
    setFurthest((f) => Math.max(f, clamped))
  }

  // Single-select: picking advances; picking on the LAST question submits after the
  // same beat (typing in "Other" or navigating cancels the pending action)
  const scheduleAdvance = (from: number) => {
    cancelAdvance()
    const last = from >= questions.length - 1
    advanceTimerRef.current = window.setTimeout(() => {
      advanceTimerRef.current = null
      if (last) submit()
      else goTo(from + 1)
    }, 250)
  }

  useEffect(() => cancelAdvance, [])

  const toggle = (qid: string, opt: string, multi?: boolean, exclusive?: boolean) => {
    setPicked((prev) => {
      const cur = new Set(prev[qid] ?? [])
      if (multi) {
        if (exclusive) {
          cur.clear()
          cur.add(opt)
        } else {
          cur.delete(decideAnswer)
          if (cur.has(opt)) cur.delete(opt)
          else cur.add(opt)
        }
      } else {
        cur.clear()
        cur.add(opt)
      }
      return { ...prev, [qid]: cur }
    })
  }

  const submit = () => {
    const qa = questions.map((q) => {
      const chosen = [...(pickedRef.current[q.id] ?? [])]
      const ot = (otherRef.current[q.id] ?? '').trim()
      if (ot) chosen.push(ot)
      const ans = chosen.length ? chosen.join('、') : t('aiClarifyDecideAnswer')
      return { q: q.label, a: ans }
    })
    onSubmit(qa.map(({ q, a }) => `${q}: ${a}`).join('\n'), qa)
  }

  const q = questions[qIdx]
  const isLast = qIdx === questions.length - 1
  const selCount = picked[q.id]?.size ?? 0
  const hasAnswer = selCount > 0 || !!other[q.id]?.trim()

  const renderOpt = (opt: string, label: string) => (
    <button
      key={opt}
      className={`ai-clarify-opt${q.multi ? ' multi' : ''}${picked[q.id]?.has(opt) ? ' ai-clarify-opt-on' : ''}`}
      role={q.multi ? 'checkbox' : 'radio'}
      aria-checked={picked[q.id]?.has(opt) ?? false}
      onClick={() => {
        const isDecide = opt === decideAnswer
        toggle(q.id, opt, q.multi, isDecide)
        if (isDecide) {
          setOther((prev) => ({ ...prev, [q.id]: '' }))
          scheduleAdvance(qIdx)
        } else if (q.multi) {
          cancelAdvance()
        } else {
          scheduleAdvance(qIdx)
        }
      }}
    >
      <span className="ai-clarify-opt-box" aria-hidden />
      <span className="ai-clarify-opt-label">{label}</span>
      {!q.multi && (
        <span className="ai-clarify-opt-arrow" aria-hidden>
          ›
        </span>
      )}
    </button>
  )

  return (
    <div className="ai-clarify-card">
      <div className="ai-clarify-head">
        <span className="ai-clarify-head-label">{t('aiClarifyTitle')}</span>
        <span className="ai-clarify-head-progress" aria-live="polite">
          {`${qIdx + 1} / ${questions.length}`}
        </span>
        <span className="ai-clarify-head-arrows">
          <button
            type="button"
            className="ai-clarify-head-arrow"
            disabled={qIdx === 0}
            onClick={() => goTo(qIdx - 1)}
            aria-label="‹"
          >
            ‹
          </button>
          <button
            type="button"
            className="ai-clarify-head-arrow"
            disabled={qIdx >= furthest || !hasAnswer}
            onClick={() => goTo(qIdx + 1)}
            aria-label="›"
          >
            ›
          </button>
        </span>
      </div>
      <div key={q.id} className={`ai-clarify-q slide-${slideDir}`}>
        <div className="ai-clarify-question-head">
          <div className="ai-clarify-label">{q.label}</div>
          {q.multi && <span className="ai-clarify-multi-badge">{t('aiClarifyMulti')}</span>}
        </div>
        {q.description && <div className="ai-clarify-desc">{q.description}</div>}
        <div className="ai-clarify-opts">
          {q.options.map((opt) => renderOpt(opt, opt))}
          {renderOpt(decideAnswer, t('aiClarifyDecide'))}
        </div>
        <input
          className="ai-clarify-other"
          placeholder={t('aiClarifyOther')}
          value={other[q.id] ?? ''}
          onChange={(e) => {
            cancelAdvance()
            const value = e.target.value
            setOther((p) => ({ ...p, [q.id]: value }))
            if (value.trim()) {
              setPicked((prev) => {
                const cur = new Set(prev[q.id] ?? [])
                cur.delete(decideAnswer)
                return { ...prev, [q.id]: cur }
              })
            }
          }}
        />
      </div>
      <div className={`ai-clarify-actions${q.multi ? ' multi' : ''}`}>
        {q.multi && selCount > 0 && (
          <span className="ai-clarify-count">{t('aiClarifySelected', { n: selCount })}</span>
        )}
        <span className="ai-clarify-actions-btns">
          <button className="ai-clarify-skip" onClick={onSkip}>
            {t('aiClarifySkip')}
          </button>
          {isLast ? (
            <button className="ai-clarify-submit" onClick={submit} disabled={!hasAnswer}>
              {t('aiClarifySubmit')}
            </button>
          ) : (
            /* Multi-select waits for an explicit Next; single-select advances by picking. */
            q.multi && (
              <button
                type="button"
                className="ai-clarify-next"
                onClick={() => goTo(qIdx + 1)}
                disabled={!hasAnswer}
                aria-label={t('aiClarifyNext')}
              >
                <span>{t('aiClarifyNext')}</span>
                <span className="ai-clarify-next-arrow" aria-hidden>
                  →
                </span>
              </button>
            )
          )}
        </span>
      </div>
    </div>
  )
}
