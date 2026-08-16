import { useState } from 'react'

import type { GoalSeekResult } from './goal-seek'
import { useI18n } from './i18n/locale'

/// Excel's Goal Seek: three fields, then an async solve whose guesses are
/// journaled edits (undo restores everything). The dialog stays open to show
/// the outcome; Cancel mid-solve is not offered (runs are a few seconds).
export function GoalSeekDialog({
  initialSetCell,
  onSolve,
  onClose,
}: {
  readonly initialSetCell: string
  /// Rejects with an Error whose message is user-facing.
  readonly onSolve: (setCell: string, toValue: number, byCell: string) => Promise<GoalSeekResult>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [setCell, setSetCell] = useState(initialSetCell)
  const [toValue, setToValue] = useState('')
  const [byCell, setByCell] = useState('')
  const [busy, setBusy] = useState(false)
  const [outcome, setOutcome] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  const CELL_REF = /^[A-Za-z]{1,3}\d{1,7}$/
  const valid =
    CELL_REF.test(setCell.trim()) && CELL_REF.test(byCell.trim()) && toValue.trim() !== ''

  const solve = async (): Promise<void> => {
    const target = Number(toValue)
    if (!Number.isFinite(target)) {
      setError(t('dlgGoalSeekNeedNumber'))
      return
    }
    setBusy(true)
    setError(null)
    setOutcome(null)
    try {
      const result = await onSolve(setCell.trim(), target, byCell.trim())
      // A failed run can end on a guess whose formula errored (NaN reached).
      const round = (value: number): string =>
        Number.isFinite(value) ? String(Number(value.toFixed(6))) : '—'
      setOutcome(
        t(result.found ? 'dlgGoalSeekFound' : 'dlgGoalSeekNotFound', {
          value: round(result.solution),
          reached: round(result.reached),
        }),
      )
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="dialog-backdrop" onClick={busy ? undefined : onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgGoalSeekTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgGoalSeekTitle')}</header>
        <section className="dialog-body">
          <div className="dialog-grid">
            <label>
              {t('dlgGoalSeekSetCell')}
              <input
                autoFocus
                value={setCell}
                placeholder="B5"
                disabled={busy}
                onChange={(event) => setSetCell(event.target.value)}
              />
            </label>
            <label>
              {t('dlgGoalSeekToValue')}
              <input
                value={toValue}
                placeholder="0"
                disabled={busy}
                onChange={(event) => setToValue(event.target.value)}
              />
            </label>
            <label>
              {t('dlgGoalSeekByCell')}
              <input
                value={byCell}
                placeholder="B2"
                disabled={busy}
                onChange={(event) => setByCell(event.target.value)}
              />
            </label>
          </div>
          <p className="dialog-note">{t('dlgGoalSeekNote')}</p>
          {outcome && <p className="dialog-note">{outcome}</p>}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" disabled={busy} onClick={onClose}>
            {t('dlgClose')}
          </button>
          <button className="primary-action" disabled={!valid || busy} onClick={() => void solve()}>
            {busy ? t('dlgGoalSeekSolving') : t('dlgGoalSeekSolve')}
          </button>
        </div>
      </div>
    </div>
  )
}
