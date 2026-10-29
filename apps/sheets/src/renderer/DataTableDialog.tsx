import { useState } from 'react'

import type { DataTableInput, DataTableOutcome } from './data-table'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Excel's Data Table: the table region (corner formula + substitution value
/// headers already on the sheet) plus the row and/or column input cell. The
/// run is one undo batch of journaled edits, so ⌘Z removes the whole table;
/// results land as plain values and do not refresh on recalculation (the
/// dialog note says so — Excel keeps them live via {=TABLE()}).
export function DataTableDialog({
  initialRange,
  onApply,
  onClose,
}: {
  /// Prefill for the range field: the multi-cell selection at open time.
  readonly initialRange: string
  /// Rejects with an Error whose message is user-facing; `onProgress` on the
  /// input mirrors the engine's per-cell progress.
  readonly onApply: (input: DataTableInput) => Promise<DataTableOutcome>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [range, setRange] = useState(initialRange)
  const [rowInput, setRowInput] = useState('')
  const [columnInput, setColumnInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState<{
    readonly done: number
    readonly total: number
  } | null>(null)
  const [outcome, setOutcome] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  const CELL_REF = /^[A-Za-z]{1,3}\d{1,7}$/
  const RANGE_REF = /^[A-Za-z]{1,3}\d{1,7}(:[A-Za-z]{1,3}\d{1,7})?$/
  const rowFilled = rowInput.trim() !== ''
  const columnFilled = columnInput.trim() !== ''
  const hasInput = rowFilled || columnFilled
  const valid =
    RANGE_REF.test(range.trim()) &&
    (!rowFilled || CELL_REF.test(rowInput.trim())) &&
    (!columnFilled || CELL_REF.test(columnInput.trim())) &&
    hasInput

  const apply = async (): Promise<void> => {
    setBusy(true)
    setError(null)
    setOutcome(null)
    setProgress(null)
    try {
      // Progress re-renders are throttled to whole percent steps: a 10,000
      // cell run must not cost 10,000 renders.
      const result = await onApply({
        tableRange: range.trim().toUpperCase(),
        rowInputCell: rowFilled ? rowInput.trim().toUpperCase() : null,
        columnInputCell: columnFilled ? columnInput.trim().toUpperCase() : null,
        onProgress: (done, total) => {
          const percent = Math.floor((done / total) * 100)
          setProgress((previous) =>
            previous && Math.floor((previous.done / previous.total) * 100) === percent
              ? previous
              : { done, total },
          )
        },
      })
      setOutcome(t('dlgDataTableDone', { count: result.cells }))
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure))
    } finally {
      setBusy(false)
    }
  }

  const modal = useModalDialog(() => {
    if (!busy) onClose()
  })
  return (
    <div className="dialog-backdrop" onClick={busy ? undefined : onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgDataTableTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgDataTableTitle')}</header>
        <section className="dialog-body">
          <div className="dialog-grid">
            <label>
              {t('dlgDataTableRange')}
              <input
                autoFocus
                value={range}
                placeholder="B4:D8"
                disabled={busy}
                onChange={(event) => setRange(event.target.value)}
              />
            </label>
            <label>
              {t('dlgDataTableRowInput')}
              <input
                value={rowInput}
                placeholder="B1"
                disabled={busy}
                onChange={(event) => setRowInput(event.target.value)}
              />
            </label>
            <label>
              {t('dlgDataTableColumnInput')}
              <input
                value={columnInput}
                placeholder="A4"
                disabled={busy}
                onChange={(event) => setColumnInput(event.target.value)}
              />
            </label>
          </div>
          <p className="dialog-note">{t('dlgDataTableNote')}</p>
          {valid ? null : <p className="dialog-note">{t('dlgDataTableNeedInput')}</p>}
          {busy && progress && (
            <p className="dialog-note">
              {t('dlgDataTableRunning', { done: progress.done, total: progress.total })}
            </p>
          )}
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
          <button className="primary-action" disabled={!valid || busy} onClick={() => void apply()}>
            {busy ? t('dlgDataTableRunningShort') : t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
