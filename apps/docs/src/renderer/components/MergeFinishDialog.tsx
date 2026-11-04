import { useState } from 'react'
import type { MergeDataSource } from '../../shared/ipc'
import type { MergeOutputMode } from '../../shared/ipc'
import type { MergeRecord } from '@genoffice/docx-engine'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * The "Finish & Merge" dialog: the record range to merge, then the merge
 * itself. The data source and the fields come from the tab; this only asks
 * which records the letters are written for. genoffice#2097.
 */

export interface MergeFinishDialogProps {
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** the records the merge walks: the ones the recipient list kept */
  records: readonly MergeRecord[]
  /** the record the user is sitting on, for the "current record" range */
  currentIndex: number
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
  onClose: () => void
}

export function MergeFinishDialog({
  source,
  records,
  currentIndex,
  onMerge,
  onClose,
}: MergeFinishDialogProps) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [mode, setMode] = useState<MergeOutputMode>('single')
  const [namePattern, setNamePattern] = useState('')
  const [range, setRange] = useState<'all' | 'current' | 'fromTo'>('all')
  const [from, setFrom] = useState(1)
  const [to, setTo] = useState(Math.max(records.length, 1))
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState('')

  // the record window the merge walks: the "from X to Y" bounds are clamped to
  // the list so a typo cannot push the merge past the end
  const rangeWindow = (): { start: number; end: number } => {
    const last = records.length - 1
    if (range === 'current') return { start: currentIndex, end: currentIndex }
    if (range === 'fromTo') {
      const a = Math.min(Math.max(from, 1), records.length)
      const b = Math.min(Math.max(to, a), records.length)
      return { start: a - 1, end: b - 1 }
    }
    return { start: 0, end: last }
  }

  const runMerge = async () => {
    if (!source || records.length === 0) return
    setBusy(true)
    setError('')
    setDone('')
    try {
      const result = await onMerge({
        mode,
        records,
        fields: source.fields,
        namePattern,
        range: rangeWindow(),
      })
      if (result.canceled) {
        setBusy(false)
        return
      }
      if (!result.ok) {
        setError(result.error ?? t('ribbonMailingsMergeFailed'))
        setBusy(false)
        return
      }
      setDone(t('ribbonMailingsMerged', { n: result.count ?? records.length }))
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal modal-mailmerge">
        <h2>{t('ribbonMailingsFinish')}</h2>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsSource')}</div>
          <div className="modal-row">
            <span className="mailmerge-source-name" title={source?.path}>
              {source ? source.name : t('ribbonMailingsNoSource')}
            </span>
          </div>
          {source?.sheet && (
            <div className="mailmerge-sheet-name">
              {t('ribbonMailingsSheetLabel', { sheet: source.sheet })}
            </div>
          )}
        </div>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsRange')}</div>
          <div className="modal-row mailmerge-option-row">
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-range"
                checked={range === 'all'}
                onChange={() => setRange('all')}
              />
              {t('ribbonMailingsRangeAll', { n: records.length })}
            </label>
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-range"
                checked={range === 'current'}
                onChange={() => setRange('current')}
              />
              {t('ribbonMailingsRangeCurrent', { n: currentIndex + 1 })}
            </label>
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-range"
                checked={range === 'fromTo'}
                onChange={() => setRange('fromTo')}
              />
              {t('ribbonMailingsRangeFromTo')}
              <input
                className="mailmerge-range-num"
                type="number"
                min={1}
                max={records.length}
                value={from}
                onChange={(e) => setFrom(Number(e.target.value))}
              />
              {t('ribbonMailingsRangeTo')}
              <input
                className="mailmerge-range-num"
                type="number"
                min={1}
                max={records.length}
                value={to}
                onChange={(e) => setTo(Number(e.target.value))}
              />
            </label>
          </div>
        </div>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsOutput')}</div>
          <div className="modal-row mailmerge-option-row">
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-output"
                checked={mode === 'single'}
                onChange={() => setMode('single')}
              />
              {t('ribbonMailingsOutputSingle')}
            </label>
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-output"
                checked={mode === 'perRecord'}
                onChange={() => setMode('perRecord')}
              />
              {t('ribbonMailingsOutputPerRecord')}
            </label>
          </div>
          {mode === 'perRecord' && (
            <div className="modal-row mailmerge-option-row">
              <span className="mailmerge-name-label">{t('ribbonMailingsNamePattern')}</span>
              <input
                type="text"
                className="mailmerge-name-input"
                placeholder={t('ribbonMailingsNamePatternPh')}
                value={namePattern}
                onChange={(e) => setNamePattern(e.target.value)}
              />
            </div>
          )}
        </div>

        {error && <div className="modal-error">{error}</div>}
        {done && <div className="mailmerge-done">{done}</div>}

        <div className="modal-actions">
          <button
            className="btn-primary"
            disabled={busy || !source || records.length === 0}
            onClick={() => void runMerge()}
          >
            {t('ribbonMailingsFinish')}
          </button>
          <button onClick={onClose}>{t('ribbonMailingsClose')}</button>
        </div>
      </div>
    </div>
  )
}
