import { useCallback, useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { MergeDataSource, MergeOutputMode } from '../../shared/ipc'
import type { MergeRecord } from '@genoffice/docx-engine'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import {
  AVERY_5160,
  clampGridCount,
  customGridLayout,
  ENVELOPE_PRESETS,
  envelopeLayout,
  unsupportedLabelBlocks,
  type EnvelopePresetId,
  type MergePageLayout,
} from '../mail-merge/layout'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * The "Finish & Merge" dialog: the record range, the output shape (plain
 * letters or multi-up labels / envelopes), then the merge itself.
 *
 * The data source and the fields come from the tab; this only asks which
 * records the letters are written for and how they are laid out. genoffice#2097,
 * genoffice#2374.
 */

export interface MergeFinishDialogProps {
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** the records the merge walks: the ones the recipient list kept */
  records: readonly MergeRecord[]
  /** the record the user is sitting on, for the "current record" range */
  currentIndex: number
  /** the active editor, used to peek at label-layout concerns */
  editor: Editor
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
  onClose: () => void
}

/** the three layout families the dialog offers */
type LayoutFamily = 'plain' | 'label' | 'envelope'

export function MergeFinishDialog({
  source,
  records,
  currentIndex,
  editor,
  onMerge,
  onClose,
}: MergeFinishDialogProps) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)

  // --- plain-letter state (mode + file-name pattern) ----------
  const [plainMode, setPlainMode] = useState<'single' | 'perRecord'>('single')
  const [namePattern, setNamePattern] = useState('')

  // --- layout family + its parameters -------------------------
  const [layoutFamily, setLayoutFamily] = useState<LayoutFamily>('plain')
  const [labelStock, setLabelStock] = useState<'avery5160' | 'custom'>('avery5160')
  const [paper, setPaper] = useState<'letter' | 'a4'>('letter')
  const [gridColumns, setGridColumns] = useState(3)
  const [gridRows, setGridRows] = useState(10)
  const [envelopeSize, setEnvelopeSize] = useState<EnvelopePresetId>('size10')

  // --- record range -------------------------------------------
  const [range, setRange] = useState<'all' | 'current' | 'fromTo'>('all')
  const [from, setFrom] = useState(1)
  const [to, setTo] = useState(Math.max(records.length, 1))

  // --- dialog state -------------------------------------------
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState('')

  // --- label-layout concern reading the template --------------
  const droppedLabelBlocks = useMemo(
    () => (layoutFamily === 'label' ? unsupportedLabelBlocks(editor.getJSON()) : 0),
    [editor, layoutFamily],
  )

  // --- record window clamping ---------------------------------
  const rangeWindow = useCallback((): { start: number; end: number } => {
    const last = records.length - 1
    if (range === 'current') return { start: currentIndex, end: currentIndex }
    if (range === 'fromTo') {
      const a = Math.min(Math.max(from, 1), records.length)
      const b = Math.min(Math.max(to, a), records.length)
      return { start: a - 1, end: b - 1 }
    }
    return { start: 0, end: last }
  }, [currentIndex, from, range, records.length, to])

  // --- build mode + layout from dialog state ------------------
  const resolveLayout = (): { mode: MergeOutputMode; layout?: MergePageLayout } => {
    if (layoutFamily === 'label') {
      const layout = labelStock === 'avery5160'
        ? AVERY_5160
        : customGridLayout(paper, gridColumns, gridRows)
      return { mode: 'labels', layout }
    }
    if (layoutFamily === 'envelope') {
      return { mode: 'envelopes', layout: envelopeLayout(envelopeSize) }
    }
    return { mode: plainMode }
  }

  // --- fire the merge -----------------------------------------
  const runMerge = async () => {
    if (!source || records.length === 0) return
    setBusy(true)
    setError('')
    setDone('')
    try {
      const { mode, layout } = resolveLayout()
      const result = await onMerge({
        mode,
        records,
        fields: source.fields,
        namePattern,
        range: rangeWindow(),
        layout,
        labelStock,
        customGrid:
          layoutFamily === 'label' && labelStock === 'custom'
            ? { paper, columns: clampGridCount(gridColumns), rows: clampGridCount(gridRows) }
            : undefined,
        envelopePreset: layoutFamily === 'envelope' ? envelopeSize : undefined,
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

  // --- helpers: drop-hint & grid-size hint -------
  const labelGridHint =
    labelStock === 'avery5160'
      ? `${AVERY_5160.grid?.columns} × ${AVERY_5160.grid?.rows}`
      : `${clampGridCount(gridColumns)} × ${clampGridCount(gridRows)}`

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
                name="mailmerge-layout"
                checked={layoutFamily === 'plain'}
                onChange={() => setLayoutFamily('plain')}
              />
              {t('ribbonMailingsOutputPlain')}
            </label>
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-layout"
                checked={layoutFamily === 'label'}
                onChange={() => setLayoutFamily('label')}
              />
              {t('ribbonMailingsOutputLabels')}
            </label>
            <label className="mailmerge-radio">
              <input
                type="radio"
                name="mailmerge-layout"
                checked={layoutFamily === 'envelope'}
                onChange={() => setLayoutFamily('envelope')}
              />
              {t('ribbonMailingsOutputEnvelopes')}
            </label>
          </div>

          {/* ---- plain letters: per-record vs single-document ---- */}
          {layoutFamily === 'plain' && (
            <>
              <div className="modal-row mailmerge-option-row">
                <label className="mailmerge-radio">
                  <input
                    type="radio"
                    name="mailmerge-plain-mode"
                    checked={plainMode === 'single'}
                    onChange={() => setPlainMode('single')}
                  />
                  {t('ribbonMailingsOutputSingle')}
                </label>
                <label className="mailmerge-radio">
                  <input
                    type="radio"
                    name="mailmerge-plain-mode"
                    checked={plainMode === 'perRecord'}
                    onChange={() => setPlainMode('perRecord')}
                  />
                  {t('ribbonMailingsOutputPerRecord')}
                </label>
              </div>
              {plainMode === 'perRecord' && (
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
            </>
          )}

          {/* ---- labels: stock + custom grid ---- */}
          {layoutFamily === 'label' && (
            <div className="mailmerge-pattern">
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-stock"
                  checked={labelStock === 'avery5160'}
                  onChange={() => setLabelStock('avery5160')}
                />
                {t('ribbonMailingsStock5160')}
              </label>
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-stock"
                  checked={labelStock === 'custom'}
                  onChange={() => setLabelStock('custom')}
                />
                {t('ribbonMailingsStockCustom')}
              </label>
              <div className="mailmerge-hint">
                {t('ribbonMailingsLabelGridHint', { grid: labelGridHint })}
              </div>
              {labelStock === 'custom' && (
                <div className="modal-row">
                  <select
                    value={paper}
                    onChange={(e) => setPaper(e.target.value as 'letter' | 'a4')}
                  >
                    <option value="letter">Letter</option>
                    <option value="a4">A4</option>
                  </select>
                  <label className="mailmerge-radio">
                    {t('ribbonMailingsLabelColumns')}
                    <input
                      type="number"
                      min={1}
                      max={20}
                      value={gridColumns}
                      onChange={(e) => setGridColumns(clampGridCount(Number(e.target.value)))}
                    />
                  </label>
                  <label className="mailmerge-radio">
                    {t('ribbonMailingsLabelRows')}
                    <input
                      type="number"
                      min={1}
                      max={20}
                      value={gridRows}
                      onChange={(e) => setGridRows(clampGridCount(Number(e.target.value)))}
                    />
                  </label>
                </div>
              )}
              {droppedLabelBlocks > 0 && (
                <div className="mailmerge-note">
                  {t('ribbonMailingsLabelsDroppedHint', { n: droppedLabelBlocks })}
                </div>
              )}
            </div>
          )}

          {/* ---- envelopes: size preset ---- */}
          {layoutFamily === 'envelope' && (
            <div className="mailmerge-pattern">
              <label className="mailmerge-radio">
                {t('ribbonMailingsEnvelopeSize')}
                <select
                  value={envelopeSize}
                  onChange={(e) => setEnvelopeSize(e.target.value as EnvelopePresetId)}
                >
                  {ENVELOPE_PRESETS.map((entry) => (
                    <option key={entry.id} value={entry.id}>
                      {entry.name}
                    </option>
                  ))}
                </select>
              </label>
              <div className="mailmerge-note">{t('ribbonMailingsEnvelopeNote')}</div>
              <div className="modal-row">
                <span className="mailmerge-name-label">{t('ribbonMailingsNamePattern')}</span>
                <input
                  type="text"
                  className="mailmerge-name-input"
                  placeholder={t('ribbonMailingsNamePatternPh')}
                  value={namePattern}
                  onChange={(e) => setNamePattern(e.target.value)}
                />
              </div>
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
