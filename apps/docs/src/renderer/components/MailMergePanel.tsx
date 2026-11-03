import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { mergeFieldInstruction } from '@genoffice/docx-engine'
import type { MergeDataSource, MergeOutputMode } from '../../shared/ipc'
import type { PmNode } from '../editor/convert'
import { documentText, substituteMergeFields } from '../mail-merge/template'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import {
  AVERY_5160,
  clampGridCount,
  customGridLayout,
  envelopeLayout,
  unsupportedLabelBlocks,
  type EnvelopePresetId,
  type MergePageLayout,
} from '../mail-merge/layout'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * The mail merge panel: the data source, the fields it holds, the preview of
 * one record and the merge itself.
 *
 * The panel asks for things and shows what comes back — the record list comes
 * from the main process, the preview from the merge module, the documents from
 * the app. It holds no merge logic of its own.
 * genoffice#2097.
 */

export interface MailMergePanelProps {
  editor: Editor
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** open the file dialog and read the recipient list */
  onPickSource: () => void
  /** insert a merge field at the caret */
  onInsertField: (instr: string) => void
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
  /** the record the preview shows */
  recordIndex: number
  onRecordIndex: (index: number) => void
  onClose: () => void
}

export function MailMergePanel({
  editor,
  source,
  onPickSource,
  onInsertField,
  onMerge,
  recordIndex,
  onRecordIndex,
  onClose,
}: MailMergePanelProps) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [mode, setMode] = useState<MergeOutputMode>('perRecord')
  const [namePattern, setNamePattern] = useState('')
  // the label stock: the built-in sheet, or an even custom grid on Letter/A4
  const [labelStock, setLabelStock] = useState<'avery5160' | 'custom'>('avery5160')
  const [paper, setPaper] = useState<'letter' | 'a4'>('letter')
  const [gridColumns, setGridColumns] = useState(3)
  const [gridRows, setGridRows] = useState(10)
  const [envelopeSize, setEnvelopeSize] = useState<EnvelopePresetId>('size10')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState('')

  /** the paper the label or envelope modes print on, from the options above */
  const layoutOf = (): MergePageLayout | undefined => {
    if (mode === 'labels') {
      return labelStock === 'avery5160'
        ? AVERY_5160
        : customGridLayout(paper, gridColumns, gridRows)
    }
    if (mode === 'envelopes') return envelopeLayout(envelopeSize)
    return undefined
  }

  // a label cell only takes paragraphs: tell the user when the template has
  // blocks the sheet would drop, instead of letting the merge waste a sheet
  const droppedLabelBlocks = useMemo(
    () => (mode === 'labels' ? unsupportedLabelBlocks(editor.getJSON() as PmNode) : 0),
    [editor, mode],
  )

  const records = source?.records ?? []
  const record = records[recordIndex]
  // the template with this record's values in it, as the user reads it
  const preview = useMemo(() => {
    if (!record) return ''
    return documentText(substituteMergeFields(editor.getJSON() as PmNode, record))
  }, [editor, record])

  const insertField = (name: string) => {
    onInsertField(mergeFieldInstruction(name))
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
        layout: layoutOf(),
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
        <h2>{t('ribbonMailingsPanelTitle')}</h2>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsSource')}</div>
          <div className="modal-row">
            <span className="mailmerge-source-name">
              {source ? source.name : t('ribbonMailingsNoSource')}
            </span>
            <button onClick={onPickSource}>{t('ribbonMailingsSelectSource')}</button>
          </div>
        </div>

        <div className="mailmerge-section">
          <div className="mailmerge-label">{t('ribbonMailingsFields')}</div>
          {source && source.fields.length > 0 ? (
            <div className="mailmerge-fields">
              {source.fields.map((name) => (
                <button
                  key={name}
                  className="mailmerge-field"
                  data-tip={t('ribbonMailingsInsertFieldTip')}
                  onClick={() => insertField(name)}
                >
                  {name}
                </button>
              ))}
            </div>
          ) : (
            <div className="mailmerge-empty">{t('ribbonMailingsNoFields')}</div>
          )}
        </div>

        {source && records.length > 0 && (
          <div className="mailmerge-section">
            <div className="mailmerge-label">{t('ribbonMailingsPreview')}</div>
            <div className="modal-row">
              <button disabled={recordIndex <= 0} onClick={() => onRecordIndex(recordIndex - 1)}>
                {t('ribbonMailingsPrev')}
              </button>
              <span className="mailmerge-counter">
                {t('ribbonMailingsRecord', { n: recordIndex + 1, total: records.length })}
              </span>
              <button
                disabled={recordIndex >= records.length - 1}
                onClick={() => onRecordIndex(recordIndex + 1)}
              >
                {t('ribbonMailingsNext')}
              </button>
            </div>
            <pre className="mailmerge-preview">{preview}</pre>
          </div>
        )}

        {source && records.length > 0 && (
          <div className="mailmerge-section">
            <div className="mailmerge-label">{t('ribbonMailingsOutput')}</div>
            <div className="modal-row">
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-output"
                  checked={mode === 'perRecord'}
                  onChange={() => setMode('perRecord')}
                />
                {t('ribbonMailingsOutputPerRecord')}
              </label>
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-output"
                  checked={mode === 'single'}
                  onChange={() => setMode('single')}
                />
                {t('ribbonMailingsOutputSingle')}
              </label>
            </div>
            <div className="modal-row">
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-output"
                  checked={mode === 'labels'}
                  onChange={() => setMode('labels')}
                />
                {t('ribbonMailingsOutputLabels')}
              </label>
              <label className="mailmerge-radio">
                <input
                  type="radio"
                  name="mailmerge-output"
                  checked={mode === 'envelopes'}
                  onChange={() => setMode('envelopes')}
                />
                {t('ribbonMailingsOutputEnvelopes')}
              </label>
            </div>
            {mode === 'labels' && (
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
                {droppedLabelBlocks > 0 && (
                  <div className="mailmerge-note">
                    {t('ribbonMailingsLabelsDroppedHint', { n: droppedLabelBlocks })}
                  </div>
                )}
                {labelStock === 'custom' && (
                  <div className="modal-row">
                    <select
                      value={paper}
                      onChange={(e) => setPaper(e.target.value as 'letter' | 'a4')}
                    >
                      <option value="letter">Letter</option>
                      <option value="a4">A4</option>
                    </select>
                    <label>
                      {t('ribbonMailingsLabelColumns')}
                      <input
                        type="number"
                        min={1}
                        max={20}
                        value={gridColumns}
                        onChange={(e) => setGridColumns(clampGridCount(Number(e.target.value)))}
                      />
                    </label>
                    <label>
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
              </div>
            )}
            {mode === 'envelopes' && (
              <>
                <label className="mailmerge-pattern">
                  {t('ribbonMailingsEnvelopeSize')}
                  <select
                    value={envelopeSize}
                    onChange={(e) => setEnvelopeSize(e.target.value as EnvelopePresetId)}
                  >
                    <option value="size10">Size 10</option>
                    <option value="dl">DL</option>
                    <option value="c5">C5</option>
                    <option value="c6">C6</option>
                  </select>
                </label>
                <div className="mailmerge-note">{t('ribbonMailingsEnvelopeNote')}</div>
              </>
            )}
            {(mode === 'perRecord' || mode === 'envelopes') && (
              <label className="mailmerge-pattern">
                {t('ribbonMailingsNamePattern')}
                <input
                  value={namePattern}
                  placeholder={t('ribbonMailingsNamePatternPh')}
                  onChange={(e) => setNamePattern(e.target.value)}
                />
              </label>
            )}
            <div className="modal-row">
              <button className="btn-primary" disabled={busy} onClick={() => void runMerge()}>
                {t('ribbonMailingsFinish')}
              </button>
            </div>
          </div>
        )}

        {error && <div className="modal-error">{error}</div>}
        {done && <div className="mailmerge-done">{done}</div>}

        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonMailingsClose')}</button>
        </div>
      </div>
    </div>
  )
}
