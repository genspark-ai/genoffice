import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { MergeRecord } from '@genoffice/docx-engine'
import type { MergeDataSource, MergeOutputMode } from '../../shared/ipc'
import type { PmNode } from '../editor/convert'
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
import { MergeFinishDialog } from './MergeFinishDialog'

/**
 * The mail merge panel: the data source and the merge output options.
 *
 * The panel asks for things and shows what comes back — the record list comes
 * from the main process and the documents from the app. It holds no merge logic
 * of its own. genoffice#2097.
 */

export interface MailMergePanelProps {
  editor: Editor
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** the records the merge walks: the source's, minus the ones taken out */
  records: readonly MergeRecord[]
  /** open the file dialog and read the recipient list */
  onPickSource: () => void
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
  /** the record the finish section starts its range at */
  recordIndex: number
  onClose: () => void
}

export function MailMergePanel({
  editor,
  source,
  records,
  onPickSource,
  onMerge,
  recordIndex,
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

            <MergeFinishDialog
              source={source}
              records={records}
              currentIndex={recordIndex}
              editor={editor}
              mode={mode}
              layout={layoutOf()}
              namePattern={namePattern}
              onMerge={onMerge}
            />
          </div>
        )}

        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonMailingsClose')}</button>
        </div>
      </div>
    </div>
  )
}
