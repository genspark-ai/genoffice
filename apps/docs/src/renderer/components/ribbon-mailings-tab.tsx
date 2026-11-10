import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { mergeFieldInstruction } from '@genoffice/docx-engine'
import type { MergeDataSource } from '../../shared/ipc'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import { useI18n } from '../i18n/locale'
import { MailMergePanel } from './MailMergePanel'
import { IconCaret, IconChevronLeft, IconChevronRight, IconMail, IconMergeField } from './icons'
import { BIG, type SetDropdown, toggleDropdown } from './ribbon-tabs'

/**
 * The Mailings ribbon tab: the data source, the fields it holds, the preview
 * of a record and the merge.
 *
 * The tab holds what the user picked (the data source) and which record the
 * preview shows; the documents themselves are built by the app and written by
 * the main process. genoffice#2097.
 */

export interface MailingsTabProps {
  editor: Editor
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** open the file dialog and read the recipient list */
  onPickSource: () => void
  /** insert a merge field at the caret */
  onInsertField: (instr: string) => void
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
}

export function MailingsTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  source,
  onPickSource,
  onInsertField,
  onMerge,
}: MailingsTabProps) {
  const { t } = useI18n()
  const [panelOpen, setPanelOpen] = useState(false)
  const [recordIndex, setRecordIndex] = useState(0)
  // a new list starts at its first record; the old index may be past its end
  useEffect(() => setRecordIndex(0), [source?.path])

  const records = source?.records ?? []
  const hasRecords = records.length > 0
  const insertField = (name: string) => {
    onInsertField(mergeFieldInstruction(name))
    setDropdown(() => null)
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonMailingsStartTip')}
            onClick={() => setPanelOpen(true)}
          >
            <span className="rb-big-icon">
              <IconMail size={BIG} />
            </span>
            <span>{t('ribbonMailingsStart')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupMailMerge')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || !source || source.fields.length === 0}
              data-tip={t('ribbonMailingsInsertFieldTip')}
              onClick={() => toggleDropdown(setDropdown, 'mergefield')}
            >
              <span className="rb-big-icon">
                <IconMergeField size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonMailingsInsertField')}</span>
            </button>
            {dropdown === 'mergefield' && source && (
              <div data-rb-panel="" className="rb-field-list">
                <div className="rb-drop-title">{t('ribbonMailingsFields')}</div>
                {source.fields.map((name) => (
                  <button key={name} className="rb-field-item" onClick={() => insertField(name)}>
                    {name}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonMailingsInsertField')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasRecords}
            data-tip={t('ribbonMailingsPrevTip')}
            onClick={() => setRecordIndex((index) => Math.max(index - 1, 0))}
          >
            <span className="rb-big-icon">
              <IconChevronLeft size={BIG} />
            </span>
            <span>{t('ribbonMailingsPrev')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasRecords}
            data-tip={t('ribbonMailingsNextTip')}
            onClick={() => setRecordIndex((index) => Math.min(index + 1, records.length - 1))}
          >
            <span className="rb-big-icon">
              <IconChevronRight size={BIG} />
            </span>
            <span>{t('ribbonMailingsNext')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">
          {hasRecords
            ? t('ribbonMailingsRecord', { n: recordIndex + 1, total: records.length })
            : t('ribbonMailingsNoSource')}
        </div>
      </div>

      {panelOpen && (
        <MailMergePanel
          editor={editor}
          source={source}
          onPickSource={onPickSource}
          onInsertField={onInsertField}
          onMerge={onMerge}
          recordIndex={recordIndex}
          onRecordIndex={setRecordIndex}
          onClose={() => setPanelOpen(false)}
        />
      )}
    </>
  )
}
