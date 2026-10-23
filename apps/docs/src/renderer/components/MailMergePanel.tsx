import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { mergeFieldInstruction } from '@genoffice/docx-engine'
import type { MergeDataSource } from '../../shared/ipc'
import type { PmNode } from '../editor/convert'
import { documentText, substituteMergeFields } from '../mail-merge/template'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
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
  const [mode, setMode] = useState<'single' | 'perRecord'>('perRecord')
  const [namePattern, setNamePattern] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState('')

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
            {mode === 'perRecord' && (
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
