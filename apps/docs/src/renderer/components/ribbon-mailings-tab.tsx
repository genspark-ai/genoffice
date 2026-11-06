import { useEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { mergeFieldInstruction } from '@genoffice/docx-engine'
import type { MergeDataSource, MergeDataSourceFile } from '../../shared/ipc'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import { useI18n } from '../i18n/locale'
import type { PmNode } from '../editor/convert'
import {
  countRecordSlots,
  hasNextField,
  recordPage,
  substituteMergeFields,
  substituteWithCursor,
} from '../mail-merge/template'
import { viewRecipients, type FilterCondition, type SortKey } from '../mail-merge/recipients'
import {
  MAIL_MERGE_PREVIEW_BYPASS,
  setPreviewReadOnly,
  setPreviewTemplate,
} from '../mail-merge/preview-extension'
import { MailMergePanel } from './MailMergePanel'
import { MergeDataSourceDialog } from './MergeDataSourceDialog'
import { MergeRecipientList } from './MergeRecipientList'
import {
  IconCaret,
  IconChevronLeft,
  IconChevronRight,
  IconEye,
  IconEyeOff,
  IconMail,
  IconMergeField,
  IconTable,
} from './icons'
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
  /**
   * The path of the document in the editor. The preview belongs to the document
   * it was started on: when another document takes its place, the preview leaves
   * instead of writing the old template over the new one.
   */
  docPath?: string | null
  dropdown: string | null
  setDropdown: SetDropdown
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** pick the file the recipient list comes from; null when the user canceled */
  onPickFile: () => Promise<MergeDataSourceFile | null>
  /** read the confirmed sheet as the data source; false when it failed */
  onReadSource: (path: string, sheet: string) => Promise<boolean>
  /** insert a merge field at the caret */
  onInsertField: (instr: string) => void
  /** build the documents and write them */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
  /**
   * The preview holds the editor's document, so the host has to know: the paths
   * that write the document out or read it for the model are fenced off while
   * it is on. Called on every enter and every leave, leaving the document in
   * its hands.
   */
  onPreviewingChange?: (previewing: boolean) => void
}

export function MailingsTab({
  editor,
  hasDoc,
  docPath,
  dropdown,
  setDropdown,
  source,
  onPickFile,
  onReadSource,
  onInsertField,
  onMerge,
  onPreviewingChange,
}: MailingsTabProps) {
  const { t } = useI18n()
  const [listOpen, setListOpen] = useState(false)
  const [sourceOpen, setSourceOpen] = useState(false)
  const [panelOpen, setPanelOpen] = useState(false)
  /** the document with the merge fields in it, swapped for the preview */
  const [previewing, setPreviewing] = useState(false)
  const [recordIndex, setRecordIndex] = useState(0)
  /** the file the user picked in the data-source dialog */
  const [pickedFile, setPickedFile] = useState<MergeDataSourceFile | null>(null)
  /** the sheet of that file the merge reads */
  const [pickedSheet, setPickedSheet] = useState('')
  /** why the last read of the data source failed */
  const [sourceError, setSourceError] = useState('')
  /** the records the user took out of the merge */
  const [excluded, setExcluded] = useState<Set<number>>(() => new Set())
  /** the filter conditions the recipient list applies */
  const [conditions, setConditions] = useState<FilterCondition[]>([])
  /** the sort levels: at most three, each with a column and its direction */
  const [sorts, setSorts] = useState<SortKey[]>([])
  // the preview swaps the editor's document for the record's values; this holds
  // the document as it was so it can be put back on exit, and survives re-renders
  const savedRef = useRef<PmNode | null>(null)
  const previewingRef = useRef(false)
  /** the document the preview was started on: another one taking its place ends it */
  const previewDocPathRef = useRef<string | null | undefined>(undefined)
  // a new list starts at its first record, with nothing taken out and no filter or sort
  useEffect(() => {
    setRecordIndex(0)
    setExcluded(new Set())
    setConditions([])
    setSorts([])
  }, [source?.path])

  // the list shows the source order or the user's sort, narrowed by the filter
  // conditions
  const viewRows = useMemo(
    () => viewRecipients(source?.records ?? [], conditions, sorts),
    [source?.records, conditions, sorts],
  )
  // the records the merge walks are the visible ones, minus those the user took out
  const activeRecords = useMemo(() => {
    return viewRows.filter((row) => !excluded.has(row.index)).map((row) => row.record)
  }, [viewRows, excluded])
  const hasRecords = activeRecords.length > 0
  // taking records out can leave the cursor past the end of the list
  const activeIndex =
    activeRecords.length === 0 ? 0 : Math.min(recordIndex, activeRecords.length - 1)

  const insertField = (name: string) => {
    onInsertField(mergeFieldInstruction(name))
    setDropdown(() => null)
  }
  /** insert a NEXT rule: advance to the next record when merging (one page, many records) */
  const insertNext = () => onInsertField('NEXT')

  /** pick the file the recipient list comes from, then name its sheet */
  const openSourceDialog = async () => {
    const file = await onPickFile()
    if (!file) return
    setPickedFile(file)
    // the sheet the current source uses, when it is the same file
    setPickedSheet(
      source && source.path === file.path && source.sheet ? source.sheet : (file.sheets[0] ?? ''),
    )
    setSourceOpen(true)
  }

  /** read the confirmed sheet and go on to the merge */
  const confirmSource = async () => {
    if (!pickedFile) return
    setSourceError('')
    try {
      const ok = await onReadSource(pickedFile.path, pickedSheet)
      if (!ok) {
        setSourceError(t('ribbonMailingsSourceFailed'))
        return
      }
      setSourceOpen(false)
    } catch (err) {
      setSourceError(err instanceof Error ? err.message : String(err))
    }
  }

  /**
   * The document as the preview shows it: the active record's values in place of
   * the fields. A NEXT template fills every slot on the page with the records
   * following the page's first record (like Word's preview); the plain template
   * keeps its fields as values.
   */
  const previewDoc = (doc: PmNode, recordIndex: number): PmNode => {
    if (hasNextField(doc)) {
      const slots = countRecordSlots(doc)
      const start = recordPage(activeRecords.length, recordIndex, slots).start
      return substituteWithCursor(doc, activeRecords, start, start + slots - 1)
    }
    return substituteMergeFields(doc, activeRecords[recordIndex])
  }

  /**
   * Swap the editor's document for the active record's merged values, so the
   * fields read as the record rather than as «Name». emitUpdate:false keeps the
   * editor's dirty flag and the autosave out of it — previewing is a view, not
   * an edit. addToHistory:false keeps the preview out of the undo stack so
   * Ctrl+Z during preview does not desync the saved snapshot. A second press
   * restores the document with the fields back.
   */
  const togglePreview = () => {
    if (!previewingRef.current) {
      const record = activeRecords[activeIndex]
      if (!record) return
      savedRef.current = editor.getJSON() as PmNode
      // tell everyone who reads editor.getJSON() to use the saved original instead
      setPreviewTemplate(editor, savedRef.current)
      previewDocPathRef.current = docPath
      setPreviewReadOnly(editor, true)
      // the bypass meta lets this swap through the preview lock the App turns
      // on right after; without it the merged document could not even appear
      editor
        .chain()
        .setMeta('addToHistory', false)
        .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
        .setContent(previewDoc(savedRef.current, activeIndex) as never, { emitUpdate: false })
        .run()
      previewingRef.current = true
      setPreviewing(true)
      onPreviewingChange?.(true)
    } else {
      if (savedRef.current) {
        // restoring the template happens while the lock is still on: bypass it
        editor
          .chain()
          .setMeta('addToHistory', false)
          .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
          .setContent(savedRef.current as never, { emitUpdate: false })
          .run()
      }
      setPreviewTemplate(editor, null)
      savedRef.current = null
      previewDocPathRef.current = undefined
      setPreviewReadOnly(editor, false)
      previewingRef.current = false
      setPreviewing(false)
      onPreviewingChange?.(false)
    }
  }

  // walking to another record re-merges the preview to that record
  useEffect(() => {
    if (!previewingRef.current || !savedRef.current) return
    if (activeRecords[activeIndex])
      editor
        .chain()
        .setMeta('addToHistory', false)
        .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
        .setContent(previewDoc(savedRef.current, activeIndex) as never, { emitUpdate: false })
        .run()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeIndex])

  // a new data source means the saved document is stale: leave the preview
  useEffect(() => {
    if (!previewingRef.current || !savedRef.current) return
    editor
      .chain()
      .setMeta('addToHistory', false)
      .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
      .setContent(savedRef.current as never, { emitUpdate: false })
      .run()
    setPreviewTemplate(editor, null)
    savedRef.current = null
    setPreviewReadOnly(editor, false)
    previewingRef.current = false
    setPreviewing(false)
    onPreviewingChange?.(false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source?.path])

  // another document took the editor's place: the preview belonged to the one
  // that was there before, so it ends here — without writing the old template
  // over the new document
  useEffect(() => {
    if (!previewingRef.current) return
    if (docPath === previewDocPathRef.current) return
    setPreviewTemplate(editor, null)
    savedRef.current = null
    previewDocPathRef.current = undefined
    setPreviewReadOnly(editor, false)
    previewingRef.current = false
    setPreviewing(false)
    onPreviewingChange?.(false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [docPath, editor])

  // the tab goes inert (unmount) with the document still swapped: put it back
  useEffect(
    () => () => {
      if (previewingRef.current && savedRef.current) {
        editor
          .chain()
          .setMeta('addToHistory', false)
          .setMeta(MAIL_MERGE_PREVIEW_BYPASS, true)
          .setContent(savedRef.current as never, { emitUpdate: false })
          .run()
        setPreviewTemplate(editor, null)
        setPreviewReadOnly(editor, false)
        previewingRef.current = false
        onPreviewingChange?.(false)
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [],
  )

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc || previewing}
            data-tip={t('ribbonMailingsStartTip')}
            onClick={() => void openSourceDialog()}
          >
            <span className="rb-big-icon">
              <IconMail size={BIG} />
            </span>
            <span>{t('ribbonMailingsStart')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonMailingsRecipientListTip')}
            onClick={() => setListOpen(true)}
          >
            <span className="rb-big-icon">
              <IconTable size={BIG} />
            </span>
            <span>{t('ribbonMailingsRecipientList')}</span>
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
              disabled={!hasDoc || previewing || !source || source.fields.length === 0}
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
                <button
                  key="__next"
                  className="rb-field-item rb-field-next"
                  data-tip={t('ribbonMailingsInsertNextTip')}
                  onClick={() => {
                    insertNext()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonMailingsInsertNext')}
                </button>
                {source.fields.length > 0 && <div className="rb-drop-sep" />}
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
            data-tip={t('ribbonMailingsPreviewTip')}
            onClick={togglePreview}
          >
            <span className="rb-big-icon">
              {previewing ? <IconEyeOff size={BIG} /> : <IconEye size={BIG} />}
            </span>
            <span>{previewing ? t('ribbonMailingsExitPreview') : t('ribbonMailingsPreview')}</span>
          </button>
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
            onClick={() => setRecordIndex((index) => Math.min(index + 1, activeRecords.length - 1))}
          >
            <span className="rb-big-icon">
              <IconChevronRight size={BIG} />
            </span>
            <span>{t('ribbonMailingsNext')}</span>
          </button>
          {/* merge-outcome: the mail merge panel, with the finish section in it */}
          <button
            className="rb-big rb-big-primary"
            disabled={!hasRecords}
            data-tip={t('ribbonMailingsFinishTip')}
            onClick={() => setPanelOpen(true)}
          >
            <span className="rb-big-icon">
              <IconMail size={BIG} />
            </span>
            <span>{t('ribbonMailingsFinish')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">
          {hasRecords
            ? t('ribbonMailingsRecord', { n: activeIndex + 1, total: activeRecords.length })
            : t('ribbonMailingsNoSource')}
        </div>
      </div>

      {sourceOpen && (
        <MergeDataSourceDialog
          file={pickedFile}
          sheet={pickedSheet}
          onSheet={setPickedSheet}
          onPickFile={() => void openSourceDialog()}
          onConfirm={() => void confirmSource()}
          error={sourceError}
          onClose={() => setSourceOpen(false)}
        />
      )}

      {listOpen && (
        <MergeRecipientList
          source={source}
          rows={viewRows}
          excluded={excluded}
          onExcluded={setExcluded}
          conditions={conditions}
          onConditions={setConditions}
          sorts={sorts}
          onSorts={setSorts}
          onClose={() => setListOpen(false)}
        />
      )}

      {/*
        the merge panel: the output shape (letter / label sheet / envelope) and,
        inside it, the finish section that writes the chosen record range.
      */}
      {panelOpen && (
        <MailMergePanel
          editor={editor}
          source={source}
          records={activeRecords}
          onPickSource={() => void openSourceDialog()}
          onMerge={onMerge}
          recordIndex={activeIndex}
          onClose={() => setPanelOpen(false)}
        />
      )}
    </>
  )
}
