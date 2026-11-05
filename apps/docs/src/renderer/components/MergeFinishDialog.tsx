import { useCallback, useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { PmNode } from '../editor/convert'
import type { MergeDataSource, MergeOutputMode } from '../../shared/ipc'
import type { MergeRecord } from '@genoffice/docx-engine'
import type { MailMergeRequest, MailMergeResponse } from '../mail-merge/output'
import type { MergePageLayout } from '../mail-merge/layout'
import { countRecordSlots, mergePageCount, recordPage } from '../mail-merge/template'
import { useI18n } from '../i18n/locale'

/**
 * The "Finish & Merge" part of the mail merge: the record range, then the
 * merge itself.
 *
 * The panel that hosts it has already asked for the output shape — one batch
 * document, one file per record, a label sheet, or one envelope per record —
 * and holds the paper it prints on. This only asks which records are merged;
 * the result opens in the app (a new tab per document) and the user saves it
 * when they want it on disk. The range has no default, the user picks one.
 *
 * genoffice#2097, genoffice#2374.
 */

export interface MergeFinishDialogProps {
  /** the data source the user picked, null when there is none yet */
  source: MergeDataSource | null
  /** the records the merge walks: the ones the recipient list kept */
  records: readonly MergeRecord[]
  /** the record the user is sitting on, for the "current record" range */
  currentIndex: number
  /** the editor the template lives in: the page groups come from it */
  editor: Editor
  /** the output shape the mail merge panel picked */
  mode: MergeOutputMode
  /** the paper the label or the envelope output prints on */
  layout?: MergePageLayout
  /** the field the file names come from; the multi-file outputs need it */
  namePattern?: string
  /** build the documents and open them in the app */
  onMerge: (request: MailMergeRequest) => Promise<MailMergeResponse>
}

/** the range choices; '' means the user has not picked one yet */
type RangeChoice = '' | 'all' | 'current' | 'fromTo'

/**
 * The template the user wrote. Preview swaps the editor's content for merged
 * values, so the original sits on the editor's storage while it is on.
 */
function templateDocOf(editor: Editor): PmNode {
  const storage = editor.storage as unknown as Record<string, unknown>
  return (storage.mailmergePreviewJson as PmNode | undefined) ?? (editor.getJSON() as PmNode)
}

export function MergeFinishDialog({
  source,
  records,
  currentIndex,
  editor,
  mode,
  layout,
  namePattern = '',
  onMerge,
}: MergeFinishDialogProps) {
  const { t } = useI18n()

  // --- the record range the merge writes ---
  const [outputRange, setOutputRange] = useState<RangeChoice>('')
  const [from, setFrom] = useState(1)
  const [to, setTo] = useState(Math.max(records.length, 1))

  // --- dialog state -------------------------------------------
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [done, setDone] = useState('')

  // --- the page group "current record" stands for -------------
  // A NEXT template merges as many records per page as it has record blocks —
  // not NEXT count + 1: the NEXT at the end of the page hands off to the next
  // page and renders nothing. Without NEXT it is one record per page.
  const slotsPerPage = useMemo(() => countRecordSlots(templateDocOf(editor)), [editor])

  // Preview renders one record at a time, so the page holding currentIndex is
  // derived from the index itself: slot 1 of that page is the previewed record.
  const currentPage = recordPage(records.length, currentIndex, slotsPerPage)
  const currentPageStart = currentPage.start
  const currentPageEnd = currentPage.end
  const currentPageCount = currentPage.count
  const lastRecord = records.length - 1

  /** the per-record output writes one file per record; the other modes are single documents */
  const multiFile = mode === 'perRecord'

  // --- record window clamping ---------------------------------
  const maxRecord = Math.max(records.length, 1)
  /** a typed range number: whole, at least 1, never past the last record */
  const clampInput = (value: number): number =>
    Math.min(Math.max(Math.floor(value) || 1, 1), maxRecord)

  const windowFor = useCallback(
    (choice: RangeChoice): { start: number; end: number } => {
      // "current" is the whole page group the previewed record sits on; "fromTo"
      // stays absolute record indices and is never page-aligned.
      if (choice === 'current') return { start: currentPageStart, end: currentPageEnd }
      if (choice === 'fromTo') {
        const a = Math.min(Math.max(from, 1), records.length)
        const b = Math.min(Math.max(to, a), records.length)
        return { start: a - 1, end: b - 1 }
      }
      if (choice === 'all') return { start: 0, end: lastRecord }
      // nothing picked yet: an empty window
      return { start: 0, end: -1 }
    },
    [currentPageEnd, currentPageStart, from, lastRecord, records.length, to],
  )

  const totalPages = mergePageCount(records.length, slotsPerPage)
  const fromPage = Math.floor((from - 1) / slotsPerPage) + 1
  const fromSlot = ((from - 1) % slotsPerPage) + 1

  /** the three range choices for the merge output */
  const rangePicker = (choice: RangeChoice, pick: (c: RangeChoice) => void, name: string) => (
    <div className="modal-row mailmerge-option-row">
      <label className="mailmerge-radio">
        <input type="radio" name={name} checked={choice === 'all'} onChange={() => pick('all')} />
        {t('ribbonMailingsRangeAll', { n: records.length })}
      </label>
      <label className="mailmerge-radio">
        <input
          type="radio"
          name={name}
          checked={choice === 'current'}
          onChange={() => pick('current')}
        />
        {slotsPerPage > 1
          ? currentPageCount > 1
            ? t('ribbonMailingsRangeCurrentPage', {
                from: currentPageStart + 1,
                to: currentPageEnd + 1,
              })
            : t('ribbonMailingsRangeCurrentPageOne', { n: currentPageStart + 1 })
          : t('ribbonMailingsRangeCurrent', { n: currentIndex + 1 })}
      </label>
      <label className="mailmerge-radio">
        <input
          type="radio"
          name={name}
          checked={choice === 'fromTo'}
          onChange={() => pick('fromTo')}
        />
        {t('ribbonMailingsRangeFromTo')}
        <input
          className="mailmerge-range-num"
          type="number"
          min={1}
          max={records.length}
          value={from}
          onChange={(e) => setFrom(clampInput(Number(e.target.value)))}
        />
        {t('ribbonMailingsRangeTo')}
        <input
          className="mailmerge-range-num"
          type="number"
          min={1}
          max={records.length}
          value={to}
          onChange={(e) => setTo(clampInput(Number(e.target.value)))}
        />
      </label>
    </div>
  )

  /** the page count the records take, and what the chosen window writes */
  const rangeHint = (choice: RangeChoice) => {
    if (records.length === 0) return null
    if (!choice) return <div className="mailmerge-hint">{t('ribbonMailingsPickRange')}</div>
    const window_ = windowFor(choice)
    const count = Math.max(0, window_.end - window_.start + 1)
    const pages = mergePageCount(count, slotsPerPage)
    return (
      <div className="mailmerge-hint">
        {t('ribbonMailingsRangeSummary', {
          total: records.length,
          perPage: slotsPerPage,
          pages: totalPages,
        })}
        {choice === 'fromTo' && (
          <>
            {' · '}
            {t('ribbonMailingsRangeSlot', { n: from, page: fromPage, slot: fromSlot })}
          </>
        )}
        {' · '}
        {t('ribbonMailingsRangeWillExport', { n: count, pages })}
      </div>
    )
  }

  // --- fire the merge -----------------------------------------
  const runMerge = async (choice: RangeChoice) => {
    // a double click must not write twice: `busy` is the re-entry latch
    if (busy || !source || records.length === 0 || !choice) return
    setBusy(true)
    setError('')
    setDone('')
    try {
      const range = windowFor(choice)
      const count = Math.max(0, range.end - range.start + 1)
      const pages = mergePageCount(count, slotsPerPage)
      // The app builds the merged documents (it owns the document context the
      // merge needs) and opens them in the app — one file for the batch/label/
      // envelope outputs, one file per record for the per-record output.
      const res = await onMerge({
        mode,
        records,
        fields: source.fields,
        namePattern: multiFile ? namePattern : '',
        range,
        layout,
      })
      if (!res.ok) {
        setError(res.error ?? t('ribbonMailingsMergeFailed'))
        return
      }
      setDone(
        t('ribbonMailingsMergedPages', { files: multiFile ? count : 1, pages: pages || 1 }) +
          ' ' +
          t('ribbonMailingsOpenedInNewTab'),
      )
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  const body = (
    <div className="mailmerge-range-block">
      <div className="mailmerge-label">{t('ribbonMailingsRange')}</div>
      {rangePicker(outputRange, setOutputRange, 'mailmerge-range')}
      {rangeHint(outputRange)}
      {records.length === 0 && (
        <div className="mailmerge-hint">{t('ribbonMailingsNoRecipients')}</div>
      )}
      {error && <div className="modal-error">{error}</div>}
      {done && <div className="mailmerge-done">{done}</div>}
      <div className="modal-actions">
        <button
          className="btn-primary"
          disabled={busy || !source || records.length === 0 || !outputRange}
          onClick={() => void runMerge(outputRange)}
        >
          {t('ribbonMailingsFinish')}
        </button>
      </div>
    </div>
  )

  // the mail merge panel hosts this inside its output section: the record range
  // and the finish button, wrapped so it nests under the single section heading
  return <div className="mailmerge-finish">{body}</div>
}
