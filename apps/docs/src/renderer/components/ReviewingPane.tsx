import { useEffect, useRef, type KeyboardEvent } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { revisionCount, type RevisionRange } from '../editor/revisions'
import { summarizeRevisions, type ReviewEntry } from '../editor/revision-filters'
import { STORY_LABEL } from '../editor/story-revisions'
import { absoluteDate, inkVar } from '../editor/comment-card'
import { revWhen } from '../editor/margin-annotations'
import { IconTrackChanges } from './icons'

export const KIND_LABEL: Record<RevisionRange['kind'], StringKey> = {
  ins: 'appRevKindIns',
  del: 'appRevKindDel',
  rPrChange: 'appRevKindFormat',
  pPrChange: 'appRevKindParaFormat',
  paraMarkRPrChange: 'appRevKindParaFormat',
  paraMarkIns: 'appRevKindParaMarkIns',
  paraMarkDel: 'appRevKindParaMarkDel',
  moveFrom: 'appRevKindMoveFrom',
  moveTo: 'appRevKindMoveTo',
  rowIns: 'appRevKindRowIns',
  rowDel: 'appRevKindRowDel',
  cellIns: 'appRevKindCellIns',
  cellDel: 'appRevKindCellDel',
  blockIns: 'appRevKindBlockIns',
  blockDel: 'appRevKindBlockDel',
  tblPrChange: 'appRevKindTableFormat',
  tblGridChange: 'appRevKindTableFormat',
  trPrChange: 'appRevKindTableFormat',
  tblPrExChange: 'appRevKindTableFormat',
  tcPrChange: 'appRevKindTableFormat',
  cellMerge: 'appRevKindCellMerge',
  numberingChange: 'appRevKindNumbering',
  sectPrChange: 'appRevKindSectionFormat',
}

const DELETED_KINDS = new Set<RevisionRange['kind']>([
  'del',
  'moveFrom',
  'rowDel',
  'cellDel',
  'blockDel',
  'paraMarkDel',
])

const INSERTED_KINDS = new Set<RevisionRange['kind']>([
  'ins',
  'moveTo',
  'rowIns',
  'cellIns',
  'blockIns',
  'paraMarkIns',
])

/** One tracked change as a Word Reviewing Pane row: "Kind  Author  date", the text, Accept / Reject on hover */
export function ReviewRow({
  entry: e,
  active,
  canEdit,
  onSelect,
  onAccept,
  onReject,
}: {
  entry: ReviewEntry
  active: boolean
  canEdit: boolean
  onSelect: (rev: RevisionRange) => void
  onAccept: (rev: RevisionRange) => void
  onReject: (rev: RevisionRange) => void
}) {
  const { t } = useI18n()
  const tone = DELETED_KINDS.has(e.rev.kind)
    ? ' review-del'
    : INSERTED_KINDS.has(e.rev.kind)
      ? ' review-ins'
      : ''
  const onKey = (ev: KeyboardEvent<HTMLDivElement>) => {
    if (ev.target !== ev.currentTarget) return
    if (ev.key === 'Enter' || ev.key === ' ') {
      ev.preventDefault()
      onSelect(e.rev)
    }
  }
  return (
    <div
      className={`review-row${tone}${active ? ' active' : ''}`}
      data-rev-index={e.index}
      role="button"
      tabIndex={0}
      onClick={() => onSelect(e.rev)}
      onKeyDown={onKey}
    >
      <div className="review-row-head">
        <span className="review-row-kind" style={{ color: inkVar(e.inkIndex) }}>
          {t(KIND_LABEL[e.rev.kind])}
          {e.rev.story && ` · ${t(STORY_LABEL[e.rev.story.kind])}`}
        </span>
        <span className="review-row-author">{e.rev.author || t('editorUnknownAuthor')}</span>
        {e.rev.date && (
          <span className="review-row-when" title={absoluteDate(e.rev.date)}>
            {revWhen(e.rev.date)}
          </span>
        )}
      </div>
      {e.excerpt && <div className="review-row-text">{e.excerpt}</div>}
      <div className="review-row-actions">
        <button
          type="button"
          className="review-row-act"
          disabled={!canEdit}
          onClick={(ev) => {
            ev.stopPropagation()
            onAccept(e.rev)
          }}
        >
          ✓ {t('appRevAccept')}
        </button>
        <button
          type="button"
          className="review-row-act"
          disabled={!canEdit}
          onClick={(ev) => {
            ev.stopPropagation()
            onReject(e.rev)
          }}
        >
          ✕ {t('appRevReject')}
        </button>
      </div>
    </div>
  )
}

/**
 * Review → Reviewing Pane, the outermost right-edge panel: Word's summary
 * lines, the per-author ink legend, then every tracked change in document
 * order; a click selects it, Accept / Reject act on it.
 */
export function ReviewingPane({
  entries,
  hiddenCount,
  authors,
  openComments,
  activeIndex,
  canEdit,
  onSelect,
  onAccept,
  onReject,
  onClose,
}: {
  entries: ReviewEntry[]
  /** changes the Show Markup filters keep out of the list */
  hiddenCount: number
  /** legend: authors in first-appearance order (ink = index mod 8) */
  authors: Array<{ name: string; inkIndex: number }>
  /** open comment threads (Word's "Comments:" line) */
  openComments: number
  activeIndex: number
  canEdit: boolean
  onSelect: (rev: RevisionRange) => void
  onAccept: (rev: RevisionRange) => void
  onReject: (rev: RevisionRange) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const listRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (activeIndex < 0) return
    listRef.current
      ?.querySelector<HTMLElement>(`[data-rev-index="${activeIndex}"]`)
      ?.scrollIntoView({ block: 'nearest' })
  }, [activeIndex])

  const revs = entries.map((e) => e.rev)
  const sum = summarizeRevisions(revs)
  const lines: Array<[StringKey, number]> = [
    ['appReviewSumInsertions', sum.insertions],
    ['appReviewSumDeletions', sum.deletions],
    ['appReviewSumMoves', sum.moves],
    ['appReviewSumFormatting', sum.formatting],
    ['appReviewSumComments', openComments],
  ]

  return (
    <aside className="comments-pane review-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">
          <IconTrackChanges size={14} /> {t('appReviewPaneTitle')}
          <span className="review-pane-count">{revisionCount(revs)}</span>
        </span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      <div className="review-pane-meta">
        <dl className="review-pane-summary">
          {lines.map(([key, n]) => (
            <div key={key} className="review-pane-sum-line">
              <dt>{t(key)}:</dt>
              <dd>{n}</dd>
            </div>
          ))}
        </dl>
        {hiddenCount > 0 && (
          <div className="review-pane-hidden">{t('appReviewPaneHidden', { n: hiddenCount })}</div>
        )}
        {authors.length > 0 && (
          <div className="review-pane-legend">
            {authors.map((a) => (
              <span key={a.name} className="review-pane-author">
                <span className="review-swatch" style={{ background: inkVar(a.inkIndex) }} />
                {a.name || t('editorUnknownAuthor')}
              </span>
            ))}
          </div>
        )}
      </div>
      <div className="review-pane-list" ref={listRef}>
        {entries.length === 0 && <div className="comments-empty">{t('appReviewPaneEmpty')}</div>}
        {entries.map((e) => (
          <ReviewRow
            key={e.rev.story?.id ?? `${e.rev.from}-${e.rev.kind}-${e.rev.author}`}
            entry={e}
            active={e.index === activeIndex}
            canEdit={canEdit}
            onSelect={onSelect}
            onAccept={onAccept}
            onReject={onReject}
          />
        ))}
      </div>
    </aside>
  )
}
