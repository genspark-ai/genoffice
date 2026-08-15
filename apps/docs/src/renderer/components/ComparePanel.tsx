import { Fragment } from 'react'
import type { CompareEntry } from '../editor/compare'
import { summarize } from '../editor/compare'
import { useI18n } from '../i18n/locale'
import { IconCompare } from './icons'

/**
 * Review → Compare result pane: paragraph-level differences between the open
 * document and a chosen second file. Unchanged stretches collapse to a count.
 */
export function ComparePanel({
  otherName,
  entries,
  onClose,
}: {
  otherName: string
  entries: CompareEntry[]
  onClose: () => void
}) {
  const { t } = useI18n()
  const { added, removed, changed } = summarize(entries)
  const rows: Array<{ key: number; entry?: CompareEntry; sameCount?: number }> = []
  let sameRun = 0
  entries.forEach((entry, i) => {
    if (entry.kind === 'same') {
      sameRun++
      return
    }
    if (sameRun > 0) {
      rows.push({ key: rows.length, sameCount: sameRun })
      sameRun = 0
    }
    rows.push({ key: rows.length, entry })
    void i
  })
  if (sameRun > 0) rows.push({ key: rows.length, sameCount: sameRun })

  return (
    <aside className="comments-pane compare-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">
          <IconCompare size={14} /> {t('appCompareResult')}
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
      <div className="compare-meta">
        {t('appCompareWithPrefix')}
        <b>{otherName}</b>
        {t('appCompareStats', { added, removed, changed })}
      </div>
      <div className="comments-pane-list">
        {added + removed + changed === 0 && (
          <div className="comments-empty">{t('appCompareIdentical')}</div>
        )}
        {rows.map((row) => (
          <Fragment key={row.key}>
            {row.sameCount !== undefined && (
              <div className="compare-same">{t('appCompareSameRun', { n: row.sameCount })}</div>
            )}
            {row.entry && (
              <div className={`compare-card compare-${row.entry.kind}`}>
                {row.entry.kind !== 'added' && (
                  <div className="compare-left">{row.entry.left || t('appEmptyParagraph')}</div>
                )}
                {row.entry.kind !== 'removed' && (
                  <div className="compare-right">{row.entry.right || t('appEmptyParagraph')}</div>
                )}
              </div>
            )}
          </Fragment>
        ))}
      </div>
    </aside>
  )
}
