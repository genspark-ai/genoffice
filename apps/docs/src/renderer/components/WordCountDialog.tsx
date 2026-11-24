/**
 * Word's Word Count dialog (Tools ▸ Word Count). Read-only counts; the
 * numbers are computed in App.tsx where the paginated DOM lives.
 */
import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** fields of Word's Word Count dialog */
export interface DocStats {
  pages: number
  words: number
  asianChars: number
  nonAsianWords: number
  charsNoSpace: number
  charsWithSpace: number
  paragraphs: number
  lines: number
}

const ROWS = [
  ['appStatPages', 'pages'],
  ['appStatWords', 'words'],
  ['appStatCharsNoSpaces', 'charsNoSpace'],
  ['appStatCharsWithSpaces', 'charsWithSpace'],
  ['appStatNonAsianWords', 'nonAsianWords'],
  ['appStatAsianChars', 'asianChars'],
  ['appStatParagraphs', 'paragraphs'],
  ['appStatLines', 'lines'],
] as const

const INCLUDE_NOTES_KEY = 'docs-word-count-include-notes'

export function WordCountDialog({
  stats,
  notes,
  onClose,
}: {
  /** the document body */
  stats: DocStats
  /** what the footnotes and endnotes add to each figure */
  notes: DocStats
  onClose: () => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [includeNotes, setIncludeNotes] = useState(
    () => localStorage.getItem(INCLUDE_NOTES_KEY) !== '0',
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal word-count-dialog" role="dialog" aria-label={t('appWordCountTitle')}>
        <h2>{t('appWordCountTitle')}</h2>
        <table className="stats-table">
          <tbody>
            {ROWS.map(([labelKey, field]) => (
              <tr key={field} data-stat={field}>
                <td>{t(labelKey)}</td>
                <td>{stats[field] + (includeNotes ? notes[field] : 0)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <label className="stats-check">
          <input
            type="checkbox"
            checked={includeNotes}
            onChange={(e) => {
              setIncludeNotes(e.target.checked)
              localStorage.setItem(INCLUDE_NOTES_KEY, e.target.checked ? '1' : '0')
            }}
          />
          {t('appStatIncludeNotes')}
        </label>
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
