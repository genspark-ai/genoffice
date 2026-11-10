/**
 * Word's Word Count dialog (Review ▸ Word Count, ⇧⌘G). Read-only counts; the
 * numbers are computed in App.tsx where the paginated DOM lives.
 */
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
  ['appStatAsianChars', 'asianChars'],
  ['appStatNonAsianWords', 'nonAsianWords'],
  ['appStatCharsNoSpaces', 'charsNoSpace'],
  ['appStatCharsWithSpaces', 'charsWithSpace'],
  ['appStatParagraphs', 'paragraphs'],
  ['appStatLines', 'lines'],
] as const

export function WordCountDialog({ stats, onClose }: { stats: DocStats; onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" role="dialog" aria-label={t('appWordCountTitle')}>
        <h2>{t('appWordCountTitle')}</h2>
        <table className="stats-table">
          <tbody>
            {ROWS.map(([labelKey, field]) => (
              <tr key={field}>
                <td>{t(labelKey)}</td>
                <td>{stats[field]}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
