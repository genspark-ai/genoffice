import { citationText, type SourceInfo } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'

/** Word's Citations pane: the document's sources; a click puts a citation of one at the caret */
export function CitationsPane({
  sources,
  canInsert,
  onInsert,
  onClose,
}: {
  sources: SourceInfo[]
  /** the document can be typed into */
  canInsert: boolean
  onInsert: (tag: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  return (
    <aside className="comments-pane citations-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">{t('ribbonCitationsPane')}</span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      <div className="comments-pane-list">
        {sources.length === 0 && <div className="comments-empty">{t('ribbonNoSources')}</div>}
        {sources.map((source) => (
          <button
            key={source.tag}
            type="button"
            className="selection-pane-row"
            disabled={!canInsert}
            onClick={() => onInsert(source.tag)}
          >
            <span>{source.title || source.tag}</span>
            <span className="selection-pane-text">{citationText(source)}</span>
          </button>
        ))}
      </div>
    </aside>
  )
}
