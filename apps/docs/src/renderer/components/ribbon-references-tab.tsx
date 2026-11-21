import { useState } from 'react'
import {
  CITATION_STYLES,
  citationText,
  type CitationStyleId,
  type SourceInfo,
} from '@genoffice/docx-engine'
import { Dropdown } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import type { ActionCommandId } from '../../shared/commands'
import { RbCommand, useRunCommand } from './ribbon-command'
import {
  IconBook,
  IconCaption,
  IconCaret,
  IconCitation,
  IconEndnote,
  IconFootnote,
  IconIndex,
  IconRefresh,
  IconToc,
  IconAddText,
  IconNextFootnote,
  IconShowNotes,
  IconCitationsPane,
  IconMarkCitation,
  IconToa,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'

const SOURCE_TYPES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'Book', nameKey: 'ribbonSourceBook' },
  { key: 'JournalArticle', nameKey: 'ribbonSourceJournal' },
  { key: 'InternetSite', nameKey: 'ribbonSourceWebsite' },
  { key: 'Report', nameKey: 'ribbonSourceReport' },
  { key: 'Misc', nameKey: 'ribbonSourceMisc' },
]

/** Create Source dialog for the citation manager */
function SourceModal({
  sources,
  onAdd,
  onClose,
}: {
  sources: SourceInfo[]
  onAdd: (source: SourceInfo) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [type, setType] = useState('Book')
  const [author, setAuthor] = useState('')
  const [title, setTitle] = useState('')
  const [year, setYear] = useState('')
  const [publisher, setPublisher] = useState('')
  const [url, setUrl] = useState('')

  const add = () => {
    if (!title.trim()) return
    // unique tag derived from author/title initials + year
    const base = (author.trim() || title.trim()).replace(/[\s,，、]+/g, '').slice(0, 8) || 'Src'
    let tag = `${base}${year.trim()}`
    let n = 1
    while (sources.some((s) => s.tag === tag)) tag = `${base}${year.trim()}_${++n}`
    onAdd({
      tag,
      type,
      author: author.trim(),
      title: title.trim(),
      year: year.trim(),
      ...(publisher.trim() ? { publisher: publisher.trim() } : {}),
      ...(url.trim() ? { url: url.trim() } : {}),
    })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonSourceCreateTitle')}</h2>
        <label>
          {t('ribbonSourceType')}
          <Dropdown
            value={type}
            ariaLabel={t('ribbonSourceType')}
            options={SOURCE_TYPES.map((s) => ({ value: s.key, label: t(s.nameKey) }))}
            onPick={setType}
          />
        </label>
        <label>
          {t('ribbonSourceAuthor')}
          <input
            value={author}
            onChange={(e) => setAuthor(e.target.value)}
            placeholder={t('ribbonSourceAuthorPh')}
          />
        </label>
        <label>
          {t('ribbonSourceTitle')}
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder={t('ribbonSourceTitlePh')}
          />
        </label>
        <label>
          {t('ribbonSourceYear')}
          <input value={year} onChange={(e) => setYear(e.target.value)} placeholder="2026" />
        </label>
        <label>
          {type === 'JournalArticle'
            ? t('ribbonSourceJournalName')
            : type === 'InternetSite'
              ? t('ribbonSourceSiteName')
              : t('ribbonSourcePublisher')}
          <input value={publisher} onChange={(e) => setPublisher(e.target.value)} />
        </label>
        {type === 'InternetSite' && (
          <label>
            URL
            <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…" />
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!title.trim()} onClick={add}>
            {t('ribbonAdd')}
          </button>
        </div>
      </div>
    </div>
  )
}

interface ReferencesTabProps extends TabProps {
  sources: SourceInfo[]
  citationStyle: CitationStyleId | null
  onAddSource: (source: SourceInfo) => void
}

export function ReferencesTab({
  hasDoc,
  dropdown,
  setDropdown,
  sources,
  citationStyle,
  onAddSource,
}: ReferencesTabProps) {
  const { t } = useI18n()
  const [sourceOpen, setSourceOpen] = useState(false)
  const runCommand = useRunCommand()
  /** a pick from a menu: run the command, close the menu */
  const pick = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(() => null)
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonTocTip')}
            onClick={() => runCommand('insert-toc')}
          >
            <span className="rb-big-icon">
              <IconToc size={BIG} />
            </span>
            <span>{t('ribbonToc')}</span>
          </button>
          <div className="rb-col rb-pair">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonAddText')}
                onClick={() => toggleDropdown(setDropdown, 'add-text')}
              >
                <IconAddText size={14} />
                <span className="rb-small-label">{t('ribbonAddText')}</span>
                <IconCaret />
              </button>
              {dropdown === 'add-text' && (
                <div data-rb-panel="" className="layout-menu">
                  <button onClick={() => pick('toc-add-text', '0')}>
                    {t('ribbonAddTextNone')}
                  </button>
                  {[1, 2, 3].map((n) => (
                    <button key={n} onClick={() => pick('toc-add-text', String(n))}>
                      {t('appParaOutlineLevelN', { n })}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonTocUpdateTip')}
              onClick={() => runCommand('update-toc')}
            >
              <IconRefresh size={14} />
              <span className="rb-small-label">{t('ribbonTocUpdate')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonToc')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonFootnoteTip')}
            onClick={() => runCommand('insert-footnote')}
          >
            <span className="rb-big-icon">
              <IconFootnote size={BIG} />
            </span>
            <span>{t('ribbonFootnote')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonEndnoteTip')}
            onClick={() => runCommand('insert-endnote')}
          >
            <span className="rb-big-icon">
              <IconEndnote size={BIG} />
            </span>
            <span>{t('ribbonEndnote')}</span>
          </button>
          <div className="rb-col rb-pair">
            <RbCommand cmd="footnote-next" look="small">
              <IconNextFootnote size={14} />
            </RbCommand>
            <RbCommand cmd="view-footnotes" look="small">
              <IconShowNotes size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupFootnotes')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonCitationTip')}
              onClick={() => toggleDropdown(setDropdown, 'citation')}
            >
              <span className="rb-big-icon">
                <IconCitation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonCitation')}</span>
            </button>
            {dropdown === 'citation' && (
              <div data-rb-panel="" className="layout-menu">
                {sources.map((s) => (
                  <button
                    key={s.tag}
                    data-tip={s.title}
                    onClick={() => pick('insert-citation', s.tag)}
                  >
                    {citationText(s)} {s.title.slice(0, 12)}
                  </button>
                ))}
                <button
                  onClick={() => {
                    setSourceOpen(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonAddNewSource')}
                </button>
              </div>
            )}
          </div>
          <RbCommand cmd="citations-pane" look="big">
            <IconCitationsPane size={BIG} />
          </RbCommand>
          <div className="rb-col rb-pair">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonCitationStyle')}
                aria-haspopup="menu"
                onClick={() => toggleDropdown(setDropdown, 'citation-style')}
              >
                <IconCitation size={14} />
                <span className="rb-small-label">{t('ribbonCitationStyle')}</span>
                <small>{CITATION_STYLES.find((s) => s.id === citationStyle)?.name ?? ''}</small>
                <IconCaret />
              </button>
              {dropdown === 'citation-style' && (
                <div data-rb-panel="" className="layout-menu" role="menu">
                  {CITATION_STYLES.map((style) => (
                    <button
                      key={style.id}
                      role="menuitemradio"
                      aria-checked={style.id === citationStyle}
                      onClick={() => pick('citation-style', style.id)}
                    >
                      {style.name}
                      {style.id === citationStyle ? ' ✓' : ''}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <button
              className="rb-small"
              disabled={!hasDoc}
              data-tip={t('ribbonBibliographyTip')}
              onClick={() => pick('insert-bibliography')}
            >
              <IconBook size={14} />
              <span className="rb-small-label">{t('ribbonBibliography')}</span>
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupCitationsBibliography')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonCaptionTip')}
            onClick={() => runCommand('insert-caption')}
          >
            <span className="rb-big-icon">
              <IconCaption size={BIG} />
            </span>
            <span>{t('ribbonInsertCaption')}</span>
          </button>
          <RbCommand cmd="insert-table-of-figures" look="big">
            <IconToc size={BIG} />
          </RbCommand>
          <div className="rb-col rb-pair">
            <RbCommand cmd="update-table-of-figures" look="small">
              <IconRefresh size={14} />
            </RbCommand>
            <RbCommand cmd="insert-cross-reference" look="small">
              <IconRefresh size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupCaptions')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <RbCommand cmd="mark-index-entry" look="big">
            <IconMarkCitation size={BIG} />
          </RbCommand>
          <div className="rb-col rb-pair">
            <RbCommand cmd="insert-index" look="small">
              <IconIndex size={14} />
            </RbCommand>
            <RbCommand cmd="update-index" look="small">
              <IconRefresh size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonIndex')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <RbCommand cmd="toa-mark-citation" look="big">
            <IconMarkCitation size={BIG} />
          </RbCommand>
          <div className="rb-col rb-pair">
            <RbCommand cmd="toa-insert" look="small">
              <IconToa size={14} />
            </RbCommand>
            <RbCommand cmd="toa-update" look="small">
              <IconRefresh size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonItAuthorities')}</div>
      </div>

      {sourceOpen && (
        <SourceModal sources={sources} onAdd={onAddSource} onClose={() => setSourceOpen(false)} />
      )}
    </>
  )
}

/* ================= Review ================= */
