import { useState } from 'react'
import type { IndexEntry } from '../editor/index-entry'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** Word's Mark Index Entry: the entry, where the index sends the reader, how its page is set */
export function MarkIndexEntryDialog({
  selected,
  bookmarks,
  onMark,
  onClose,
}: {
  /** the selected text: what the main entry starts as, and what Mark All looks for */
  selected: string
  bookmarks: string[]
  onMark: (entry: IndexEntry, all: boolean) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [entry, setEntry] = useState<IndexEntry>(() => ({
    main: selected,
    sub: '',
    target: 'page',
    cross: `${t('ribbonXeCrossRefSee')} `,
    bookmark: bookmarks[0] ?? '',
    bold: false,
    italic: false,
  }))
  const set = (patch: Partial<IndexEntry>) => setEntry((prev) => ({ ...prev, ...patch }))
  const ready =
    entry.main.trim() !== '' &&
    (entry.target !== 'range' || entry.bookmark !== '') &&
    (entry.target !== 'cross' || entry.cross.trim() !== '')
  const mark = (all: boolean) => {
    if (!ready) return
    onClose()
    onMark(entry, all)
  }
  const target = (value: IndexEntry['target'], label: string) => (
    <label className="tco-check">
      <input
        type="radio"
        name="xe-target"
        checked={entry.target === value}
        onChange={() => set({ target: value })}
      />
      {label}
    </label>
  )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog mark-entry-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          mark(false)
        }}
      >
        <h2>{t('ribbonMarkEntryTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonXeMainEntry')}
            <input
              type="text"
              name="main"
              value={entry.main}
              autoFocus
              spellCheck={false}
              onChange={(e) => set({ main: e.target.value })}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonXeSubentry')}
            <input
              type="text"
              name="sub"
              value={entry.sub}
              spellCheck={false}
              onChange={(e) => set({ sub: e.target.value })}
            />
          </label>
        </div>
        <div className="xe-target">
          {target('cross', t('ribbonXeCrossRef'))}
          <input
            type="text"
            name="cross"
            aria-label={t('ribbonXeCrossRef')}
            value={entry.cross}
            disabled={entry.target !== 'cross'}
            spellCheck={false}
            onChange={(e) => set({ cross: e.target.value })}
          />
        </div>
        <div className="xe-target">{target('page', t('ribbonXeCurrentPage'))}</div>
        <div className="xe-target">
          {target('range', t('ribbonXePageRange'))}
          <select
            name="bookmark"
            aria-label={t('ribbonBookmark')}
            value={entry.bookmark}
            disabled={entry.target !== 'range' || bookmarks.length === 0}
            onChange={(e) => set({ bookmark: e.target.value })}
          >
            {bookmarks.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>
        <fieldset>
          <legend>{t('ribbonXePageFormat')}</legend>
          <label className="tco-check">
            <input
              type="checkbox"
              name="bold"
              checked={entry.bold}
              onChange={(e) => set({ bold: e.target.checked })}
            />
            {t('ribbonMarkerBold')}
          </label>
          <label className="tco-check">
            <input
              type="checkbox"
              name="italic"
              checked={entry.italic}
              onChange={(e) => set({ italic: e.target.checked })}
            />
            {t('ribbonMarkerItalic')}
          </label>
        </fieldset>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button
              type="button"
              disabled={!ready || entry.target !== 'page' || !selected}
              onClick={() => mark(true)}
            >
              {t('ribbonXeMarkAll')}
            </button>
            <button type="submit" className="primary" disabled={!ready}>
              {t('ribbonXeMark')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
