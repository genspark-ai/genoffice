import { useState } from 'react'
import type { TocEntry } from '../editor/index-entry'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const TABLE_IDS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i))
const LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9]

/** Word's Mark Table of Contents Entry: the entry's text, the table it is for and its level */
export function MarkTocEntryDialog({
  selected,
  onMark,
  onClose,
}: {
  /** the selected text, which the entry starts as */
  selected: string
  onMark: (entry: TocEntry) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [entry, setEntry] = useState<TocEntry>({ text: selected, id: 'C', level: 1 })
  const set = (patch: Partial<TocEntry>) => setEntry((prev) => ({ ...prev, ...patch }))
  const ready = entry.text.trim() !== ''
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog mark-entry-dialog mark-toc-entry-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!ready) return
          onClose()
          onMark(entry)
        }}
      >
        <h2>{t('appScMarkTocEntry')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('appTocEntryText')}
            <input
              type="text"
              name="entry"
              value={entry.text}
              autoFocus
              spellCheck={false}
              onChange={(e) => set({ text: e.target.value })}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonItTableIdentifier')}
            <select name="table" value={entry.id} onChange={(e) => set({ id: e.target.value })}>
              {TABLE_IDS.map((id) => (
                <option key={id} value={id}>
                  {id}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('ribbonListLevel')}
            <select
              name="level"
              value={entry.level}
              onChange={(e) => set({ level: Number(e.target.value) })}
            >
              {LEVELS.map((level) => (
                <option key={level} value={level}>
                  {level}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
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
