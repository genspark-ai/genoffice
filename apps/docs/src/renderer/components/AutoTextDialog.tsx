import { useState } from 'react'
import type { AutoTextEntry } from '../editor/auto-text'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * Word's Insert > AutoText: `new` names the selection to keep it, `insert`
 * lists what is kept, to put one in or delete it.
 */
export function AutoTextDialog({
  mode,
  suggestedName,
  entries,
  onSave,
  onInsert,
  onDelete,
  onClose,
}: {
  mode: 'new' | 'insert'
  /** the start of the selection: what a new entry is called unless renamed */
  suggestedName: string
  entries: AutoTextEntry[]
  onSave: (name: string) => void
  onInsert: (entry: AutoTextEntry) => void
  onDelete: (name: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [name, setName] = useState(suggestedName)
  // by place in the list: the document may hold an entry under a name this computer has too
  const [picked, setPicked] = useState(0)
  const entry = entries[picked] ?? entries[0]
  const ready = mode === 'new' ? name.trim() !== '' : entry !== undefined
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog auto-text-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!ready) return
          onClose()
          if (mode === 'new') onSave(name.trim())
          else if (entry) onInsert(entry)
        }}
      >
        <h2>{t(mode === 'new' ? 'ribbonAutoTextNew' : 'ribbonAutoText')}</h2>
        {mode === 'new' ? (
          <div className="modal-row margin-row">
            <label>
              {t('ribbonStyleDialogName')}
              <input
                type="text"
                value={name}
                autoFocus
                spellCheck={false}
                onChange={(e) => setName(e.target.value)}
              />
            </label>
          </div>
        ) : entries.length === 0 ? (
          <p className="index-tables-hint">{t('ribbonAutoTextEmpty')}</p>
        ) : (
          <>
            <select
              className="auto-text-list"
              size={8}
              autoFocus
              aria-label={t('ribbonAutoText')}
              value={entry ? entries.indexOf(entry) : 0}
              onChange={(e) => setPicked(Number(e.target.value))}
              onDoubleClick={(e) => e.currentTarget.form?.requestSubmit()}
            >
              {entries.map((e, i) => (
                <option key={i} value={i}>
                  {e.name}
                </option>
              ))}
            </select>
            <p className="index-tables-hint auto-text-preview">{entry?.text.slice(0, 160)}</p>
          </>
        )}
        <div className="modal-actions">
          {mode === 'insert' && entry ? (
            <button
              type="button"
              disabled={entry.from === 'document'}
              onClick={() => onDelete(entry.name)}
            >
              {t('ribbonGroupDelete')}
            </button>
          ) : (
            <span />
          )}
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!ready}>
              {t(mode === 'new' ? 'ribbonOk' : 'ribbonInsert')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
