import { useState } from 'react'

import { useI18n } from './i18n/locale'

import type { GoToNameEntry } from './goto'

/// Excel's Go To (⌘G / Ctrl+G), minimal: a reference input plus the
/// workbook's defined names. Go (or Enter, or double-clicking a name) hands
/// the reference to App, which resolves and selects it; failures show
/// inline and keep the dialog open.
export function GoToDialog({
  names,
  onGo,
  onClose,
}: {
  readonly names: readonly GoToNameEntry[]
  /// Returns an error message, or null on success.
  readonly onGo: (ref: string) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [reference, setReference] = useState('')
  const [error, setError] = useState<string | null>(null)

  const go = (ref: string): void => {
    const failure = onGo(ref)
    setError(failure)
    if (failure === null) onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgGoToTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgGoToTitle')}</header>
        <section className="dialog-body">
          <div className="dialog-grid">
            <label className="dialog-span">
              {t('dlgGoToRef')}
              <input
                autoFocus
                value={reference}
                placeholder={t('dlgGoToRefPlaceholder')}
                onChange={(event) => setReference(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' && reference.trim() !== '') go(reference)
                }}
              />
            </label>
          </div>
          {names.length > 0 && (
            <ul className="goto-name-list">
              {names.map((entry) => (
                <li key={entry.name}>
                  <button
                    type="button"
                    data-tip={t('dlgGoToNameHint', { ref: entry.ref })}
                    onClick={() => setReference(entry.name)}
                    onDoubleClick={() => go(entry.name)}
                  >
                    <strong>{entry.name}</strong>
                    <span>{entry.ref}</span>
                  </button>
                </li>
              ))}
            </ul>
          )}
          <p className="dialog-note">{t('dlgGoToNote')}</p>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            disabled={reference.trim() === ''}
            onClick={() => go(reference)}
          >
            {t('dlgGoToGo')}
          </button>
        </div>
      </div>
    </div>
  )
}
