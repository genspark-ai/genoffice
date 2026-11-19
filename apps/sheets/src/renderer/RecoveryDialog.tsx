import { useEffect } from 'react'

import { useI18n } from './i18n/locale'

import type { RecoveryPromptPayload } from '../shared/desktop-api'

/// Styled replacement for the native "Recovered version found" message box
/// raised by main while opening a workbook whose autosaved recovery copy is
/// newer than the file. Strings arrive pre-localized in the payload; the
/// dialog adds what the native box could not show — the file name and when
/// the unsaved work was last autosaved.
export function RecoveryDialog({
  prompt,
  onChoose,
}: {
  readonly prompt: RecoveryPromptPayload
  /// true = restore the autosaved copy, false = discard it and open the file.
  readonly onChoose: (restore: boolean) => void
}): React.JSX.Element {
  const { dateLocale } = useI18n()

  // Esc mirrors the native dialog's cancelId (Discard). No backdrop-click
  // dismissal: the choice deletes or resurrects unsaved work, so it must be
  // explicit.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') onChoose(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onChoose])

  const savedAt = new Date(prompt.savedAtMs).toLocaleString(dateLocale, {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })

  return (
    <div className="dialog-backdrop">
      <div className="format-cells-dialog recovery-dialog" role="dialog" aria-label={prompt.title}>
        <section className="dialog-body">
          <div className="recovery-heading">
            <span className="recovery-icon" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                <path d="M3 12a9 9 0 1 0 2.6-6.3" strokeLinecap="round" />
                <path d="M3 4.5V9h4.5" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M12 7.5V12l3.5 2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </span>
            <h2>{prompt.title}</h2>
          </div>
          <p className="recovery-body">{prompt.body}</p>
          <div className="recovery-file">
            <span className="recovery-file-name" title={prompt.fileName}>
              {prompt.fileName}
            </span>
            <span className="recovery-file-time">{savedAt}</span>
          </div>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={() => onChoose(false)}>
            {prompt.discardLabel}
          </button>
          <button className="primary-action" autoFocus onClick={() => onChoose(true)}>
            {prompt.restoreLabel}
          </button>
        </div>
      </div>
    </div>
  )
}
