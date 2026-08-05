import { useEffect, useRef, useState } from 'react'
import { useI18n } from './locale'
import { AiProvidersPage } from './AiProvidersPage'
import './settings.css'

interface SettingsDialogProps {
  onClose: () => void
}

export function SettingsDialog({ onClose }: SettingsDialogProps) {
  const { t } = useI18n()
  const [page, setPage] = useState<'general' | 'ai-providers'>('ai-providers')
  const dialogRef = useRef<HTMLDivElement>(null)
  const previousFocus = useRef<HTMLElement | null>(null)

  useEffect(() => {
    previousFocus.current =
      document.activeElement instanceof HTMLElement ? document.activeElement : null
    dialogRef.current?.focus()
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault()
        onClose()
        return
      }
      if (event.key !== 'Tab') return
      const dialog = dialogRef.current
      if (!dialog) return
      const focusables = Array.from(
        dialog.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [href], [tabindex]:not([tabindex="-1"])',
        ),
      )
      if (focusables.length === 0) {
        event.preventDefault()
        return
      }
      const first = focusables[0]
      const last = focusables[focusables.length - 1]
      const active = document.activeElement
      if (event.shiftKey && active === first) {
        event.preventDefault()
        last.focus()
      } else if (!event.shiftKey && active === last) {
        event.preventDefault()
        first.focus()
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('keydown', onKeyDown)
      previousFocus.current?.focus()
    }
  }, [onClose])

  return (
    <div
      className="settings-overlay"
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose()
      }}
    >
      <div
        className="settings-dialog"
        ref={dialogRef}
        tabIndex={-1}
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-dialog-title"
      >
        <header className="settings-dialog-header">
          <h1 id="settings-dialog-title">{t('settingsTitle')}</h1>
          <button
            className="settings-close"
            type="button"
            onClick={onClose}
            aria-label={t('closeSettings')}
          >
            <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
              <path
                d="m4 4 10 10M14 4 4 14"
                stroke="currentColor"
                strokeWidth="1.6"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </header>
        <div className="settings-dialog-body">
          <nav className="settings-nav" aria-label={t('settingsNavigationLabel')}>
            <button
              type="button"
              className={`settings-nav-item${page === 'general' ? ' active' : ''}`}
              aria-current={page === 'general' ? 'page' : undefined}
              onClick={() => setPage('general')}
            >
              <span className="settings-nav-icon" aria-hidden="true">
                ⚙
              </span>
              {t('generalSettings')}
            </button>
            <button
              type="button"
              className={`settings-nav-item${page === 'ai-providers' ? ' active' : ''}`}
              aria-current={page === 'ai-providers' ? 'page' : undefined}
              onClick={() => setPage('ai-providers')}
            >
              <span className="settings-nav-icon" aria-hidden="true">
                ✦
              </span>
              {t('aiProvidersTitle')}
            </button>
          </nav>
          <main className="settings-content">
            {page === 'ai-providers' ? (
              <AiProvidersPage />
            ) : (
              <section className="settings-general">
                <h2>{t('generalSettings')}</h2>
                <p>{t('generalSettingsDescription')}</p>
              </section>
            )}
          </main>
        </div>
      </div>
    </div>
  )
}
