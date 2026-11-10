import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export function ChangeWorkbookPasswordDialog({
  onSubmit,
  onClose,
}: {
  readonly onSubmit: (password: string) => Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const modal = useModalDialog(onClose)

  const submit = async (): Promise<void> => {
    if (busy || !password || !confirmation) return
    if (password !== confirmation) {
      setError(t('appPasswordsMismatch'))
      return
    }
    setBusy(true)
    setError(null)
    try {
      const failure = await onSubmit(password)
      if (failure) setError(failure)
      else onClose()
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : t('appPasswordUpdateFailed'))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="dialog-backdrop" onClick={busy ? undefined : onClose}>
      <div
        className="format-cells-dialog encrypted-workbook-password-dialog change-workbook-password-dialog"
        role="dialog"
        aria-labelledby="change-workbook-password-title"
        aria-describedby="change-workbook-password-copy"
        {...modal}
        onClick={(event) => event.stopPropagation()}
      >
        <header id="change-workbook-password-title">{t('appChangePasswordTitle')}</header>
        <section className="dialog-body">
          <p className="encrypted-workbook-password-copy" id="change-workbook-password-copy">
            {t('appChangePasswordCopy')}
          </p>
          <label>
            {t('appNewPasswordLabel')}
            <input
              autoFocus
              autoComplete="new-password"
              maxLength={1024}
              type="password"
              value={password}
              disabled={busy}
              onChange={(event) => setPassword(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') void submit()
              }}
            />
          </label>
          <label>
            {t('appConfirmPasswordLabel')}
            <input
              autoComplete="new-password"
              maxLength={1024}
              type="password"
              value={confirmation}
              disabled={busy}
              onChange={(event) => setConfirmation(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') void submit()
              }}
            />
          </label>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" type="button" disabled={busy} onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            type="button"
            disabled={busy || !password || !confirmation}
            onClick={() => void submit()}
          >
            {busy ? t('appPasswordUpdating') : t('appUpdatePassword')}
          </button>
        </div>
      </div>
    </div>
  )
}
