import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export function EncryptedWorkbookPasswordDialog({
  error,
  onSubmit,
  onClose,
}: {
  readonly error: string | null
  readonly onSubmit: (password: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [password, setPassword] = useState('')
  const modal = useModalDialog(onClose)
  const submit = (): void => {
    if (password.length > 0) onSubmit(password)
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog encrypted-workbook-password-dialog"
        role="dialog"
        aria-labelledby="encrypted-workbook-password-title"
        aria-describedby="encrypted-workbook-password-copy"
        {...modal}
        onClick={(event) => event.stopPropagation()}
      >
        <header id="encrypted-workbook-password-title">{t('appEncryptedOpenTitle')}</header>
        <section className="dialog-body">
          <p className="encrypted-workbook-password-copy" id="encrypted-workbook-password-copy">
            {t('appEncryptedOpenCopy')}
          </p>
          <label>
            {t('appEncryptedPasswordLabel')}
            <input
              autoFocus
              autoComplete="current-password"
              type="password"
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') submit()
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
          <button className="secondary" type="button" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" type="button" disabled={!password} onClick={submit}>
            {t('appEncryptedOpen')}
          </button>
        </div>
      </div>
    </div>
  )
}
