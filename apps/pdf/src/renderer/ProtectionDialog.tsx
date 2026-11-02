import { useState } from 'react'
import type { TFunc } from './i18n/locale'
import type { PdfPermissions } from '../shared/ipc'

/**
 * Review → Protection: set, change or lift the password the document is saved
 * under, and the limits a reader has once it is open. Applying only updates the
 * session — the next save writes the file with this configuration.
 */
export function ProtectionDialog({
  t,
  hasPassword,
  permissions,
  onApply,
  onRemove,
  onClose,
}: {
  t: TFunc
  /** whether the document on disk is currently password-protected */
  hasPassword: boolean
  permissions: PdfPermissions
  onApply: (result: {
    userPassword: string
    ownerPassword: string
    permissions: PdfPermissions
  }) => void
  onRemove: () => void
  onClose: () => void
}): React.JSX.Element {
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [owner, setOwner] = useState('')
  const [perms, setPerms] = useState<PdfPermissions>(permissions)
  const [error, setError] = useState<string | null>(null)

  const apply = (): void => {
    if (password.length === 0 || password !== confirm) {
      setError(t('protectMismatch'))
      return
    }
    onApply({
      userPassword: password,
      ownerPassword: owner.length > 0 ? owner : password,
      permissions: perms,
    })
  }

  const toggle = (key: keyof PdfPermissions): void => {
    setPerms((current) => ({ ...current, [key]: !current[key] }))
  }

  const checkbox = (key: keyof PdfPermissions, label: string): React.JSX.Element => (
    <label className="pdf-perm-row">
      <input type="checkbox" checked={perms[key]} onChange={() => toggle(key)} />
      <span>{label}</span>
    </label>
  )

  return (
    <div className="pdf-modal-mask" onClick={onClose}>
      <div
        className="pdf-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('protectTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{t('protectTitle')}</div>
        <div className="pdf-modal-hint">
          {hasPassword ? t('protectChangeHint') : t('protectSetHint')}
        </div>

        <div className="pdf-field">
          <label className="pdf-perm-row">
            <span>{t('protectPasswordLabel')}</span>
            <input
              type="password"
              className="pdf-modal-input"
              value={password}
              onChange={(e) => {
                setPassword(e.target.value)
                setError(null)
              }}
              autoFocus
            />
          </label>
          <label className="pdf-perm-row">
            <span>{t('protectConfirmLabel')}</span>
            <input
              type="password"
              className="pdf-modal-input"
              value={confirm}
              onChange={(e) => {
                setConfirm(e.target.value)
                setError(null)
              }}
            />
          </label>
          <label className="pdf-perm-row">
            <span>{t('protectOwnerLabel')}</span>
            <input
              type="password"
              className="pdf-modal-input"
              value={owner}
              onChange={(e) => setOwner(e.target.value)}
            />
          </label>
        </div>

        <div className="pdf-field">
          <div className="pdf-modal-hint">{t('protectPermsLabel')}</div>
          {checkbox('printing', t('permPrint'))}
          {checkbox('modifying', t('permModify'))}
          {checkbox('annotating', t('permAnnotate'))}
          {checkbox('copying', t('permCopy'))}
        </div>

        {error && (
          <div className="pdf-field">
            <div className="pdf-cert-error" role="alert">
              {error}
            </div>
          </div>
        )}

        <div className="pdf-modal-actions">
          {hasPassword && (
            <button
              className="pdf-modal-btn"
              onClick={() => {
                onRemove()
              }}
            >
              {t('protectRemove')}
            </button>
          )}
          <button className="pdf-modal-btn" onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" onClick={apply}>
            {t('protectApply')}
          </button>
        </div>
      </div>
    </div>
  )
}
