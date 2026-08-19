/**
 * Word-style "Protect Document" dialog (Review > Protect), modeled on Word for
 * Mac's Password Protect sheet. One dialog covers all protection features:
 *
 * - Security: password to open (ECMA-376 whole-file encryption) and password
 *   to modify (settings.xml w:writeProtection, honor-system).
 * - Protection: editing restriction (w:documentProtection) with the four Word
 *   modes (tracked changes / comments / read only / forms) + optional password.
 * - Privacy: remove known author and organization metadata on save.
 *
 * The dialog only computes a diff (ProtectDialogResult); applying it (IPC for
 * the open password, dirty-state for the rest) is the caller's job.
 */
import { useState } from 'react'
import {
  hashProtectionPassword,
  verifyProtectionPassword,
  type DocProtection,
  type WriteProtection,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'

/** every field: undefined = unchanged; null = remove; value = set */
export interface ProtectDialogResult {
  openPassword?: string | null
  writeProtection?: WriteProtection | null
  protection?: DocProtection | null
  removePersonalInfo?: boolean
}

export const PROTECTION_MODES = ['trackedChanges', 'comments', 'readOnly', 'forms'] as const
export type ProtectionMode = (typeof PROTECTION_MODES)[number]

const MODE_LABEL_KEYS = {
  trackedChanges: 'appProtectModeTracked',
  comments: 'appProtectModeComments',
  readOnly: 'appProtectModeReadOnly',
  forms: 'appProtectModeForms',
} as const

/**
 * Existing passwords are only known as hashes, so the fields are prefilled
 * with an untypable sentinel that renders as dots: submitting it unchanged
 * keeps the password, clearing the field removes it.
 */
const KEEP = '\u0001'.repeat(8)

type ErrorKey = '' | 'appEncMismatch' | 'appWrongPassword'

export function ProtectDialog({
  encrypted,
  writeProtection,
  protection,
  removePersonalInfo,
  onCancel,
  onApply,
}: {
  /** an open password is desired for the next save */
  encrypted: boolean
  writeProtection: WriteProtection | null
  protection: DocProtection | null
  removePersonalInfo: boolean
  onCancel: () => void
  onApply: (result: ProtectDialogResult) => void
}) {
  const { t } = useI18n()
  const hadModifyPwd = !!writeProtection?.hash
  const wasEnforced = !!protection?.enforced
  /** changing/removing an enforced password-protected restriction needs the password */
  const locked = wasEnforced && !!protection?.hash

  const [openPwd, setOpenPwd] = useState(encrypted ? KEEP : '')
  const [openPwd2, setOpenPwd2] = useState(encrypted ? KEEP : '')
  const [modifyPwd, setModifyPwd] = useState(hadModifyPwd ? KEEP : '')
  const [modifyPwd2, setModifyPwd2] = useState(hadModifyPwd ? KEEP : '')
  const [protectOn, setProtectOn] = useState(wasEnforced)
  const [mode, setMode] = useState<ProtectionMode>(
    PROTECTION_MODES.includes(protection?.edit as ProtectionMode)
      ? (protection?.edit as ProtectionMode)
      : 'trackedChanges',
  )
  const [protectPwd, setProtectPwd] = useState(locked ? KEEP : '')
  const [unlockPwd, setUnlockPwd] = useState('')
  const [removePersonal, setRemovePersonal] = useState(removePersonalInfo)
  const [errorKey, setErrorKey] = useState<ErrorKey>('')
  const [busy, setBusy] = useState(false)

  const protectionChanged =
    protectOn !== wasEnforced ||
    (protectOn && (mode !== protection?.edit || protectPwd !== (locked ? KEEP : '')))

  const submit = async () => {
    if (busy) return
    if (openPwd !== openPwd2 || modifyPwd !== modifyPwd2) {
      setErrorKey('appEncMismatch')
      return
    }
    setBusy(true)
    try {
      const result: ProtectDialogResult = {}

      if (openPwd !== (encrypted ? KEEP : '')) {
        result.openPassword = openPwd === '' ? null : openPwd
      }

      if (modifyPwd !== (hadModifyPwd ? KEEP : '')) {
        const recommended = writeProtection?.recommended ? { recommended: true } : {}
        result.writeProtection =
          modifyPwd === ''
            ? writeProtection?.recommended
              ? { recommended: true }
              : null
            : { ...recommended, ...(await hashProtectionPassword(modifyPwd)) }
      }

      if (protectionChanged) {
        if (locked && !(await verifyProtectionPassword(unlockPwd, protection!))) {
          setErrorKey('appWrongPassword')
          return
        }
        if (!protectOn) {
          result.protection = null
        } else {
          const creds =
            protectPwd === KEEP && protection?.hash
              ? {
                  hash: protection.hash,
                  salt: protection.salt,
                  spinCount: protection.spinCount,
                  algorithmSid: protection.algorithmSid,
                }
              : protectPwd && protectPwd !== KEEP
                ? await hashProtectionPassword(protectPwd)
                : {}
          result.protection = { edit: mode, enforced: true, ...creds }
        }
      }

      if (removePersonal !== removePersonalInfo) {
        result.removePersonalInfo = removePersonal
      }

      onApply(result)
    } finally {
      setBusy(false)
    }
  }

  const onEnter = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') void submit()
  }
  const clearError = () => setErrorKey('')

  return (
    <div className="modal-backdrop">
      <div className="modal protect-dialog">
        <h2>{t('appProtectTitle')}</h2>

        <h3 className="protect-section-title">{t('appProtectSecurity')}</h3>
        <label>
          {t('appProtectOpenPwd')}
          <input
            type="password"
            autoFocus
            maxLength={255}
            value={openPwd}
            onChange={(e) => {
              setOpenPwd(e.target.value)
              clearError()
            }}
            onKeyDown={onEnter}
          />
        </label>
        <label>
          {t('appEncConfirmLabel')}
          <input
            type="password"
            maxLength={255}
            value={openPwd2}
            onChange={(e) => {
              setOpenPwd2(e.target.value)
              clearError()
            }}
            onKeyDown={onEnter}
          />
        </label>
        <label>
          {t('appProtectModifyPwd')}
          <input
            type="password"
            maxLength={255}
            value={modifyPwd}
            onChange={(e) => {
              setModifyPwd(e.target.value)
              clearError()
            }}
            onKeyDown={onEnter}
          />
        </label>
        <label>
          {t('appEncConfirmLabel')}
          <input
            type="password"
            maxLength={255}
            value={modifyPwd2}
            onChange={(e) => {
              setModifyPwd2(e.target.value)
              clearError()
            }}
            onKeyDown={onEnter}
          />
        </label>
        {(encrypted || hadModifyPwd) && <p className="pgnum-hint">{t('appProtectPwdKeepHint')}</p>}

        <h3 className="protect-section-title">{t('appProtectSectionTitle')}</h3>
        <label className="protect-check">
          <input
            type="checkbox"
            checked={protectOn}
            onChange={(e) => {
              setProtectOn(e.target.checked)
              clearError()
            }}
          />
          {t('appProtectFor')}
        </label>
        <div className="protect-modes">
          {PROTECTION_MODES.map((m) => (
            <label key={m} className="protect-check">
              <input
                type="radio"
                name="protect-mode"
                disabled={!protectOn}
                checked={mode === m}
                onChange={() => {
                  setMode(m)
                  clearError()
                }}
              />
              {t(MODE_LABEL_KEYS[m])}
            </label>
          ))}
        </div>
        {protectOn && (
          <label>
            {t('appProtectPwdOptional')}
            <input
              type="password"
              maxLength={255}
              value={protectPwd}
              onChange={(e) => {
                setProtectPwd(e.target.value)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
        )}
        {locked && protectionChanged && (
          <label>
            {t('appProtectUnlockPwd')}
            <input
              type="password"
              value={unlockPwd}
              onChange={(e) => {
                setUnlockPwd(e.target.value)
                clearError()
              }}
              onKeyDown={onEnter}
            />
          </label>
        )}

        <h3 className="protect-section-title">{t('appProtectPrivacy')}</h3>
        <label className="protect-check">
          <input
            type="checkbox"
            checked={removePersonal}
            onChange={(e) => setRemovePersonal(e.target.checked)}
          />
          {t('appProtectRemovePersonal')}
        </label>

        {errorKey && <p className="modal-error">{t(errorKey)}</p>}
        <div className="modal-actions">
          <button onClick={onCancel}>{t('appCancel')}</button>
          <button className="btn-primary" disabled={busy} onClick={() => void submit()}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
