/**
 * Review > Protect Sheet / Unprotect Sheet (Excel's dialogs): an optional
 * password plus the list of actions that stay allowed while protected, or
 * the password prompt that lifts a password-protected sheet.
 */
import { useState } from 'react'

import {
  DEFAULT_SHEET_PROTECTION_ALLOW,
  SHEET_PROTECTION_PERMISSIONS,
  type SheetProtectionAllow,
  type SheetProtectionPermission,
} from '@genoffice/xlsx-gateway/gateway/xlsx-protection'

import { useI18n } from './i18n/locale'
import type { zh } from './i18n/dialogs/zh'
import { useModalDialog } from './modal-dialog'

const PERMISSION_LABELS: Record<SheetProtectionPermission, keyof typeof zh> = {
  selectLockedCells: 'dlgProtAllowSelectLockedCells',
  selectUnlockedCells: 'dlgProtAllowSelectUnlockedCells',
  formatCells: 'dlgProtAllowFormatCells',
  formatColumns: 'dlgProtAllowFormatColumns',
  formatRows: 'dlgProtAllowFormatRows',
  insertColumns: 'dlgProtAllowInsertColumns',
  insertRows: 'dlgProtAllowInsertRows',
  insertHyperlinks: 'dlgProtAllowInsertHyperlinks',
  deleteColumns: 'dlgProtAllowDeleteColumns',
  deleteRows: 'dlgProtAllowDeleteRows',
  sort: 'dlgProtAllowSort',
  autoFilter: 'dlgProtAllowAutoFilter',
  pivotTables: 'dlgProtAllowPivotTables',
  objects: 'dlgProtAllowObjects',
  scenarios: 'dlgProtAllowScenarios',
}

export function ProtectSheetDialog({
  mode,
  onProtect,
  onUnprotect,
  onClose,
}: {
  readonly mode: 'protect' | 'unprotect'
  readonly onProtect: (request: {
    password: string | null
    allow: SheetProtectionAllow
  }) => Promise<string | null>
  readonly onUnprotect: (password: string) => Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [allow, setAllow] = useState<SheetProtectionAllow>(DEFAULT_SHEET_PROTECTION_ALLOW)
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const modal = useModalDialog(onClose)
  const title = t(mode === 'protect' ? 'dlgProtectSheetTitle' : 'dlgUnprotectSheetTitle')

  const apply = (): void => {
    if (busy) return
    if (mode === 'protect' && password !== confirm) {
      setError(t('dlgProtectSheetMismatch'))
      return
    }
    setBusy(true)
    const run =
      mode === 'protect'
        ? onProtect({ password: password === '' ? null : password, allow })
        : onUnprotect(password)
    void run
      .then((failed) => {
        if (failed !== null) setError(failed)
        else onClose()
      })
      .catch((reason: unknown) => {
        setError(reason instanceof Error ? reason.message : String(reason))
      })
      .finally(() => setBusy(false))
  }

  const onEnter = (event: React.KeyboardEvent): void => {
    if (event.key === 'Enter') apply()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{title}</header>
        <section className="dialog-body">
          <div className="dialog-grid">
            <label>
              {t(mode === 'protect' ? 'dlgProtectSheetPassword' : 'dlgUnprotectSheetPrompt')}
              <input
                autoFocus
                type="password"
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                onKeyDown={onEnter}
              />
            </label>
            {mode === 'protect' && (
              <label>
                {t('dlgProtectSheetConfirm')}
                <input
                  type="password"
                  value={confirm}
                  onChange={(event) => setConfirm(event.target.value)}
                  onKeyDown={onEnter}
                />
              </label>
            )}
          </div>
          {mode === 'protect' && (
            <>
              <p className="dialog-note">{t('dlgProtectSheetAllow')}</p>
              <ul className="protect-sheet-allow">
                {SHEET_PROTECTION_PERMISSIONS.map((permission) => (
                  <li key={permission}>
                    <label>
                      <input
                        type="checkbox"
                        checked={allow[permission]}
                        onChange={(event) =>
                          setAllow({ ...allow, [permission]: event.target.checked })
                        }
                      />
                      {t(PERMISSION_LABELS[permission])}
                    </label>
                  </li>
                ))}
              </ul>
            </>
          )}
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
          <button className="primary-action" disabled={busy} onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
