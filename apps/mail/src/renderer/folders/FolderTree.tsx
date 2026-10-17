import { useState } from 'react'
import type { MailFolder } from '@genoffice/mail-protocols'
import type { MailAccountSummary } from '../../shared/ipc'
import { useI18n, type TFunc } from '../i18n/locale'
import './folders.css'

/**
 * Account + folder tree (Lane R owns). Accounts come from
 * `mail:list-accounts`, folders from `mail:list-folders`; unread badges
 * use the folder counters. Removing an account is a two-step inline
 * confirm — no blocking dialogs, so the flow stays testable.
 */

export interface FolderSelection {
  accountId: string
  folderId: string
}

export interface FolderTreeProps {
  accounts: MailAccountSummary[]
  foldersByAccount: Record<string, MailFolder[]>
  selected: FolderSelection | null
  onSelect(selection: FolderSelection): void
  onAddAccount(): void
  onRemoveAccount(accountId: string): void
}

function folderLabel(t: TFunc, folder: MailFolder): string {
  switch (folder.role) {
    case 'inbox':
      return t('mail.folders.inbox')
    case 'sent':
      return t('mail.folders.sent')
    case 'drafts':
      return t('mail.folders.drafts')
    case 'trash':
      return t('mail.folders.trash')
    case 'spam':
      return t('mail.folders.spam')
    case 'archive':
      return t('mail.folders.archive')
    case 'custom':
      return folder.name
  }
}

export function FolderTree({
  accounts,
  foldersByAccount,
  selected,
  onSelect,
  onAddAccount,
  onRemoveAccount,
}: FolderTreeProps) {
  const { t } = useI18n()
  const [confirmingRemove, setConfirmingRemove] = useState<string | null>(null)

  return (
    <div className="mail-folders" aria-label={t('mail.accounts.title')}>
      <div className="mail-folders-accounts">
        <span className="mail-folders-heading">{t('mail.accounts.title')}</span>
        <button type="button" className="mail-folders-add" onClick={onAddAccount}>
          {t('mail.accounts.add')}
        </button>
      </div>
      {accounts.map((account) => (
        <section
          key={account.accountId}
          className="mail-folders-account"
          aria-label={account.label}
        >
          <div className="mail-folders-account-row">
            <span className="mail-folders-account-name">{account.label}</span>
            {account.sessionOnly && (
              <span className="mail-folders-session">{t('mail.accounts.sessionOnly')}</span>
            )}
            {confirmingRemove === account.accountId ? (
              <span className="mail-folders-confirm">
                <span>{t('mail.accounts.removeConfirm')}</span>
                <button type="button" onClick={() => onRemoveAccount(account.accountId)}>
                  {t('mail.accounts.remove')}
                </button>
                <button type="button" onClick={() => setConfirmingRemove(null)}>
                  {t('mail.btnCancel')}
                </button>
              </span>
            ) : (
              <button
                type="button"
                className="mail-folders-remove"
                aria-label={`${t('mail.accounts.remove')}: ${account.label}`}
                onClick={() => setConfirmingRemove(account.accountId)}
              >
                {t('mail.accounts.remove')}
              </button>
            )}
          </div>
          <ul className="mail-folders-list" role="tree" aria-label={account.label}>
            {(foldersByAccount[account.accountId] ?? []).map((folder) => {
              const active =
                selected?.accountId === account.accountId && selected?.folderId === folder.id
              return (
                <li key={folder.id} role="none">
                  <button
                    type="button"
                    role="treeitem"
                    aria-selected={active}
                    aria-current={active ? 'true' : undefined}
                    className={active ? 'mail-folder is-active' : 'mail-folder'}
                    onClick={() => onSelect({ accountId: account.accountId, folderId: folder.id })}
                  >
                    <span
                      className={`mail-folder-glyph mail-folder-${folder.role}`}
                      aria-hidden="true"
                    />
                    <span className="mail-folder-name">{folderLabel(t, folder)}</span>
                    {folder.unread > 0 && (
                      <span
                        className="mail-folder-unread"
                        aria-label={`${folder.unread} ${t('mail.folders.unreadCount')}`}
                      >
                        {folder.unread}
                      </span>
                    )}
                  </button>
                </li>
              )
            })}
          </ul>
        </section>
      ))}
    </div>
  )
}
