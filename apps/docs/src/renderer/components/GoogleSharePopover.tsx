import { useEffect, useRef, useState } from 'react'
import type { GooglePermissionSummary, GoogleRole } from '../../shared/ipc'
import { useModalKeys } from './modal-keys'
import { showToast } from './toast-bus'

const ROLE_LABELS: Record<GoogleRole, string> = {
  reader: 'Viewer',
  commenter: 'Commenter',
  writer: 'Editor',
}

function IconCaretDown14() {
  return (
    <svg width={14} height={14} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M6 9l6 6 6-6"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/** Hand-built dropdown (button + absolutely-positioned menu), matching the
 *  rest of the app's gdh-tb-menu pattern — never a native <select>, so it
 *  never picks up platform select chrome (border/arrow) that a bare <select>
 *  renders regardless of CSS. Closes on outside click or Escape. */
function InlineDropdown({
  label,
  options,
  onSelect,
  ddClassName,
  buttonClassName,
  menuClassName,
}: {
  label: string
  options: Array<{ value: string; label: string; danger?: boolean }>
  onSelect: (value: string) => void
  ddClassName: string
  buttonClassName: string
  menuClassName: string
}) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false)
    }
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false)
    window.addEventListener('mousedown', onDoc)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDoc)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])

  return (
    <div className={ddClassName} ref={ref}>
      <button type="button" className={buttonClassName} onClick={() => setOpen((o) => !o)}>
        {label}
        <IconCaretDown14 />
      </button>
      {open && (
        <div className={`gdh-tb-menu ${menuClassName}`}>
          {options.map((o) => (
            <button
              key={o.value}
              className={o.danger ? 'gdh-share-remove-row' : ''}
              onClick={() => {
                onSelect(o.value)
                setOpen(false)
              }}
            >
              {o.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

function initials(name?: string, email?: string): string {
  const base = (name || email || '?').trim()
  const parts = base.split(/\s+/)
  if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase()
  return base.slice(0, 2).toUpperCase()
}

/** stable per-email hash -> avatar background color, matching the way Docs
 *  gives every collaborator a consistent color across sessions */
const AVATAR_COLORS = [
  '#1a73e8',
  '#d93025',
  '#1e8e3e',
  '#f9ab00',
  '#9334e6',
  '#12b5cb',
  '#e8710a',
  '#795548',
]
function avatarColor(seed: string): string {
  let hash = 0
  for (let i = 0; i < seed.length; i++) hash = (hash * 31 + seed.charCodeAt(i)) >>> 0
  return AVATAR_COLORS[hash % AVATAR_COLORS.length]
}

function IconHelpCircle() {
  return (
    <svg width={20} height={20} viewBox="0 0 24 24" fill="none" aria-hidden>
      <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.6" />
      <path
        d="M9.5 9.2a2.5 2.5 0 1 1 3.7 2.2c-.8.5-1.2.9-1.2 1.8v.3"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
      />
      <circle cx="12" cy="17" r="1" fill="currentColor" />
    </svg>
  )
}

function IconGlobeMedium() {
  return (
    <svg width={20} height={20} viewBox="0 0 24 24" fill="none" aria-hidden>
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.6" />
      <path
        d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"
        stroke="currentColor"
        strokeWidth="1.6"
      />
    </svg>
  )
}

function IconLockMedium() {
  return (
    <svg width={18} height={18} viewBox="0 0 24 24" fill="none" aria-hidden>
      <rect x="5" y="11" width="14" height="9" rx="2" stroke="currentColor" strokeWidth="1.6" />
      <path
        d="M8 11V8a4 4 0 0 1 8 0v3"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
      />
    </svg>
  )
}

function IconLinkSmall() {
  return (
    <svg width={16} height={16} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M9 15l6-6M8 16l-2 2a3.5 3.5 0 0 1-5-5l3-3a3.5 3.5 0 0 1 5-1M16 8l2-2a3.5 3.5 0 0 1 5 5l-3 3a3.5 3.5 0 0 1-5 1"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function IconOpenExternal() {
  return (
    <svg width={16} height={16} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M14 4h6v6M20 4l-9 9M19 13v5a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/**
 * Google-Docs-style share popover, anchored under the header's Share pill.
 * `webViewLink`/`fileId` are null when the current document hasn't been sent
 * to Google Docs yet — that state offers a "Save to Google Docs" action first.
 */
export function GoogleSharePopover({
  fileId,
  webViewLink,
  docName,
  onClose,
  onSaveToGoogleFirst,
  onMoveToFolder,
}: {
  fileId: string | null
  webViewLink: string | null
  docName: string
  onClose: () => void
  /** runs "Send to Google Docs" and resolves once the doc has a fileId */
  onSaveToGoogleFirst: () => Promise<void>
  /** opens the folder picker in move mode for the current googleFileId */
  onMoveToFolder: () => void
}) {
  const modalKeys = useModalKeys(onClose)
  const [saving, setSaving] = useState(false)
  const [permissions, setPermissions] = useState<GooglePermissionSummary[]>([])
  const [loading, setLoading] = useState(false)
  const [email, setEmail] = useState('')
  const [role, setRole] = useState<GoogleRole>('writer')
  const [error, setError] = useState<string | null>(null)
  const [anyoneRole, setAnyoneRole] = useState<'reader' | 'writer' | null>(null)

  const load = async (id: string) => {
    setLoading(true)
    setError(null)
    const result = await window.desktop.googleListPermissions(id)
    setLoading(false)
    if (!result.ok) {
      setError(result.error)
      return
    }
    setPermissions(result.data)
    // 'domain' permissions (Workspace's org-wide default sharing policy) grant
    // the same kind of open access as 'anyone' — treat them the same so the
    // dialog never shows "Restricted" while the file is actually org-visible.
    const openAccess = result.data.find((p) => p.type === 'anyone' || p.type === 'domain')
    setAnyoneRole(openAccess ? (openAccess.role === 'writer' ? 'writer' : 'reader') : null)
  }

  useEffect(() => {
    if (fileId) void load(fileId)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileId])

  const saveFirst = async () => {
    setSaving(true)
    setError(null)
    try {
      await onSaveToGoogleFirst()
    } catch (err) {
      setError(String(err))
    } finally {
      setSaving(false)
    }
  }

  const invite = async () => {
    if (!fileId || !email.trim()) return
    setLoading(true)
    setError(null)
    const result = await window.desktop.googleAddPermission(fileId, email.trim(), role)
    setLoading(false)
    if (!result.ok) {
      setError(result.error)
      return
    }
    setEmail('')
    await load(fileId)
  }

  const changeRole = async (permissionId: string, next: GoogleRole) => {
    if (!fileId) return
    const result = await window.desktop.googleUpdatePermission(fileId, permissionId, next)
    if (!result.ok) {
      setError(result.error)
      return
    }
    await load(fileId)
  }

  const remove = async (permissionId: string) => {
    if (!fileId) return
    const result = await window.desktop.googleRemovePermission(fileId, permissionId)
    if (!result.ok) {
      setError(result.error)
      return
    }
    await load(fileId)
  }

  const toggleAnyone = async (next: 'reader' | 'writer' | null) => {
    if (!fileId) return
    setLoading(true)
    const result = await window.desktop.googleSetAnyoneAccess(fileId, next)
    setLoading(false)
    if (!result.ok) {
      setError(result.error)
      return
    }
    setAnyoneRole(next)
  }

  const copyLink = () => {
    if (!webViewLink) return
    void navigator.clipboard.writeText(webViewLink)
    showToast('Link copied')
  }

  return (
    <div
      className="modal-backdrop gdh-share-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gdh-share-card">
        <h2>Share &ldquo;{docName}&rdquo;</h2>
        <button
          type="button"
          className="gdh-share-help"
          title="Learn about sharing"
          onClick={() =>
            window.desktop.googleOpenExternal('https://support.google.com/drive/answer/7166529')
          }
        >
          <IconHelpCircle />
        </button>

        {!fileId ? (
          <div className="gdh-google-signin">
            <p className="gdh-google-status">
              Save this document to Google Docs to share it with others.
            </p>
            {error && <p className="gdh-google-error">{error}</p>}
            <button
              type="button"
              className="btn-primary"
              disabled={saving}
              onClick={() => void saveFirst()}
            >
              {saving ? 'Saving…' : 'Save to Google Docs'}
            </button>
          </div>
        ) : (
          <>
            {error && <p className="gdh-google-error">{error}</p>}

            <div className="gdh-share-invite">
              <div className="gdh-share-invite-row">
                <div className="gdh-share-invite-field">
                  <input
                    id="gdh-share-email"
                    className="gdh-google-search"
                    placeholder=""
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') void invite()
                    }}
                  />
                  <label className="gdh-share-invite-label" htmlFor="gdh-share-email">
                    Add people, groups, spaces, and calendar events
                  </label>
                </div>
                {email.trim() && (
                  <>
                    <InlineDropdown
                      label={ROLE_LABELS[role]}
                      ddClassName="gdh-share-role-dd"
                      buttonClassName="gdh-share-role-btn"
                      menuClassName="gdh-share-role-menu"
                      options={[
                        { value: 'reader', label: 'Viewer' },
                        { value: 'commenter', label: 'Commenter' },
                        { value: 'writer', label: 'Editor' },
                      ]}
                      onSelect={(v) => setRole(v as GoogleRole)}
                    />
                    <button
                      type="button"
                      className="btn-primary"
                      disabled={loading}
                      onClick={() => void invite()}
                    >
                      Save
                    </button>
                  </>
                )}
              </div>
            </div>

            <div className="gdh-share-section">
              <span className="gdh-share-section-title">People with access</span>
            </div>

            <div className="gdh-share-people">
              {permissions
                .filter((p) => p.type === 'user')
                .map((p) => (
                  <div key={p.id} className="gdh-share-person">
                    <span
                      className="gdh-share-avatar"
                      style={{ background: avatarColor(p.emailAddress || p.displayName || p.id) }}
                    >
                      {initials(p.displayName, p.emailAddress)}
                    </span>
                    <span className="gdh-share-person-text">
                      <span className="gdh-share-person-name">
                        {p.displayName || p.emailAddress}
                      </span>
                      {p.displayName && p.emailAddress && (
                        <span className="gdh-share-person-email">{p.emailAddress}</span>
                      )}
                    </span>
                    {p.role === 'owner' ? (
                      <span className="gdh-share-role-fixed">Owner</span>
                    ) : (
                      <InlineDropdown
                        label={ROLE_LABELS[p.role as GoogleRole] ?? 'Viewer'}
                        ddClassName="gdh-share-role-dd"
                        buttonClassName="gdh-share-role-btn"
                        menuClassName="gdh-share-role-menu"
                        options={[
                          { value: 'reader', label: 'Viewer' },
                          { value: 'commenter', label: 'Commenter' },
                          { value: 'writer', label: 'Editor' },
                          { value: 'remove', label: 'Remove access', danger: true },
                        ]}
                        onSelect={(v) =>
                          v === 'remove'
                            ? void remove(p.id)
                            : void changeRole(p.id, v as GoogleRole)
                        }
                      />
                    )}
                  </div>
                ))}
            </div>

            <div className="gdh-share-general">
              <div className="gdh-share-general-title">General access</div>
              <div className="gdh-share-general-row">
                <span
                  className={`gdh-share-general-icon ${anyoneRole === null ? 'restricted' : ''}`}
                >
                  {anyoneRole === null ? <IconLockMedium /> : <IconGlobeMedium />}
                </span>
                <div className="gdh-share-general-text">
                  <InlineDropdown
                    label={anyoneRole === null ? 'Restricted' : 'Anyone with the link'}
                    ddClassName="gdh-share-general-dd"
                    buttonClassName="gdh-share-general-select"
                    menuClassName="gdh-share-general-menu"
                    options={[
                      { value: 'restricted', label: 'Restricted' },
                      { value: 'anyone', label: 'Anyone with the link' },
                    ]}
                    onSelect={(v) => void toggleAnyone(v === 'anyone' ? 'reader' : null)}
                  />
                  <span className="gdh-share-general-sub">
                    {anyoneRole === null
                      ? 'Only people with access can open with the link'
                      : `Anyone on the internet with the link can ${anyoneRole === 'writer' ? 'edit' : 'view'}`}
                  </span>
                </div>
                {anyoneRole !== null && (
                  <InlineDropdown
                    label={anyoneRole === 'writer' ? 'Editor' : 'Viewer'}
                    ddClassName="gdh-share-role-dd"
                    buttonClassName="gdh-share-role-btn"
                    menuClassName="gdh-share-role-menu"
                    options={[
                      { value: 'reader', label: 'Viewer' },
                      { value: 'writer', label: 'Editor' },
                    ]}
                    onSelect={(v) => void toggleAnyone(v as 'reader' | 'writer')}
                  />
                )}
              </div>
            </div>

            {/* Google's real footer is Copy link + Done only — "Move to Drive
                folder…" and "Open in Google Docs" already live in the File
                menu and the header's send-to-Google icon, so they aren't
                duplicated here. Handlers are still accepted as props (used
                by those other entry points) even though this footer no
                longer renders buttons for them. */}
            <div className="modal-actions gdh-share-footer">
              <button type="button" className="gdh-share-copylink" onClick={copyLink}>
                <IconLinkSmall />
                Copy link
              </button>
              <button type="button" className="gdh-share-done" onClick={onClose}>
                Done
              </button>
            </div>
          </>
        )}

        {!fileId && (
          <div className="modal-actions">
            <button type="button" className="btn-primary" onClick={onClose}>
              Done
            </button>
          </div>
        )}
      </div>
    </div>
  )
}
