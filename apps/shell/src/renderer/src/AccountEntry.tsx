import { useCallback, useEffect, useRef, useState } from 'react'
import type { AccountStatus } from '../../shared/home-api'
import { useI18n } from './locale'
import { SettingsModal } from './SettingsModal'
import type { SettingsTarget } from './SettingsModal'
import { skillUpdateDue } from './IntegrationsPane'

// ── Account entry (bottom-left) ──────────────────────────
// Currently the Genspark (gsk) login entry; to be upgraded to a signup/account system later.
// Clicking it opens the settings modal directly (SettingsModal.tsx), which hosts
// login/logout plus preferences (language, theme, save location, update channel).

const LOGIN_POLL_MS = 2500
/** fallback deadline when the CLI does not report expires_in (device codes live ~300s) */
const LOGIN_MAX_WAIT_MS = 300_000

export function AccountEntry({
  onStatusChange,
  onFileSearchSettingsChange,
  openRequest,
}: {
  onStatusChange?: (status: AccountStatus | null) => void
  onFileSearchSettingsChange?: () => void
  /** a fresh object per request: the home list asks to open the modal on a given block */
  openRequest?: SettingsTarget | null
}) {
  const { t } = useI18n()
  const [status, setStatus] = useState<AccountStatus | null>(null)

  useEffect(() => {
    onStatusChange?.(status)
  }, [status, onStatusChange])
  const [waiting, setWaiting] = useState(false)
  // incremented on login retry, resetting the polling timer
  const [loginNonce, setLoginNonce] = useState(0)
  const [loginError, setLoginError] = useState<
    'timeout' | 'launch' | 'network' | 'expired' | 'failed' | null
  >(null)
  // auth URL reported by the login CLI — rescue entry when the browser did not open
  const [authUrl, setAuthUrl] = useState<string | null>(null)
  const [urlCopied, setUrlCopied] = useState(false)
  const loginDeadline = useRef(0)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [skillUpdate, setSkillUpdate] = useState(false)
  const [loggingOut, setLoggingOut] = useState(false)
  // bumped on logout so an in-flight status refresh (which can still
  // report logged-in) is discarded instead of resurrecting the UI
  const statusSeq = useRef(0)

  // query login state once on mount
  useEffect(() => {
    let alive = true
    void window.aiOffice.accountStatus?.().then((s) => {
      if (alive) setStatus(s)
    })
    return () => {
      alive = false
    }
  }, [])

  // the skill state is a few file reads; re-probe after the modal closes so an
  // update done inside it clears the dot
  useEffect(() => {
    if (settingsOpen) return
    let alive = true
    void window.aiOfficeIntegrations?.status().then((st) => {
      if (alive) setSkillUpdate(skillUpdateDue(st))
    })
    return () => {
      alive = false
    }
  }, [settingsOpen])

  // login progress pushed from main (gsk login CLI output)
  useEffect(() => {
    const off = window.aiOffice.onAccountLogin?.((ev) => {
      if (ev.phase === 'url') {
        if (ev.url) setAuthUrl(ev.url)
        if (ev.expiresInSec) loginDeadline.current = Date.now() + ev.expiresInSec * 1000
      } else if (ev.phase === 'success') {
        void window.aiOffice.accountStatus().then((s) => {
          if (s.loggedIn) {
            setStatus(s)
            setWaiting(false)
            setAuthUrl(null)
          }
        })
      } else if (ev.phase === 'error') {
        setWaiting(false)
        setAuthUrl(null)
        setLoginError(
          ev.error === 'network' ? 'network' : ev.error === 'expired' ? 'expired' : 'failed',
        )
      }
    })
    return off
  }, [])

  // config-file polling stays as the fallback success path (works even if progress events are lost)
  useEffect(() => {
    if (!waiting) return
    const timer = setInterval(() => {
      void window.aiOffice.accountStatus().then((s) => {
        if (s.loggedIn) {
          setStatus(s)
          setWaiting(false)
          setAuthUrl(null)
        } else if (Date.now() > loginDeadline.current) {
          setWaiting(false)
          setAuthUrl(null)
          setLoginError('timeout')
        }
      })
    }, LOGIN_POLL_MS)
    return () => clearInterval(timer)
  }, [waiting, loginNonce])

  const loggedIn = status?.loggedIn ?? false
  const email = status?.email ?? ''
  const initial = email ? email[0].toUpperCase() : loggedIn ? 'G' : '?'
  const errorText = loginError
    ? {
        timeout: t('loginTimeout'),
        launch: t('loginLaunchFailed'),
        network: t('loginNetworkError'),
        expired: t('loginExpired'),
        failed: t('loginFailed'),
      }[loginError]
    : null

  const doLogout = () => {
    setLoggingOut(true)
    statusSeq.current++
    void window.aiOffice.accountLogout().then(() => {
      setLoggingOut(false)
      setStatus({ loggedIn: false })
    })
  }

  const startLogin = () => {
    // clicking again while waiting = relaunch the login (main kills the stale CLI, so the new device code is the live one)
    setLoginError(null)
    setWaiting(true)
    setAuthUrl(null)
    setUrlCopied(false)
    loginDeadline.current = Date.now() + LOGIN_MAX_WAIT_MS
    setLoginNonce((n) => n + 1)
    void window.aiOffice.accountLogin().then((launched) => {
      if (!launched) {
        setWaiting(false)
        setLoginError('launch')
      }
    })
  }

  const openLoginUrl = () => void window.aiOffice.openLoginUrl?.()

  const copyLoginUrl = () => {
    if (!authUrl) return
    void navigator.clipboard.writeText(authUrl).then(() => {
      setUrlCopied(true)
      window.setTimeout(() => setUrlCopied(false), 2000)
    })
  }

  const [target, setTarget] = useState<SettingsTarget | null>(null)
  const openSettings = useCallback((to: SettingsTarget | null) => {
    // refresh the login state / credit balance; drop the response
    // when a logout happened while it was in flight
    const seq = statusSeq.current
    void window.aiOffice.accountStatus?.().then((s) => {
      if (seq === statusSeq.current) setStatus(s)
    })
    setTarget(to)
    setSettingsOpen(true)
  }, [])
  const handleClick = () => openSettings(null)
  useEffect(() => {
    if (openRequest) openSettings(openRequest)
  }, [openRequest, openSettings])

  return (
    <div className="account-entry">
      {settingsOpen && (
        <SettingsModal
          status={status}
          loggingOut={loggingOut}
          loginWaiting={waiting}
          loginUrl={authUrl}
          urlCopied={urlCopied}
          onOpenLoginUrl={openLoginUrl}
          onCopyLoginUrl={copyLoginUrl}
          onClose={() => setSettingsOpen(false)}
          onFileSearchChange={onFileSearchSettingsChange}
          onLogin={() => {
            setSettingsOpen(false)
            startLogin()
          }}
          onLogout={doLogout}
          skillUpdateDue={skillUpdate}
          onSkillUpdateDue={setSkillUpdate}
          target={target}
        />
      )}
      {!settingsOpen && waiting && authUrl && (
        <div className="login-hint" role="status">
          <button className="login-hint-open" onClick={openLoginUrl}>
            {t('loginOpenShort')}
          </button>
          <button
            className={`login-hint-copy${urlCopied ? ' copied' : ''}`}
            onClick={copyLoginUrl}
            // static tip: screentips are suppressed from pointerdown until the pointer
            // leaves the control, so a swapped-in "copied" tip would never show — the
            // check-mark icon is the visible feedback
            data-tip={t('loginCopyUrl')}
            aria-label={urlCopied ? t('loginCopied') : t('loginCopyUrl')}
          >
            {urlCopied ? (
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path
                  d="m3.5 8.5 3 3 6-7"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            ) : (
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <rect
                  x="5.5"
                  y="5.5"
                  width="7"
                  height="7"
                  rx="1.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                />
                <path
                  d="M3.5 10.5V5a1.5 1.5 0 0 1 1.5-1.5h5.5"
                  stroke="currentColor"
                  strokeWidth="1.3"
                  strokeLinecap="round"
                />
              </svg>
            )}
          </button>
        </div>
      )}
      <button
        className="account-btn"
        onClick={handleClick}
        aria-haspopup="dialog"
        aria-expanded={settingsOpen}
        data-tip={
          loggedIn
            ? email || t('loggedInGenspark')
            : waiting
              ? t('waitingLogin')
              : (errorText ?? t('loginGenspark'))
        }
        aria-label={t('settings')}
      >
        <span
          className={`account-avatar${loggedIn ? ' logged-in' : ''}${waiting ? ' waiting' : ''}`}
        >
          {waiting ? (
            <svg
              className="account-spinner"
              width="14"
              height="14"
              viewBox="0 0 16 16"
              aria-hidden="true"
            >
              <circle
                cx="8"
                cy="8"
                r="6"
                stroke="currentColor"
                strokeWidth="1.8"
                fill="none"
                strokeDasharray="26"
                strokeDashoffset="18"
                strokeLinecap="round"
              />
            </svg>
          ) : (
            initial
          )}
          {skillUpdate && (
            <span className="account-badge" role="img" aria-label={t('intgUpdateDue')} />
          )}
        </span>
        <span className="account-text">
          <span className="account-name">
            {loggedIn
              ? email
                ? email.split('@')[0]
                : t('loggedIn')
              : waiting
                ? t('waitingShort')
                : t('login')}
          </span>
          {!loggedIn && !waiting && errorText && (
            <span className="account-sub error">{errorText}</span>
          )}
        </span>
        <svg
          className="account-chevron"
          width="14"
          height="14"
          viewBox="0 0 16 16"
          fill="none"
          aria-hidden="true"
        >
          <path
            d="M5 6.2 8 3.4l3 2.8M5 9.8l3 2.8 3-2.8"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
    </div>
  )
}
