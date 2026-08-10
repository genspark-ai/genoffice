import { useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { AccountStatus, UiTheme } from '../../shared/home-api'
import './settings.css'

// ── Settings modal (opened from the account menu) ─────────
// Zoom-style two-pane dialog: section nav on the left, fields on the right.
// All values go through the existing home IPC; nothing is stored locally.

// sorted by ISO 639 language code — native-script labels have no natural
// shared alphabet, so the code is the ordering key
const LANG_OPTIONS = [
  { value: 'ar', label: 'العربية' },
  { value: 'de', label: 'Deutsch' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'fr', label: 'Français' },
  { value: 'he', label: 'עברית' },
  { value: 'hi', label: 'हिन्दी' },
  { value: 'id', label: 'Bahasa Indonesia' },
  { value: 'it', label: 'Italiano' },
  { value: 'ja', label: '日本語' },
  { value: 'ko', label: '한국어' },
  { value: 'ms', label: 'Bahasa Melayu' },
  { value: 'nl', label: 'Nederlands' },
  { value: 'pl', label: 'Polski' },
  { value: 'pt', label: 'Português' },
  { value: 'ru', label: 'Русский' },
  { value: 'th', label: 'ไทย' },
  { value: 'zh', label: '简体中文' },
  { value: 'zh-TW', label: '繁體中文' },
] as const

const THEME_OPTIONS = [
  { value: 'light', labelKey: 'themeLight' },
  { value: 'dark', labelKey: 'themeDark' },
  { value: 'system', labelKey: 'themeSystem' },
] as const satisfies readonly { value: UiTheme; labelKey: StringKey }[]

const CHANNEL_OPTIONS = [
  { value: 'stable', labelKey: 'channelStable' },
  { value: 'beta', labelKey: 'channelBeta' },
] as const satisfies readonly { value: 'stable' | 'beta'; labelKey: StringKey }[]

type SectionId = 'account' | 'general' | 'about'

const SECTIONS: readonly { id: SectionId; labelKey: StringKey }[] = [
  { id: 'account', labelKey: 'setSecAccount' },
  { id: 'general', labelKey: 'setSecGeneral' },
  { id: 'about', labelKey: 'setSecAbout' },
]

function SectionIcon({ id }: { id: SectionId }) {
  if (id === 'account') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <circle cx="8" cy="5.2" r="2.9" stroke="currentColor" strokeWidth="1.3" />
        <path
          d="M2.7 13.6a5.5 5.5 0 0 1 10.6 0"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    )
  }
  if (id === 'general') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M2 5h8M13 5h1M2 11h1M6 11h8"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
        <circle cx="11.5" cy="5" r="1.7" stroke="currentColor" strokeWidth="1.3" />
        <circle cx="4.5" cy="11" r="1.7" stroke="currentColor" strokeWidth="1.3" />
      </svg>
    )
  }
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.3" />
      <path d="M8 7.4v3.4" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      <circle cx="8" cy="5.1" r="0.8" fill="currentColor" />
    </svg>
  )
}

/** label-over-value field row with an optional right-aligned action */
function Field({
  label,
  value,
  valueTitle,
  action,
}: {
  label: string
  value: string
  valueTitle?: string
  action?: ReactNode
}) {
  return (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-label">{label}</div>
        <div className="set-field-value" data-tip={valueTitle}>
          {value}
        </div>
      </div>
      {action}
    </div>
  )
}

export interface SettingsModalProps {
  status: AccountStatus | null
  loggingOut: boolean
  /** browser sign-in in progress (spinner shows on the account entry) */
  loginWaiting: boolean
  /** device auth URL while waiting — rescue actions when the browser did not auto-open */
  loginUrl: string | null
  urlCopied: boolean
  onOpenLoginUrl: () => void
  onCopyLoginUrl: () => void
  onClose: () => void
  /** closes the modal and launches the Genspark login flow (progress shows on the account entry) */
  onLogin: () => void
  onLogout: () => void
}

export function SettingsModal({
  status,
  loggingOut,
  loginWaiting,
  loginUrl,
  urlCopied,
  onOpenLoginUrl,
  onCopyLoginUrl,
  onClose,
  onLogin,
  onLogout,
}: SettingsModalProps) {
  const { lang, setLang, t } = useI18n()
  const [section, setSection] = useState<SectionId>('account')
  const [theme, setTheme] = useState<UiTheme>('system')
  const [saveDir, setSaveDir] = useState('')
  const [channel, setChannel] = useState<'stable' | 'beta'>('stable')
  const [appVersion, setAppVersion] = useState('')

  useEffect(() => {
    let alive = true
    void window.aiOffice.getTheme?.().then((th) => {
      if (alive) setTheme(th)
    })
    void window.aiOffice.getDefaultSaveDir?.().then((dir) => {
      if (alive && dir) setSaveDir(dir)
    })
    void window.aiOffice.getUpdateChannel?.().then((ch) => {
      if (alive) setChannel(ch)
    })
    void window.aiOffice.getAppVersion?.().then((v) => {
      if (alive && v) setAppVersion(v)
    })
    return () => {
      alive = false
    }
  }, [])

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  const applyTheme = (next: UiTheme) => {
    setTheme(next)
    void window.aiOffice.setTheme(next)
    if (next === 'system') document.documentElement.removeAttribute('data-theme')
    else document.documentElement.setAttribute('data-theme', next)
  }

  const changeSaveDir = () => {
    void window.aiOffice.pickDefaultSaveDir?.().then((dir) => {
      if (dir) setSaveDir(dir)
    })
  }

  const loggedIn = status?.loggedIn ?? false
  const email = status?.email ?? ''

  return (
    <div
      className="set-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div className="set-dialog" role="dialog" aria-modal="true" aria-label={t('settings')}>
        <div className="set-header">
          <h2 className="set-title">{t('settings')}</h2>
          <button className="set-close" onClick={onClose} aria-label={t('cancel')}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="set-body">
          <nav className="set-nav" aria-label={t('settings')}>
            {SECTIONS.map((s) => (
              <button
                key={s.id}
                className={`set-nav-item${section === s.id ? ' active' : ''}`}
                aria-current={section === s.id}
                onClick={() => setSection(s.id)}
              >
                <SectionIcon id={s.id} />
                {t(s.labelKey)}
              </button>
            ))}
          </nav>
          <div className="set-pane">
            {section === 'account' && (
              <>
                <h3 className="set-pane-title">{t('setSecAccount')}</h3>
                <Field label={t('setEmail')} value={loggedIn ? email : t('setNotLoggedIn')} />
                {loggedIn && (
                  <Field
                    label={t('credits')}
                    value={
                      status?.creditBalance === undefined
                        ? '—'
                        : Math.floor(status.creditBalance).toLocaleString('en-US')
                    }
                    action={
                      <button
                        className="set-btn"
                        data-tip={t('creditsTip')}
                        onClick={() => void window.aiOffice.openCreditUsage?.()}
                      >
                        {t('setViewUsage')}
                      </button>
                    }
                  />
                )}
                <div className="set-pane-footer">
                  {loggedIn ? (
                    <button className="set-btn danger" disabled={loggingOut} onClick={onLogout}>
                      {loggingOut ? t('loggingOut') : t('logout')}
                    </button>
                  ) : (
                    <>
                      {loginWaiting && loginUrl && (
                        <>
                          <button className="set-btn" onClick={onOpenLoginUrl}>
                            {t('loginOpenManually')}
                          </button>
                          <button className="set-btn" onClick={onCopyLoginUrl}>
                            {urlCopied ? t('loginCopied') : t('loginCopyUrl')}
                          </button>
                        </>
                      )}
                      <button className="set-btn primary" onClick={onLogin}>
                        {loginWaiting ? t('waitingShort') : t('loginGenspark')}
                      </button>
                    </>
                  )}
                </div>
              </>
            )}
            {section === 'general' && (
              <>
                <h3 className="set-pane-title">{t('setSecGeneral')}</h3>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-lang">
                      {t('language')}
                    </label>
                  </div>
                  <select
                    id="set-lang"
                    className="set-select"
                    value={lang}
                    onChange={(e) => setLang(e.target.value as typeof lang)}
                  >
                    {LANG_OPTIONS.map((opt) => (
                      <option key={opt.value} value={opt.value}>
                        {opt.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-theme">
                      {t('theme')}
                    </label>
                  </div>
                  <select
                    id="set-theme"
                    className="set-select"
                    value={theme}
                    onChange={(e) => applyTheme(e.target.value as UiTheme)}
                  >
                    {THEME_OPTIONS.map((opt) => (
                      <option key={opt.value} value={opt.value}>
                        {t(opt.labelKey)}
                      </option>
                    ))}
                  </select>
                </div>
                <Field
                  label={t('saveLocation')}
                  value={saveDir || '—'}
                  valueTitle={saveDir}
                  action={
                    <button className="set-btn" onClick={changeSaveDir}>
                      {t('setChange')}
                    </button>
                  }
                />
              </>
            )}
            {section === 'about' && (
              <>
                <h3 className="set-pane-title">{t('setSecAbout')}</h3>
                <Field label={t('versionLabel')} value={appVersion || '—'} />
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-channel">
                      {t('updateChannel')}
                    </label>
                  </div>
                  <select
                    id="set-channel"
                    className="set-select"
                    value={channel}
                    onChange={(e) => {
                      const next = e.target.value === 'beta' ? 'beta' : 'stable'
                      setChannel(next)
                      void window.aiOffice.setUpdateChannel(next)
                    }}
                  >
                    {CHANNEL_OPTIONS.map((opt) => (
                      <option key={opt.value} value={opt.value}>
                        {t(opt.labelKey)}
                      </option>
                    ))}
                  </select>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
