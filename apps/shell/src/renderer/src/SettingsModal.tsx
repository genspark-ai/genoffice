import {
  DEFAULT_APPEARANCE,
  applyAppearance,
  contrastRatio,
  loadAppearance,
  parseHex,
  saveAppearance,
} from './appearance'
import type { Appearance, UiFont } from './appearance'
import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import type { CSSProperties, ReactNode } from 'react'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { AccountStatus, UiTheme } from '../../shared/home-api'
import './settings.css'

// ── Settings modal (opened from the account menu) ─────────
// Genspark-style two-pane dialog: section nav on the left, fields on the right.
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

// GenMail's option order: follow-system first, then the manual picks
const ACCENT_PRESETS = ['#339cff', '#12b5a6', '#3dd68c', '#f5a623', '#8a8a8a'] as const

const FONT_OPTIONS: readonly { value: UiFont; labelKey: StringKey }[] = [
  { value: 'system', labelKey: 'fontSystem' },
  { value: 'sans', labelKey: 'fontSans' },
  { value: 'serif', labelKey: 'fontSerif' },
  { value: 'mono', labelKey: 'fontMono' },
]

const THEME_OPTIONS = [
  { value: 'system', labelKey: 'themeSystem' },
  { value: 'light', labelKey: 'themeLight' },
  { value: 'dark', labelKey: 'themeDark' },
] as const satisfies readonly { value: UiTheme; labelKey: StringKey }[]

const CHANNEL_OPTIONS = [
  { value: 'stable', labelKey: 'channelStable' },
  { value: 'beta', labelKey: 'channelBeta' },
] as const satisfies readonly { value: 'stable' | 'beta'; labelKey: StringKey }[]

type SectionId = 'account' | 'general' | 'appearance' | 'shortcuts' | 'about'

const SECTIONS: readonly { id: SectionId; labelKey: StringKey }[] = [
  { id: 'account', labelKey: 'setSecAccount' },
  { id: 'general', labelKey: 'setSecGeneral' },
  { id: 'appearance', labelKey: 'setSecAppearance' },
  { id: 'shortcuts', labelKey: 'setSecShortcuts' },
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
  if (id === 'appearance') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.3" />
        <path d="M8 1.7v12.6" stroke="currentColor" strokeWidth="1.3" />
        <path d="M8 1.7a6.3 6.3 0 0 1 0 12.6z" fill="currentColor" />
      </svg>
    )
  }
  if (id === 'shortcuts') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect
          x="1.7"
          y="3.4"
          width="12.6"
          height="9.2"
          rx="2"
          stroke="currentColor"
          strokeWidth="1.3"
        />
        <path
          d="M4.6 6.6h.01M7.2 6.6h.01M9.8 6.6h.01M12 6.6h.01M5 9.6h6"
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
  const [appearance, setAppearance] = useState<Appearance>(() =>
    loadAppearance(window.localStorage),
  )
  const [hexText, setHexText] = useState(appearance.accent ?? '')

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

  const updateAppearance = (patch: Partial<Appearance>) => {
    const next = { ...appearance, ...patch }
    setAppearance(next)
    saveAppearance(window.localStorage, next)
    applyAppearance(document.documentElement, next)
  }

  const resetAppearance = () => {
    setHexText('')
    updateAppearance({ ...DEFAULT_APPEARANCE })
  }

  const surfaceHex = parseHex(
    getComputedStyle(document.documentElement).getPropertyValue('--surface'),
  )
  const accentLowContrast =
    appearance.accent !== null &&
    surfaceHex !== null &&
    contrastRatio(appearance.accent, surfaceHex) < 3
  const hexInvalid = hexText.trim() !== '' && parseHex(hexText) === null
  const modKey = navigator.platform.toLowerCase().includes('mac') ? '⌘' : 'Ctrl'

  const changeSaveDir = () => {
    void window.aiOffice.pickDefaultSaveDir?.().then((dir) => {
      if (dir) setSaveDir(dir)
    })
  }

  const loggedIn = status?.loggedIn ?? false
  const email = status?.email ?? ''

  const modal = (
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
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">
                      {LANG_OPTIONS.find((o) => o.value === lang)?.label ?? lang}
                    </span>
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
                  </span>
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-theme">
                      {t('theme')}
                    </label>
                  </div>
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">
                      {t(THEME_OPTIONS.find((o) => o.value === theme)?.labelKey ?? 'themeSystem')}
                    </span>
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
                  </span>
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
            {section === 'appearance' && (
              <>
                <h3 className="set-pane-title">{t('setSecAppearance')}</h3>
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-label">{t('appearanceMode')}</div>
                  </div>
                  <div className="set-mode-row" role="radiogroup" aria-label={t('appearanceMode')}>
                    {THEME_OPTIONS.map((opt) => (
                      <button
                        key={opt.value}
                        role="radio"
                        aria-checked={theme === opt.value}
                        className={`set-mode-card${theme === opt.value ? ' active' : ''}`}
                        onClick={() => applyTheme(opt.value)}
                      >
                        {t(opt.labelKey)}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="set-field set-field-col">
                  <div className="set-field-text">
                    <div className="set-field-label">{t('appearanceAccent')}</div>
                  </div>
                  <div className="set-accent-row">
                    <button
                      className={`set-swatch default${appearance.accent === null ? ' active' : ''}`}
                      aria-label={t('appearanceAccentDefault')}
                      aria-pressed={appearance.accent === null}
                      title={t('appearanceAccentDefault')}
                      onClick={() => {
                        setHexText('')
                        updateAppearance({ accent: null })
                      }}
                    />
                    {ACCENT_PRESETS.map((hex) => (
                      <button
                        key={hex}
                        className={`set-swatch${appearance.accent === hex ? ' active' : ''}`}
                        style={{ '--swatch': hex } as CSSProperties}
                        aria-label={hex}
                        aria-pressed={appearance.accent === hex}
                        title={hex}
                        onClick={() => {
                          setHexText(hex)
                          updateAppearance({ accent: hex })
                        }}
                      />
                    ))}
                    <input
                      className="set-hex"
                      value={hexText}
                      placeholder="#339CFF"
                      spellCheck={false}
                      maxLength={7}
                      aria-label={t('appearanceAccentHex')}
                      aria-invalid={hexInvalid}
                      onChange={(e) => {
                        setHexText(e.target.value)
                        const parsed = parseHex(e.target.value)
                        if (parsed) updateAppearance({ accent: parsed })
                      }}
                    />
                  </div>
                  {hexInvalid && <div className="set-warn">{t('appearanceAccentInvalid')}</div>}
                  {!hexInvalid && accentLowContrast && (
                    <div className="set-warn">{t('appearanceAccentLow')}</div>
                  )}
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-font">
                      {t('appearanceFont')}
                    </label>
                  </div>
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">
                      {t(
                        FONT_OPTIONS.find((o) => o.value === appearance.font)?.labelKey ??
                          'fontSystem',
                      )}
                    </span>
                    <select
                      id="set-font"
                      className="set-select"
                      value={appearance.font}
                      onChange={(e) => updateAppearance({ font: e.target.value as UiFont })}
                    >
                      {FONT_OPTIONS.map((opt) => (
                        <option key={opt.value} value={opt.value}>
                          {t(opt.labelKey)}
                        </option>
                      ))}
                    </select>
                  </span>
                </div>
                <label className="set-field set-switch-row">
                  <div className="set-field-text">
                    <div className="set-field-label">{t('appearanceTransparency')}</div>
                    <div className="set-field-value">{t('appearanceTransparencyHint')}</div>
                  </div>
                  <input
                    type="checkbox"
                    role="switch"
                    className="set-switch"
                    checked={appearance.transparency}
                    onChange={(e) => updateAppearance({ transparency: e.target.checked })}
                  />
                </label>
                <label className="set-field set-switch-row">
                  <div className="set-field-text">
                    <div className="set-field-label">{t('appearanceConfirmDelete')}</div>
                    <div className="set-field-value">{t('appearanceConfirmDeleteHint')}</div>
                  </div>
                  <input
                    type="checkbox"
                    role="switch"
                    className="set-switch"
                    checked={appearance.confirmDelete}
                    onChange={(e) => updateAppearance({ confirmDelete: e.target.checked })}
                  />
                </label>
                <div className="set-field set-reset-row">
                  <div className="set-field-text">
                    <div className="set-field-value">{t('appearanceScope')}</div>
                  </div>
                  <button className="set-btn" onClick={resetAppearance}>
                    {t('appearanceReset')}
                  </button>
                </div>
              </>
            )}
            {section === 'shortcuts' && (
              <>
                <h3 className="set-pane-title">{t('setSecShortcuts')}</h3>
                {(
                  [
                    ['scOpen', ['Enter']],
                    ['scNavigate', ['↑', '↓']],
                    ['scRename', ['F2']],
                    ['scDelete', ['Delete']],
                    ['scSelectAll', [modKey, 'A']],
                    ['searchFiles', [modKey, 'F']],
                    ['scRange', ['Shift', 'Click']],
                    ['scMenu', ['Right-click']],
                    ['scEscape', ['Esc']],
                  ] as const
                ).map(([labelKey, keys]) => (
                  <div className="set-field set-shortcut" key={labelKey}>
                    <div className="set-field-text">
                      <div className="set-field-label">{t(labelKey)}</div>
                    </div>
                    <span className="set-keys">
                      {keys.map((k) => (
                        <kbd key={k}>{k}</kbd>
                      ))}
                    </span>
                  </div>
                ))}
                <div className="set-field">
                  <div className="set-field-value">{t('shortcutsNote')}</div>
                </div>
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
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">
                      {t(
                        CHANNEL_OPTIONS.find((o) => o.value === channel)?.labelKey ??
                          'channelStable',
                      )}
                    </span>
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
                  </span>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )

  return typeof document !== 'undefined' ? createPortal(modal, document.body) : modal
}
