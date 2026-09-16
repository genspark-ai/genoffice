import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import {
  AI_CUSTOM_FONT_MAX_PX,
  AI_CUSTOM_FONT_MIN_PX,
  DEFAULT_AI_PANEL_PREFS,
  Dropdown,
  aiPanelFontPx,
  clampAiCustomFontSize,
} from '@genoffice/ui'
import type { AiFontSize, AiPanelPrefs } from '@genoffice/ui'
import {
  DEFAULT_MAX_OUTPUT_TOKENS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  clampMaxOutputTokens,
} from '@genoffice/ai-provider/browser'
import type {
  MarketplaceCategory,
  MarketplaceCategoryInfo,
  MarketplaceEntry,
  MarketplaceUploadEntry,
  MarketplaceUploadPayload,
  PiResourceReport,
  PluginEntry,
  PluginKind,
  SkillEntry,
  SkillKind,
} from '../../shared/home-api'
import type {
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiMediaSettings,
  AiSearchProviderMeta,
  AiSearchSettings,
  AiSettings,
} from '@genoffice/ai-provider'
import { useI18n } from './locale'
import type { StringKey, TFunc } from './locale'
import type {
  AccountStatus,
  AiCatalogEntry,
  ModuleEntry,
  ModuleKind,
  TranslateFileStatus,
  TranslationKbEntry,
  TranslationKbSchema,
  TranslationKbScope,
  UiTheme,
} from '../../shared/home-api'
import { ProviderLogo } from './provider-logos'
import { IntegrationsPane, skillUpdateDue } from './IntegrationsPane'
import './settings.css'

// ── Settings modal (opened from the account menu) ─────────
// Genspark-style two-pane dialog: section nav on the left, fields on the right.
// All values go through the existing home IPC; nothing is stored locally.

// sorted by ISO 639 language code — native-script labels have no natural
// shared alphabet, so the code is the ordering key
const LANG_OPTIONS = [
  { value: 'ar', label: 'العربية' },
  { value: 'cs', label: 'Čeština' },
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
const THEME_OPTIONS = [
  { value: 'system', labelKey: 'themeSystem' },
  { value: 'light', labelKey: 'themeLight' },
  { value: 'dark', labelKey: 'themeDark' },
] as const satisfies readonly { value: UiTheme; labelKey: StringKey }[]

const AI_FONT_SIZE_OPTIONS = [
  { value: 'default', labelKey: 'aiFontSizeDefault' },
  { value: 'large', labelKey: 'aiFontSizeLarge' },
  { value: 'xlarge', labelKey: 'aiFontSizeXLarge' },
  { value: 'custom', labelKey: 'aiFontSizeCustom' },
] as const satisfies readonly { value: AiFontSize; labelKey: StringKey }[]

const CHANNEL_OPTIONS = [
  { value: 'stable', labelKey: 'channelStable' },
  { value: 'beta', labelKey: 'channelBeta' },
] as const satisfies readonly { value: 'stable' | 'beta'; labelKey: StringKey }[]

/** GitHub-style abbreviated stargazer count (2591 → "2.6k") — the number is
 * social proof, not a metric; the cached/exact value would only look stale */
function formatStars(n: number): string {
  if (n < 1000) return String(n)
  const k = n / 1000
  return `${k >= 100 ? Math.round(k) : (Math.round(k * 10) / 10).toString().replace(/\.0$/, '')}k`
}

/** px stepper for the custom AI panel text size; in-range values apply live,
 * out-of-range or partial input is clamped on blur */
function CustomFontSizeInput({
  value,
  label,
  onCommit,
}: {
  value: number
  label: string
  onCommit: (px: number) => void
}) {
  const [draft, setDraft] = useState(String(value))
  const [editing, setEditing] = useState(false)
  const shown = editing ? draft : String(value)
  const commit = (raw: string) => {
    const px = clampAiCustomFontSize(raw)
    if (px !== null && px !== value) onCommit(px)
  }
  return (
    <label className="set-num">
      <input
        type="number"
        className="set-input set-num-input"
        aria-label={label}
        min={AI_CUSTOM_FONT_MIN_PX}
        max={AI_CUSTOM_FONT_MAX_PX}
        step={1}
        value={shown}
        onFocus={() => {
          setDraft(String(value))
          setEditing(true)
        }}
        onChange={(e) => {
          setDraft(e.target.value)
          const n = Number(e.target.value)
          if (Number.isInteger(n) && n >= AI_CUSTOM_FONT_MIN_PX && n <= AI_CUSTOM_FONT_MAX_PX) {
            onCommit(n)
          }
        }}
        onBlur={() => {
          commit(draft)
          setEditing(false)
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') e.currentTarget.blur()
        }}
      />
      <span className="set-num-unit">px</span>
    </label>
  )
}

type SectionId =
  | 'account'
  | 'aiModel'
  | 'aiMedia'
  | 'translationKb'
  | 'general'
  | 'integrations'
  | 'modules'
  | 'skillsPlugins'
  | 'about'

const SECTIONS: readonly { id: SectionId; labelKey: StringKey }[] = [
  { id: 'account', labelKey: 'setSecAccount' },
  { id: 'aiModel', labelKey: 'setSecAiModel' },
  { id: 'aiMedia', labelKey: 'setSecAiMedia' },
  { id: 'translationKb', labelKey: 'setSecTranslationKb' },
  { id: 'general', labelKey: 'setSecGeneral' },
  { id: 'integrations', labelKey: 'setSecIntegrations' },
  { id: 'modules', labelKey: 'setSecModules' },
  { id: 'skillsPlugins', labelKey: 'setSecSkillsPlugins' },
  { id: 'about', labelKey: 'setSecAbout' },
]

function SectionIcon({ id }: { id: SectionId }) {
  if (id === 'aiModel') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M8 1.8 9.5 6l4.2 1.5L9.5 9 8 13.2 6.5 9 2.3 7.5 6.5 6 8 1.8Z"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinejoin="round"
        />
        <path
          d="M12.8 11.2v3M11.3 12.7h3"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    )
  }
  if (id === 'translationKb') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M3 2.5h6.5A2 2 0 0 1 11.5 4.5v9H5a2 2 0 0 0-2 2v-13Z"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinejoin="round"
        />
        <path
          d="M13 4.5v7M5.2 5.6h4M5.2 8h4"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    )
  }
  if (id === 'aiMedia') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect x="2" y="3" width="12" height="10" rx="1.5" stroke="currentColor" strokeWidth="1.3" />
        <path
          d="M2.5 11.5 6 8l2.5 2.5L10.5 9l3 2.8"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
        <circle cx="10.5" cy="6" r="1.1" fill="currentColor" />
      </svg>
    )
  }
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
  }  if (id === 'skillsPlugins') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path d="M8 1.5 9.5 5l3.5 1L9.5 7 8 10.5 6.5 7 3 6l3.5-1L8 1.5Z" stroke="currentColor" strokeWidth="1.3" strokeLinejoin="round" />
        <path d="M3 12.5 4 11.5l1 1M13 12.5l-1-1-1 1" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
        <path d="M8 14v-2.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      </svg>
    )
  }

  if (id === 'integrations') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M5.5 2v3M10.5 2v3M4 5h8v2.5a4 4 0 0 1-8 0V5ZM8 11.5V14"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          strokeLinejoin="round"
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
  if (id === 'modules') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect
          x="1.8"
          y="1.8"
          width="5.4"
          height="5.4"
          rx="1"
          stroke="currentColor"
          strokeWidth="1.3"
        />
        <rect
          x="8.8"
          y="1.8"
          width="5.4"
          height="5.4"
          rx="1"
          stroke="currentColor"
          strokeWidth="1.3"
        />
        <rect
          x="1.8"
          y="8.8"
          width="5.4"
          height="5.4"
          rx="1"
          stroke="currentColor"
          strokeWidth="1.3"
        />
        <rect
          x="8.8"
          y="8.8"
          width="5.4"
          height="5.4"
          rx="1"
          stroke="currentColor"
          strokeWidth="1.3"
        />
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

/** AI model pane: provider / model / key / base URL, saved to userData/ai-settings.json */
function AiModelPane({ t }: { t: TFunc }) {
  const [catalog, setCatalog] = useState<AiCatalogEntry[]>(
    () => window.aiOffice.getAiProviders?.() ?? [],
  )
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saved, setSaved] = useState(false)
  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState<{ ok: boolean; error?: string } | null>(null)
  const [capabilities, setCapabilities] = useState<{
    search: { available: boolean; via: string; configured: boolean; fallback?: string }
    image_search: { available: boolean; via: string; configured: boolean; fallback?: string }
    image_generation: { available: boolean; via: string; configured: boolean }
    media_analysis: { available: boolean; via: string; configured: boolean }
  } | null>(null)

  useEffect(() => {
    let alive = true
    void window.aiOffice.getAiCapabilities?.().then((c) => {
      if (alive && c) setCapabilities(c.capabilities)
    })
    return () => { alive = false }
  }, [])
  /** free-typed value of the output-cap field; committed (and clamped) on blur */
  const [maxTokensDraft, setMaxTokensDraft] = useState<string | null>(null)

  const refreshCodexModels = useCallback(async (cliPath = '', selectedModel = '') => {
    if (!window.aiOffice.getCodexModels) return
    const live = await window.aiOffice.getCodexModels(cliPath)
    setCatalog((current) =>
      current.map((entry) => {
        if (entry.id !== 'codex') return entry
        const models =
          selectedModel && !live.models.includes(selectedModel)
            ? [selectedModel, ...live.models]
            : live.models
        return { ...entry, models, defaultModel: live.defaultModel }
      }),
    )
  }, [])

  useEffect(() => {
    let alive = true
    void window.aiOffice.getAiSettings?.().then((s) => {
      if (!alive || !s) return
      // The switch is disabled with genspark, so never present it stranded
      // off. Display-only: s.provider may be the activeProvider fallback for
      // a half-configured BYOK selection, so writing anything back here would
      // clobber the stored choice — the main process heals a genuine legacy
      // genspark+off file itself, judged on the raw stored provider.
      if (s.provider === 'genspark' && s.gskToolsEnabled === false) {
        s = { ...s, gskToolsEnabled: true }
      }
      setSettings(s)
      const codex = s.providers.codex
      if (codex) {
        void refreshCodexModels(codex.cliPath ?? '', codex.model).catch(() => undefined)
      }
    })
    return () => {
      alive = false
    }
  }, [refreshCodexModels])

  if (!settings) return null
  const provider = settings.provider
  const meta = catalog.find((c) => c.id === provider)
  const config = settings.providers[provider] ?? {
    apiKey: '',
    model: meta?.defaultModel ?? '',
    baseUrl: undefined,
    cliPath: undefined,
  }
  const isGenspark = provider === 'genspark'
  const isCodex = provider === 'codex'

  const touch = () => {
    setDirty(true)
    setSaved(false)
    setTestResult(null)
  }
  const updateConfig = (patch: Partial<typeof config>) => {
    setSettings({
      ...settings,
      providers: { ...settings.providers, [provider]: { ...config, ...patch } },
    })
    touch()
  }
  /** Commit the output-cap input: clamp what was typed and drop a no-op edit */
  const commitMaxTokens = () => {
    if (maxTokensDraft === null) return
    setMaxTokensDraft(null)
    const next = clampMaxOutputTokens(Number.parseInt(maxTokensDraft, 10))
    if (next === (settings.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS)) return
    setSettings({ ...settings, maxOutputTokens: next })
    touch()
  }
  const selectProvider = (id: AiSettings['provider']) => {
    // cloud tools cannot be off with genspark (chat runs through gsk anyway)
    setSettings({
      ...settings,
      provider: id,
      ...(id === 'genspark' ? { gskToolsEnabled: true } : {}),
    })
    touch()
  }
  const save = () => {
    window.aiOffice
      .setAiSettings?.(settings)
      .then(() => {
        setDirty(false)
        setSaved(true)
      })
      .catch((error) => {
        window.alert(error instanceof Error ? error.message : String(error))
      })
  }
  const test = () => {
    setTesting(true)
    setTestResult(null)
    window.aiOffice
      .testAiSettings?.(settings)
      .then((r) => {
        setTestResult(r ?? { ok: false })
        if (r?.ok && isCodex) {
          void refreshCodexModels(config.cliPath ?? '', config.model).catch(() => undefined)
        }
      })
      .catch((error) =>
        setTestResult({ ok: false, error: error instanceof Error ? error.message : String(error) }),
      )
      .finally(() => setTesting(false))
  }

  return (
    <>
      <h3 className="set-pane-title">{t('setSecAiModel')}</h3>
      <div className="set-field">
        <div className="set-field-text">
          <label className="set-field-label">{t('setAiProvider')}</label>
        </div>
        <Dropdown
          className="set-dd"
          value={provider}
          ariaLabel={t('setAiProvider')}
          options={catalog.map((c) => ({
            value: c.id,
            label: c.label,
            render: (
              <>
                <ProviderLogo id={c.id} />
                {c.label}
              </>
            ),
          }))}
          onPick={(v) => selectProvider(v as AiSettings['provider'])}
        />
      </div>
      <div className="set-field-desc set-ai-note">
        {isGenspark ? t('setAiGensparkHint') : isCodex ? t('setAiCodexHint') : t('setAiByokNote')}
      </div>
      <div className="set-field">
        <div className="set-field-text">
          <label className="set-field-label">{t('setAiModelId')}</label>
        </div>
        {meta && meta.models.length > 0 ? (
          <Dropdown
            className="set-dd"
            value={config.model || meta.defaultModel}
            ariaLabel={t('setAiModelId')}
            options={meta.models.map((m) => ({ value: m, label: m }))}
            onPick={(m) => updateConfig({ model: m })}
          />
        ) : (
          <input
            id="set-ai-model"
            className="set-input"
            type="text"
            value={config.model}
            placeholder="model-id"
            spellCheck={false}
            onChange={(e) => updateConfig({ model: e.target.value })}
          />
        )}
      </div>
      {isCodex ? (
        <div className="set-field">
          <div className="set-field-text">
            <div className="set-field-stack">
              <label className="set-field-label" htmlFor="set-ai-cli-path">
                {t('setAiCodexPath')}
              </label>
              <div className="set-field-desc">{t('setAiCodexPathHint')}</div>
            </div>
          </div>
          <input
            id="set-ai-cli-path"
            className="set-input"
            type="text"
            value={config.cliPath ?? ''}
            placeholder={t('setAiCodexAutoPlaceholder')}
            spellCheck={false}
            autoComplete="off"
            onChange={(e) => updateConfig({ cliPath: e.target.value.trim() })}
            onBlur={(e) => {
              const cliPath = e.target.value.trim()
              void refreshCodexModels(cliPath, config.model).catch(() => undefined)
            }}
          />
        </div>
      ) : !isGenspark ? (
        <>
          <div className="set-field">
            <div className="set-field-text">
              <div className="set-field-stack">
                <label className="set-field-label" htmlFor="set-ai-key">
                  {t('setAiApiKey')}
                </label>
                <div className="set-field-desc">{t('setAiKeyHint')}</div>
              </div>
            </div>
            <input
              id="set-ai-key"
              className="set-input"
              type="password"
              value={config.apiKey}
              placeholder={meta?.keyPlaceholder ?? 'API Key'}
              spellCheck={false}
              autoComplete="off"
              onChange={(e) => updateConfig({ apiKey: e.target.value.trim() })}
            />
          </div>
          <div className="set-field">
            <div className="set-field-text">
              <div className="set-field-stack">
                <label className="set-field-label" htmlFor="set-ai-base-url">
                  {t('setAiBaseUrl')}
                </label>
                {!meta?.needsBaseUrl && (
                  <div className="set-field-desc">{t('setAiBaseUrlHint')}</div>
                )}
              </div>
            </div>
            <input
              id="set-ai-base-url"
              className="set-input"
              type="text"
              value={config.baseUrl ?? ''}
              placeholder={meta?.needsBaseUrl ? 'https://…/v1' : meta?.defaultBaseUrl}
              spellCheck={false}
              onChange={(e) => updateConfig({ baseUrl: e.target.value.trim() })}
            />
          </div>
        </>
      ) : null}
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <label className="set-field-label" htmlFor="set-ai-max-tokens">
              {t('setAiMaxTokens')}
            </label>
            <div className="set-field-desc">{t('setAiMaxTokensDesc')}</div>
          </div>
        </div>
        <input
          id="set-ai-max-tokens"
          className="set-input"
          type="number"
          min={MIN_MAX_OUTPUT_TOKENS}
          max={MAX_MAX_OUTPUT_TOKENS}
          step={1024}
          value={maxTokensDraft ?? String(settings.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS)}
          onChange={(e) => setMaxTokensDraft(e.target.value)}
          onBlur={commitMaxTokens}
        />
      </div>
      <div className="set-field">
        <div className="set-field-text">
          <div className="set-field-stack">
            <div className="set-field-label">{t('setAiGskTools')}</div>
            <div className="set-field-desc">{t('setAiGskToolsDesc')}</div>
          </div>
        </div>
        {/* locked on with the genspark provider — chat runs through gsk anyway */}
        <button
          className="set-switch"
          role="switch"
          aria-checked={settings.gskToolsEnabled !== false}
          aria-label={t('setAiGskTools')}
          disabled={isGenspark}
          onClick={() => {
            setSettings({ ...settings, gskToolsEnabled: settings.gskToolsEnabled === false })
            touch()
          }}
        />
      </div>
      <div className="set-pane-footer">
        <AiStatusPill
          status={
            testing
              ? { kind: 'testing', text: t('setAiTesting') }
              : testResult
                ? testResult.ok
                  ? { kind: 'ok', text: t('setAiTestOk') }
                  : { kind: 'err', text: testResult.error || t('setAiTestFail') }
                : saved
                  ? { kind: 'ok', text: t('setAiSaved') }
                  : null
          }
        />
        <button className="set-btn" disabled={testing} onClick={test}>
          {t('setAiTest')}
        </button>
        <button className="set-btn primary" disabled={!dirty} onClick={save}>
          {t('setAiSave')}
        </button>
      </div>
    </>
  )
}

type Capability = 'image' | 'analysis' | 'video' | 'search'

/**
 * AI media & search pane, one block per capability — web search, image
 * generation, image analysis, video analysis — each with the same
 * provider / model / key / base URL rows as the AI Model pane. A vendor's key
 * and base URL are stored once and shared by every block that picks it.
 * Saved into the same ai-settings.json as the chat provider.
 */
function AiMediaPane({ t }: { t: TFunc }) {
  const [mediaCatalog] = useState<AiMediaProviderMeta[]>(
    () => window.aiOffice.getAiMediaProviders?.() ?? [],
  )
  const [searchCatalog] = useState<AiSearchProviderMeta[]>(
    () => window.aiOffice.getAiSearchProviders?.() ?? [],
  )
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saved, setSaved] = useState(false)
  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState<{ ok: boolean; error?: string } | null>(null)
  const [capabilities, setCapabilities] = useState<{
    search: { available: boolean; via: string; configured: boolean; fallback?: string }
    image_search: { available: boolean; via: string; configured: boolean; fallback?: string }
    image_generation: { available: boolean; via: string; configured: boolean }
    media_analysis: { available: boolean; via: string; configured: boolean }
  } | null>(null)

  useEffect(() => {
    let alive = true
    void window.aiOffice.getAiCapabilities?.().then((c) => {
      if (alive && c) setCapabilities(c.capabilities)
    })
    return () => { alive = false }
  }, [])

  useEffect(() => {
    let alive = true
    void window.aiOffice.getAiSettings?.().then((s) => {
      if (alive && s) setSettings(s)
    })
    return () => {
      alive = false
    }
  }, [])

  if (!settings?.media || !settings.search) return null
  const media: AiMediaSettings = settings.media
  const search: AiSearchSettings = settings.search

  const touch = () => {
    setDirty(true)
    setSaved(false)
    setTestResult(null)
  }
  const setMedia = (next: AiMediaSettings) => {
    setSettings({ ...settings, media: next })
    touch()
  }
  const setSearch = (next: AiSearchSettings) => {
    setSettings({ ...settings, search: next })
    touch()
  }
  const mediaConfigOf = (id: AiMediaProviderId) => {
    const meta = mediaCatalog.find((m) => m.id === id)
    return (
      media.providers[id] ?? {
        apiKey: '',
        imageModel: meta?.defaultImageModel ?? '',
        analysisModel: meta?.defaultAnalysisModel ?? '',
      }
    )
  }
  const updateMediaConfig = (
    id: AiMediaProviderId,
    patch: Partial<AiMediaSettings['providers'][AiMediaProviderId]>,
  ) =>
    setMedia({
      ...media,
      providers: { ...media.providers, [id]: { ...mediaConfigOf(id), ...patch } },
    })

  const save = () => {
    window.aiOffice
      .setAiSettings?.(settings)
      .then(() => {
        setDirty(false)
        setSaved(true)
      })
      .catch((error) => {
        window.alert(error instanceof Error ? error.message : String(error))
      })
  }
  // every distinct BYOK vendor the four blocks point at is checked once; first failure wins
  const test = async () => {
    setTesting(true)
    setTestResult(null)
    try {
      const vendors = new Set<AiMediaProviderId>(
        [media.imageProvider, media.analysisProvider, media.videoAnalysisProvider].filter(
          (id) => id !== 'genspark',
        ),
      )
      const checks: Promise<{ ok: boolean; error?: string } | undefined>[] = [...vendors].map(
        (id) =>
          window.aiOffice.testAiMediaSettings?.({ provider: id, config: mediaConfigOf(id) }) ??
          Promise.resolve(undefined),
      )
      if (search.provider !== 'genspark') {
        checks.push(
          window.aiOffice.testAiSearchSettings?.({
            provider: search.provider,
            apiKey: search.providers[search.provider]?.apiKey ?? '',
          }) ?? Promise.resolve(undefined),
        )
      }
      if (checks.length === 0) {
        checks.push(
          window.aiOffice.testAiMediaSettings?.({
            provider: 'genspark',
            config: mediaConfigOf('genspark'),
          }) ?? Promise.resolve(undefined),
        )
      }
      const results = await Promise.all(checks)
      setTestResult(results.find((r) => r && !r.ok) ?? { ok: true })
    } catch (error) {
      setTestResult({ ok: false, error: error instanceof Error ? error.message : String(error) })
    } finally {
      setTesting(false)
    }
  }

  const providerRow = (
    label: string,
    value: string,
    options: { id: string; label: string }[],
    onPick: (id: string) => void,
  ) => (
    <div className="set-field">
      <div className="set-field-text">
        <label className="set-field-label">{t('setAiProvider')}</label>
      </div>
      <Dropdown
        className="set-dd"
        value={value}
        ariaLabel={label}
        options={options.map((c) => ({
          value: c.id,
          label: c.label,
          render: (
            <>
              <ProviderLogo id={c.id} />
              {c.label}
            </>
          ),
        }))}
        onPick={onPick}
      />
    </div>
  )

  const modelRow = (
    id: string,
    models: string[],
    fallback: string,
    value: string,
    onChange: (v: string) => void,
  ) => (
    <div className="set-field">
      <div className="set-field-text">
        <label className="set-field-label" htmlFor={id}>
          {t('setAiModelId')}
        </label>
      </div>
      {models.length > 0 ? (
        <Dropdown
          className="set-dd"
          value={value || fallback}
          ariaLabel={t('setAiModelId')}
          options={models.map((m) => ({ value: m, label: m }))}
          onPick={onChange}
        />
      ) : (
        <input
          id={id}
          className="set-input"
          type="text"
          value={value}
          placeholder="model-id"
          spellCheck={false}
          onChange={(e) => onChange(e.target.value)}
        />
      )}
    </div>
  )

  const keyRow = (
    id: string,
    value: string,
    placeholder: string,
    onChange: (v: string) => void,
  ) => (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-stack">
          <label className="set-field-label" htmlFor={id}>
            {t('setAiApiKey')}
          </label>
          <div className="set-field-desc">{t('setAiKeyHint')}</div>
        </div>
      </div>
      <input
        id={id}
        className="set-input"
        type="password"
        value={value}
        placeholder={placeholder}
        spellCheck={false}
        autoComplete="off"
        onChange={(e) => onChange(e.target.value.trim())}
      />
    </div>
  )

  const baseUrlRow = (
    id: string,
    meta: AiMediaProviderMeta,
    value: string,
    onChange: (v: string) => void,
  ) => (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-stack">
          <label className="set-field-label" htmlFor={id}>
            {t('setAiBaseUrl')}
          </label>
          {!meta.needsBaseUrl && <div className="set-field-desc">{t('setAiBaseUrlHint')}</div>}
        </div>
      </div>
      <input
        id={id}
        className="set-input"
        type="text"
        value={value}
        placeholder={meta.needsBaseUrl ? 'https://…/v1' : meta.defaultBaseUrl}
        spellCheck={false}
        onChange={(e) => onChange(e.target.value.trim())}
      />
    </div>
  )

  /** one media block: provider → model → key → base URL (key/base URL shared per vendor) */
  const mediaBlock = (cap: Exclude<Capability, 'search'>) => {
    const title =
      cap === 'image'
        ? t('setAiCapImage')
        : cap === 'analysis'
          ? t('setAiCapAnalysis')
          : t('setAiCapVideo')
    const options = mediaCatalog.filter((m) =>
      cap === 'image'
        ? !!m.imageProtocol
        : cap === 'video'
          ? !!m.analysisProtocol && m.videoAnalysis
          : !!m.analysisProtocol,
    )
    const current =
      cap === 'image'
        ? media.imageProvider
        : cap === 'video'
          ? media.videoAnalysisProvider
          : media.analysisProvider
    // Defensive: a build that ships without the media registry (or a future
    // capability the registry does not implement) used to crash the entire
    // settings modal here via `options[0]!.id`. Render a hint instead.
    const meta = options.find((m) => m.id === current) ?? options[0]
    if (!meta) {
      return (
        <section key={cap}>
          <h4 className="set-pane-subtitle">{title}</h4>
          <div className="set-field-desc set-ai-note">{t('setAiProviderCatalogEmpty')}</div>
        </section>
      )
    }
    const id = meta.id
    const config = mediaConfigOf(id)
    const pick = (next: string) => {
      const p = next as AiMediaProviderId
      setMedia(
        cap === 'image'
          ? { ...media, imageProvider: p }
          : cap === 'video'
            ? { ...media, videoAnalysisProvider: p }
            : { ...media, analysisProvider: p },
      )
    }
    const modelField = cap === 'image' ? 'imageModel' : 'analysisModel'
    return (
      <section key={cap}>
        <h4 className="set-pane-subtitle">{title}</h4>
        {providerRow(title, id, options, pick)}
        <div className="set-field-desc set-ai-note">
          {id === 'genspark' ? t('setAiMediaGensparkHint') : meta.description}
        </div>
        {id !== 'genspark' && (
          <>
            {modelRow(
              `set-ai-${cap}-model`,
              cap === 'image' ? meta.imageModels : meta.analysisModels,
              cap === 'image' ? meta.defaultImageModel : meta.defaultAnalysisModel,
              config[modelField],
              (m) => updateMediaConfig(id, { [modelField]: m }),
            )}
            {keyRow(`set-ai-${cap}-key`, config.apiKey, meta.keyPlaceholder, (v) =>
              updateMediaConfig(id, { apiKey: v }),
            )}
            {baseUrlRow(`set-ai-${cap}-base-url`, meta, config.baseUrl ?? '', (v) =>
              updateMediaConfig(id, { baseUrl: v }),
            )}
          </>
        )}
      </section>
    )
  }

  const searchMeta = searchCatalog.find((m) => m.id === search.provider)
  const searchKey =
    search.provider === 'genspark' ? '' : (search.providers[search.provider]?.apiKey ?? '')

  const capRow = (label: string, c: { available: boolean; via: string; configured: boolean; fallback?: string } | undefined) => {
    if (!c) return null
    const cls = c.available ? (c.configured ? 'is-on' : 'is-fallback') : 'is-off'
    return (
      <span className={`set-cap-pill ${cls}`} title={c.configured ? '' : 'via DuckDuckGo fallback'}>
        <span className="set-cap-dot" aria-hidden="true" />
        <span className="set-cap-label">{label}</span>
        <span className="set-cap-via">{c.via}</span>
      </span>
    )
  }

  return (
    <>
      <h3 className="set-pane-title">{t('setSecAiMedia')}</h3>
      <div className="set-field-desc set-ai-note">{t('setAiSharedKeyHint')}</div>
      {capabilities && (
        <div className="set-cap-status" data-ai-capability-status="1">
          {capRow(t('setAiCapSearch'), capabilities.search)}
          {capRow(t('setAiCapImageSearch'), capabilities.image_search)}
          {capRow(t('setAiCapImageGen'), capabilities.image_generation)}
          {capRow(t('setAiCapAnalysis'), capabilities.media_analysis)}
        </div>
      )}
      <section>
        <h4 className="set-pane-subtitle">{t('setAiCapSearch')}</h4>
        {providerRow(t('setAiCapSearch'), search.provider, searchCatalog, (v) =>
          setSearch({ ...search, provider: v as AiSearchSettings['provider'] }),
        )}
        <div className="set-field-desc set-ai-note">
          {search.provider === 'genspark'
            ? t('setAiSearchGensparkHint')
            : searchMeta?.imageSearch
              ? t('setAiSearchSerperHint')
              : t('setAiSearchTavilyHint')}
        </div>
        {search.provider !== 'genspark' &&
          keyRow('set-ai-search-key', searchKey, searchMeta?.keyPlaceholder ?? 'API Key', (v) =>
            setSearch({
              ...search,
              providers: { ...search.providers, [search.provider]: { apiKey: v } },
            }),
          )}
      </section>
      {mediaBlock('image')}
      {mediaBlock('analysis')}
      {mediaBlock('video')}
      <div className="set-pane-footer">
        <AiStatusPill
          status={
            testing
              ? { kind: 'testing', text: t('setAiTesting') }
              : testResult
                ? testResult.ok
                  ? { kind: 'ok', text: t('setAiTestOk') }
                  : { kind: 'err', text: testResult.error || t('setAiTestFail') }
                : saved
                  ? { kind: 'ok', text: t('setAiSaved') }
                  : null
          }
        />
        <button className="set-btn" disabled={testing} onClick={() => void test()}>
          {t('setAiTest')}
        </button>
        <button className="set-btn primary" disabled={!dirty} onClick={save}>
          {t('setAiSave')}
        </button>
      </div>
    </>
  )
}

interface AiStatus {
  kind: 'testing' | 'ok' | 'err'
  text: string
}

/** colored feedback pill in the AI pane footer: spinner while testing, then success/error */
/* ── translation knowledge base ─────────────────────────────────────
 * Settings → AI → Translation KB. Five schemas (mirroring the
 * `trade.translation.*` set the upstream LumosAI translate-config skill
 * writes), each a CRUD list backed by `~/.genoffice/translation-kb.json`
 * through the ai:translation-kb-* channels.
 *
 * The same page also drives the file pipeline: pick a document, let the KB +
 * the active provider mine a `--dictionary`, then optionally translate the
 * file in one go.
 */

const TKB_SCHEMAS: readonly { key: TranslationKbSchema; labelKey: StringKey }[] = [
  { key: 'term', labelKey: 'tkbSchemaTerm' },
  { key: 'forbidden', labelKey: 'tkbSchemaForbidden' },
  { key: 'brand', labelKey: 'tkbSchemaBrand' },
  { key: 'styleRule', labelKey: 'tkbSchemaStyle' },
  { key: 'customerPreference', labelKey: 'tkbSchemaCustomer' },
]

const TKB_SCOPES: readonly TranslationKbScope[] = [
  'session',
  'customer',
  'project',
  'company',
  'global',
]

/** Field ids double as the KB entry property names they fill in. */
const TKB_FIELDS: Record<TranslationKbSchema, readonly { id: string; labelKey: StringKey }[]> = {
  term: [
    { id: 'sourceTerm', labelKey: 'tkbFieldSource' },
    { id: 'targetTerm', labelKey: 'tkbFieldTarget' },
  ],
  forbidden: [
    { id: 'forbiddenText', labelKey: 'tkbFieldForbidden' },
    { id: 'replacement', labelKey: 'tkbFieldReplacement' },
  ],
  brand: [
    { id: 'word', labelKey: 'tkbFieldWord' },
    { id: 'translateAs', labelKey: 'tkbFieldTranslateAs' },
  ],
  styleRule: [
    { id: 'name', labelKey: 'tkbFieldName' },
    { id: 'description', labelKey: 'tkbFieldDesc' },
  ],
  customerPreference: [
    { id: 'customerName', labelKey: 'tkbCustomer' },
    { id: 'preferenceType', labelKey: 'tkbFieldPrefType' },
    { id: 'value', labelKey: 'tkbFieldPrefValue' },
  ],
}

const TKB_BRAND_POLICIES = ['neverTranslate', 'keep', 'translateAs'] as const

/** Language tags offered to the dictionary builder; labels stay native so the
 *  picker reads the same in every UI locale. */
const TKB_LANGS: readonly { value: string; label: string }[] = [
  { value: 'auto', label: 'Auto' },
  { value: 'zh-CN', label: '中文（简体）' },
  { value: 'zh-TW', label: '中文（繁體）' },
  { value: 'en-US', label: 'English' },
  { value: 'ja-JP', label: '日本語' },
  { value: 'ko-KR', label: '한국어' },
  { value: 'de-DE', label: 'Deutsch' },
  { value: 'fr-FR', label: 'Français' },
  { value: 'es-ES', label: 'Español' },
  { value: 'pt-BR', label: 'Português' },
  { value: 'it-IT', label: 'Italiano' },
  { value: 'ru-RU', label: 'Русский' },
  { value: 'ar-SA', label: 'العربية' },
  { value: 'th-TH', label: 'ไทย' },
  { value: 'vi-VN', label: 'Tiếng Việt' },
  { value: 'id-ID', label: 'Bahasa Indonesia' },
]

/**
 * The KB is a bag of rows discriminated by their own field names, not by a tag
 * — `kind` never lands on disk, so it is derived here for the UI to group by.
 */
function tkbKindOf(entry: TranslationKbEntry): TranslationKbSchema {
  if (typeof entry.sourceTerm === 'string') return 'term'
  if (typeof entry.forbiddenText === 'string') return 'forbidden'
  if (typeof entry.word === 'string') return 'brand'
  if (typeof entry.name === 'string') return 'styleRule'
  return 'customerPreference'
}

/** Short source → target pair shown on a row, whatever the schema. */
interface DictSegment { source: string; target: string; origin: 'kb' | 'llm' }
interface DictResult {
  dictionaryPath: string
  kbEntries: number
  llmEntries: number
  missed: number
  totalSegments: number
  elapsedMs: number
  segments: DictSegment[]
}

function tkbRowText(entry: TranslationKbEntry): { primary: string; secondary: string } {
  switch (tkbKindOf(entry)) {
    case 'term':
      return { primary: entry.sourceTerm ?? '', secondary: entry.targetTerm ?? '' }
    case 'forbidden':
      return { primary: entry.forbiddenText ?? '', secondary: entry.replacement ?? '' }
    case 'brand':
      return {
        primary: entry.word ?? '',
        secondary: entry.translateAs ?? entry.policy ?? '',
      }
    case 'styleRule':
      return { primary: entry.name ?? '', secondary: entry.description ?? '' }
    default:
      return {
        primary: entry.customerName ?? '',
        secondary: [entry.preferenceType, entry.value].filter(Boolean).join(' = '),
      }
  }
}

function TranslationKbPane({ t }: { t: TFunc }) {
  const [entries, setEntries] = useState<TranslationKbEntry[]>([])
  const [schema, setSchema] = useState<TranslationKbSchema>('term')
  const [draft, setDraft] = useState<Record<string, string>>({})
  const [scope, setScope] = useState<TranslationKbScope>('company')
  const [priority, setPriority] = useState(50)
  const [policy, setPolicy] = useState<(typeof TKB_BRAND_POLICIES)[number]>('neverTranslate')
  const [flash, setFlash] = useState(false)
  const [error, setError] = useState('')
  const [aiSettings, setAiSettings] = useState<AiSettings | null>(null)
  const [gskEnabled, setGskEnabled] = useState(false)

  const [filePath, setFilePath] = useState('')
  const [sourceLang, setSourceLang] = useState('auto')
  const [targetLang, setTargetLang] = useState('en-US')
  const [customerName, setCustomerName] = useState('')
  const [busy, setBusy] = useState(false)
  const [kbOnly, setKbOnly] = useState(false)
  const [dictError, setDictError] = useState('')
  const [outputPath, setOutputPath] = useState('')
  const [dict, setDict] = useState<DictResult | null>(null)
  const [status, setStatus] = useState<TranslateFileStatus | null>(null)
  const [savingDict, setSavingDict] = useState(false)
  const [savedDictCount, setSavedDictCount] = useState(0)

  const [snippet, setSnippet] = useState('')
  const [snippetBusy, setSnippetBusy] = useState(false)
  const [snippetResult, setSnippetResult] = useState<{
    translation: string
    status: 'translated' | 'memory-hit' | 'failed'
    matchedTerms: string[]
    elapsedMs: number
  } | null>(null)
  const [snippetError, setSnippetError] = useState('')
  const [snippetCopied, setSnippetCopied] = useState(false)

  const reload = useCallback(async () => {
    const result = await window.aiOffice.listTranslationKb?.()
    setEntries(result?.entries ?? [])
  }, [])

  useEffect(() => {
    void reload()
    void window.aiOffice.getTranslateFileStatus?.().then((s) => s && setStatus(s))
    void window.aiOffice.getAiSettings?.().then((s) => s && setAiSettings(s))
    void window.aiOffice.getAiCapabilities?.().then((c) => {
      if (c) setGskEnabled(c.gskToolsEnabled)
    })
    // Restore the last-used target/source/customer so the user does not have
    // to re-pick them every time the modal opens.
    try {
      const raw = localStorage.getItem('genoffice:tkb-prefs')
      if (raw) {
        const p = JSON.parse(raw) as Partial<{
          sourceLang: string
          targetLang: string
          customerName: string
          kbOnly: boolean
        }>
        if (p.sourceLang) setSourceLang(p.sourceLang)
        if (p.targetLang) setTargetLang(p.targetLang)
        if (typeof p.customerName === 'string') setCustomerName(p.customerName)
        if (typeof p.kbOnly === 'boolean') setKbOnly(p.kbOnly)
      }
    } catch {
      /* ignore corrupt local cache */
    }
  }, [reload])

  useEffect(() => {
    try {
      localStorage.setItem(
        'genoffice:tkb-prefs',
        JSON.stringify({ sourceLang, targetLang, customerName, kbOnly }),
      )
    } catch {
      /* private mode / quota — non-fatal */
    }
  }, [sourceLang, targetLang, customerName, kbOnly])

  useEffect(() => {
    if (!flash) return
    const id = window.setTimeout(() => setFlash(false), 1600)
    return () => window.clearTimeout(id)
  }, [flash])

  const add = async () => {
    const get = (id: string) => (draft[id] ?? '').trim()
    const base = {
      id: `${schema}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      scope,
      priority,
    }
    let entry: TranslationKbEntry | null = null
    if (schema === 'term') {
      if (!get('sourceTerm') || !get('targetTerm')) return
      entry = {
        ...base,
        sourceTerm: get('sourceTerm'),
        targetTerm: get('targetTerm'),
        ...(sourceLang !== 'auto' ? { sourceLang } : {}),
        targetLang,
      }
    } else if (schema === 'forbidden') {
      if (!get('forbiddenText')) return
      entry = {
        ...base,
        forbiddenText: get('forbiddenText'),
        ...(get('replacement') ? { replacement: get('replacement') } : {}),
      }
    } else if (schema === 'brand') {
      if (!get('word')) return
      entry = {
        ...base,
        word: get('word'),
        policy,
        ...(get('translateAs') ? { translateAs: get('translateAs') } : {}),
      }
    } else if (schema === 'styleRule') {
      if (!get('name') || !get('description')) return
      entry = { ...base, name: get('name'), description: get('description') }
    } else {
      if (!get('customerName') || !get('value')) return
      entry = {
        ...base,
        customerName: get('customerName'),
        preferenceType: get('preferenceType') || 'style',
        value: get('value'),
        targetLang,
      }
    }
    const result = await window.aiOffice.upsertTranslationKb?.(entry)
    if (!result?.ok) {
      setError(result?.error ?? t('tkbError', { error: 'upsert' }))
      return
    }
    setError('')
    setDraft({})
    setFlash(true)
    await reload()
  }

  const remove = async (id: string) => {
    await window.aiOffice.removeTranslationKb?.(id)
    await reload()
  }

  const pickFile = async () => {
    const picked = await window.aiOffice.pickTranslationFile?.(t('tkbDictTitle'))
    if (picked?.ok && picked.path) setFilePath(picked.path)
  }

  /** Save the LLM-extracted segments from the last dictionary build into the KB
   *  so future translations automatically pick them up. KB-only segments are
   *  skipped because they already came from the KB. */
  const saveDictToKb = async () => {
    if (!dict) return
    const llmOnly = dict.segments.filter((s) => s.origin === 'llm' && s.source && s.target)
    if (llmOnly.length === 0) return
    setSavingDict(true)
    let ok = 0
    for (const seg of llmOnly) {
      const id = `llm-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`
      const result = await window.aiOffice.upsertTranslationKb?.({
        id,
        scope: 'company',
        priority: 30,
        sourceTerm: seg.source,
        targetTerm: seg.target,
        targetLang,
      })
      if (result?.ok) ok++
    }
    setSavedDictCount(ok)
    setSavingDict(false)
    setFlash(true)
    await reload()
  }

  const build = async (thenTranslate: boolean) => {
    if (!filePath) {
      setDictError(t('tkbNoFile'))
      return
    }
    setBusy(true)
    setDictError('')
    setDict(null)
    setOutputPath('')
    setSavedDictCount(0)
    const common = {
      inputPath: filePath,
      sourceLang,
      targetLang,
      ...(customerName.trim() ? { customerName: customerName.trim() } : {}),
    }
    try {
      if (thenTranslate) {
        const result = await window.aiOffice.translateFileAuto?.(common)
        if (!result?.ok) {
          setDictError(t('tkbError', { error: result?.error ?? 'translate' }))
          return
        }
        setDict({
          dictionaryPath: result.dictionaryPath ?? '',
          kbEntries: result.dictionary?.kbEntries ?? 0,
          llmEntries: result.dictionary?.llmEntries ?? 0,
          missed: result.dictionary?.missed?.length ?? 0,
          totalSegments: result.dictionary?.totalSegments ?? 0,
          elapsedMs: result.dictionary?.elapsedMs ?? 0,
          segments: result.dictionary?.segments ?? [],
        })
        if (result.outputPath) {
          setOutputPath(result.outputPath)
          await reload()
        }
        return
      }
      const result = await window.aiOffice.buildTranslationDictionary?.({
        ...common,
        ...(kbOnly ? { useLlm: false } : {}),
      })
      if (!result?.ok) {
        setDictError(t('tkbError', { error: result?.error ?? 'dictionary' }))
        return
      }
      setDict({
        dictionaryPath: result.dictionaryPath ?? '',
        kbEntries: result.kbEntries ?? 0,
        llmEntries: result.llmEntries ?? 0,
        missed: result.missed?.length ?? 0,
        totalSegments: result.totalSegments ?? 0,
        elapsedMs: result.elapsedMs ?? 0,
        segments: result.segments ?? [],
      })
    } catch (err) {
      setDictError(t('tkbError', { error: err instanceof Error ? err.message : String(err) }))
    } finally {
      setBusy(false)
    }
  }

  const translateSnippet = async () => {
    if (!snippet.trim()) return
    setSnippetBusy(true)
    setSnippetError('')
    setSnippetResult(null)
    setSnippetCopied(false)
    try {
      const result = await window.aiOffice.translateSnippet?.({
        text: snippet.trim(),
        sourceLang,
        targetLang,
        ...(customerName.trim() ? { customerName: customerName.trim() } : {}),
      })
      if (!result?.ok) {
        setSnippetError(t('tkbSnippetError', { error: result?.error ?? 'translate' }))
        return
      }
      setSnippetResult({
        translation: result.translation ?? '',
        status: result.status ?? 'translated',
        matchedTerms: result.matchedTerms ?? [],
        elapsedMs: result.elapsedMs ?? 0,
      })
    } catch (err) {
      setSnippetError(
        t('tkbSnippetError', { error: err instanceof Error ? err.message : String(err) }),
      )
    } finally {
      setSnippetBusy(false)
    }
  }

  const swapLanguages = () => {
    // Snippet + dictionary share these targets; swapping both in one click
    // is the standard translator UX.
    setSourceLang(targetLang)
    setTargetLang(sourceLang)
  }

  const copySnippet = async () => {
    if (!snippetResult) return
    try {
      await navigator.clipboard.writeText(snippetResult.translation)
      setSnippetCopied(true)
      window.setTimeout(() => setSnippetCopied(false), 1500)
    } catch {
      // clipboard denied in non-secure context — non-fatal
    }
  }

  const counts = useMemo(() => {
    const out: Partial<Record<TranslationKbSchema, number>> = {}
    for (const entry of entries) {
      const kind = tkbKindOf(entry)
      out[kind] = (out[kind] ?? 0) + 1
    }
    return out
  }, [entries])

  const visible = entries.filter((e) => tkbKindOf(e) === schema)
  const fields = TKB_FIELDS[schema]
  const skillMissing = status !== null && !status.available
  const providerId = aiSettings?.provider ?? '—'
  // genspark is the zero-config default — it counts as configured whenever
  // gsk tools are enabled (i.e. the user is signed in to genspark).
  const hasApiKey = !!aiSettings?.providers?.[aiSettings.provider]?.apiKey
  const providerConfigured = !!aiSettings && (providerId === 'genspark' ? gskEnabled : hasApiKey)
  const providerLabel = providerId

  return (
    <>
      <h3 className="set-pane-title">{t('setSecTranslationKb')}</h3>
      <div className="set-field-desc set-ai-note">{t('tkbDesc')}</div>

      <div className="set-cap-status set-tkb-status" data-tkb-status="1">
        <span
          className={`set-cap-pill ${status && status.available ? 'is-on' : 'is-off'}`}
          title={status ? `${status.source} · ${status.skillDir}` : ''}
        >
          <span className="set-cap-dot" aria-hidden="true" />
          <span className="set-cap-label">translate</span>
          <span className="set-cap-via">{status?.source ?? '…'}</span>
        </span>
        {status && (
          <span className="set-cap-pill is-on" title={status.pythonPath}>
            <span className="set-cap-dot" aria-hidden="true" />
            <span className="set-cap-label">python</span>
            <span className="set-cap-via">{status.pythonPath.split('/').pop()}</span>
          </span>
        )}
        <span
          className={`set-cap-pill ${providerConfigured ? 'is-on' : 'is-fallback'}`}
          title={t('tkbProvider')}
        >
          <span className="set-cap-dot" aria-hidden="true" />
          <span className="set-cap-label">{t('tkbProvider')}</span>
          <span className="set-cap-via">{providerLabel}</span>
        </span>
        {!providerConfigured && (
          <span className="set-cap-pill is-off" title={t('tkbNoProvider')}>
            <span className="set-cap-dot" aria-hidden="true" />
            <span className="set-cap-label">{t('tkbNoProvider')}</span>
          </span>
        )}
      </div>

      <div className="set-mp-chips" role="tablist" aria-label={t('setSecTranslationKb')}>
        {TKB_SCHEMAS.map((s) => (
          <button
            key={s.key}
            type="button"
            role="tab"
            aria-selected={schema === s.key}
            className={`set-mp-chip${schema === s.key ? ' set-mp-chip-active' : ''}`}
            onClick={() => {
              setSchema(s.key)
              setDraft({})
            }}
          >
            {t(s.labelKey)}
            <span className="set-mp-chip-count">{counts[s.key] ?? 0}</span>
          </button>
        ))}
        <span className="set-tkb-total">{t('tkbTotal', { total: entries.length })}</span>
      </div>

      <div className="set-tkb-list">
        {visible.length === 0 ? (
          <div className="set-tkb-empty">{t('tkbEmpty')}</div>
        ) : (
          visible.map((entry) => {
            const row = tkbRowText(entry)
            return (
              <div className="set-tkb-row" key={entry.id}>
                <span className="set-tkb-primary" title={row.primary}>
                  {row.primary}
                </span>
                <span className="set-tkb-arrow" aria-hidden="true">
                  →
                </span>
                <span className="set-tkb-secondary" title={row.secondary}>
                  {row.secondary || '—'}
                </span>
                <span className="set-tkb-scope" title={t('tkbScopePriority')}>
                  {entry.scope}
                </span>
                <span className="set-tkb-prio">{entry.priority}</span>
                <button
                  type="button"
                  className="set-tkb-del"
                  title={t('tkbDelete')}
                  aria-label={`${t('tkbDelete')} ${row.primary}`}
                  onClick={() => void remove(entry.id)}
                >
                  ✕
                </button>
              </div>
            )
          })
        )}
      </div>

      <div className="set-mp-upload set-tkb-add">
        <h5>{t('tkbAdd')}</h5>
        <div className="set-mp-upload-grid">
          {fields.map((f) => (
            <label className="set-mp-field" key={f.id}>
              <span>{t(f.labelKey)}</span>
              <input
                type="text"
                value={draft[f.id] ?? ''}
                spellCheck={false}
                onChange={(e) => setDraft({ ...draft, [f.id]: e.target.value })}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') void add()
                }}
              />
            </label>
          ))}
          {schema === 'brand' && (
            <label className="set-mp-field">
              <span>{t('tkbFieldPolicy')}</span>
              <select
                value={policy}
                onChange={(e) => setPolicy(e.target.value as (typeof TKB_BRAND_POLICIES)[number])}
              >
                {TKB_BRAND_POLICIES.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
            </label>
          )}
          <label className="set-mp-field">
            <span>{t('tkbScope')}</span>
            <select value={scope} onChange={(e) => setScope(e.target.value as TranslationKbScope)}>
              {TKB_SCOPES.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>
          <label className="set-mp-field">
            <span>{t('tkbPriority')}</span>
            <input
              type="number"
              min={0}
              max={999}
              value={priority}
              onChange={(e) => setPriority(Math.max(0, Math.min(999, Number(e.target.value) || 0)))}
            />
          </label>
        </div>
        <div className="set-mp-upload-actions">
          <button type="button" className="set-btn primary" onClick={() => void add()}>
            {t('tkbAdd')}
          </button>
          {flash && <span className="set-tkb-flash">{t('tkbSaved')}</span>}
          {error && <span className="set-tkb-flash is-err">{error}</span>}
        </div>
      </div>

      <h4 className="set-pane-subtitle">{t('tkbDictTitle')}</h4>
      <div className="set-field-desc set-ai-note">{t('tkbDictDesc')}</div>
      <div className="set-mp-upload">
        <div className="set-mp-upload-grid">
          <label className="set-mp-field set-mp-field-wide">
            <span>{t('tkbDictTitle')}</span>
            <div className="set-tkb-file">
              <input
                type="text"
                value={filePath}
                placeholder={t('tkbNoFile')}
                spellCheck={false}
                onChange={(e) => setFilePath(e.target.value)}
                onDragOver={(e) => {
                  e.preventDefault()
                }}
                onDrop={(e) => {
                  e.preventDefault()
                  const file = e.dataTransfer.files[0]
                  if (!file) return
                  // Electron: the path is on the File object; the browser does
                  // not expose it, so we fall back to a regular file picker.
                  const filePath =
                    (file as File & { path?: string }).path ?? ''
                  if (filePath) {
                    setFilePath(filePath)
                  } else {
                    void pickFile()
                  }
                }}
              />
              <button type="button" className="set-btn" onClick={() => void pickFile()}>
                {t('tkbPickFile')}
              </button>
            </div>
          </label>
          <label className="set-mp-field">
            <span>{t('tkbSourceLang')}</span>
            <select value={sourceLang} onChange={(e) => setSourceLang(e.target.value)}>
              {TKB_LANGS.map((l) => (
                <option key={l.value} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </label>
          <div className="set-mp-swap">
            <button
              type="button"
              className="set-tkb-swap"
              title="⇄"
              onClick={swapLanguages}
              aria-label="⇄"
            >
              ⇄
            </button>
          </div>
          <label className="set-mp-field">
            <span>{t('tkbTargetLang')}</span>
            <select value={targetLang} onChange={(e) => setTargetLang(e.target.value)}>
              {TKB_LANGS.filter((l) => l.value !== 'auto').map((l) => (
                <option key={l.value} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </label>
          <label className="set-mp-field">
            <span>{t('tkbCustomer')}</span>
            <input
              type="text"
              value={customerName}
              spellCheck={false}
              onChange={(e) => setCustomerName(e.target.value)}
            />
          </label>
        </div>
        <label className="set-tkb-check">
          <input type="checkbox" checked={kbOnly} onChange={(e) => setKbOnly(e.target.checked)} />
          {t('tkbKbOnly')}
        </label>
        <div className="set-mp-upload-actions">
          <button
            type="button"
            className="set-btn primary"
            disabled={busy || skillMissing}
            onClick={() => void build(false)}
          >
            {busy ? t('tkbDictBuilding') : t('tkbDictBuild')}
          </button>
          <button
            type="button"
            className="set-btn primary"
            disabled={busy || skillMissing || (!providerConfigured && !kbOnly)}
            onClick={() => void build(true)}
          >
            {t('tkbGenerateAndTranslate')}
          </button>
        </div>

        {dict && (
          <div className="set-mp-upload-history set-tkb-result-card">
            <div className="set-tkb-result">
              <span className="set-cap-pill is-on">{t('tkbDictKb', { count: dict.kbEntries })}</span>
              <span className="set-cap-pill is-fallback">
                {t('tkbDictLlm', { count: dict.llmEntries })}
              </span>
              {dict.missed > 0 && (
                <span className="set-cap-pill is-off">
                  {t('tkbDictMissed', { count: dict.missed })}
                </span>
              )}
              {dict.totalSegments > 0 && (
                <span className="set-cap-pill is-on" title={t('tkbExtracted', { count: dict.totalSegments })}>
                  {t('tkbExtracted', { count: dict.totalSegments })}
                </span>
              )}
              {dict.elapsedMs > 0 && (
                <span className="set-cap-pill is-fallback">
                  {t('tkbElapsed', { ms: dict.elapsedMs })}
                </span>
              )}
            </div>
            {dict.segments.length > 0 && (
              <div className="set-tkb-preview">
                <div className="set-tkb-preview-title">
                  {t('tkbPreview', { count: dict.segments.length })}
                </div>
                <div className="set-tkb-preview-chips">
                  {dict.segments.slice(0, 12).map((s, i) => (
                    <span
                      key={i}
                      className={`set-tkb-chip${s.origin === 'kb' ? ' is-kb' : ''}`}
                      title={s.source}
                    >
                      {s.source}
                      <span className="set-tkb-chip-arrow">→</span>
                      {s.target}
                    </span>
                  ))}
                  {dict.segments.length > 12 && (
                    <span className="set-tkb-chip is-more">+{dict.segments.length - 12}</span>
                  )}
                </div>
              </div>
            )}
            <div className="set-mp-upload-item">
              <code>{dict.dictionaryPath}</code>
              <button
                type="button"
                className="set-btn"
                onClick={() => void window.aiOffice.revealPath?.(dict.dictionaryPath)}
              >
                {t('tkbRevealInFinder')}
              </button>
              {dict.llmEntries > 0 && (
                <button
                  type="button"
                  className="set-btn"
                  disabled={savingDict || savedDictCount > 0}
                  onClick={() => void saveDictToKb()}
                >
                  {savedDictCount > 0
                    ? t('tkbSaved')
                    : t('tkbSaveDict', { count: dict.llmEntries })}
                </button>
              )}
            </div>
          </div>
        )}

        {outputPath && (
          <div className="set-mp-upload-history set-tkb-output">
            <div className="set-tkb-output-head">
              <span className="set-cap-pill is-on">{t('tkbOpenOutput')}</span>
              <code>{outputPath}</code>
            </div>
            <div className="set-mp-upload-actions">
              <button
                type="button"
                className="set-btn primary"
                onClick={() => void window.aiOffice.openPath(outputPath)}
              >
                {t('tkbOpenOutput')}
              </button>
              <button
                type="button"
                className="set-btn"
                onClick={() => void window.aiOffice.revealPath?.(outputPath)}
              >
                {t('tkbRevealInFinder')}
              </button>
            </div>
          </div>
        )}

        {dictError && <div className="set-tkb-flash is-err">{dictError}</div>}
        {!providerConfigured && !kbOnly && !dict && (
          <div className="set-tkb-flash is-err">
            {t('tkbNoProvider')}
          </div>
        )}
      </div>

      <h4 className="set-pane-subtitle">{t('tkbSnippet')}</h4>
      <div className="set-field-desc set-ai-note">{t('tkbSnippetDesc')}</div>
      <div className="set-mp-upload set-tkb-snippet">
        <div className="set-mp-upload-grid">
          <label className="set-mp-field set-mp-field-wide">
            <span>{t('tkbSnippetSource')}</span>
            <textarea
              className="set-tkb-snippet-area"
              rows={3}
              value={snippet}
              placeholder={t('tkbSnippetPlaceholder')}
              onChange={(e) => setSnippet(e.target.value)}
              onKeyDown={(e) => {
                if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') void translateSnippet()
              }}
            />
          </label>
        </div>
        <div className="set-mp-upload-actions">
          <button
            type="button"
            className="set-btn primary"
            disabled={snippetBusy || !snippet.trim() || skillMissing}
            onClick={() => void translateSnippet()}
          >
            {snippetBusy ? t('tkbSnippetTranslating') : t('tkbSnippetTranslate')}
          </button>
          {sourceLang !== 'auto' && sourceLang === targetLang && (
            <span className="set-cap-pill is-off">
              {t('tkbSourceLang')} = {t('tkbTargetLang')}
            </span>
          )}
          {snippetResult?.status === 'memory-hit' && (
            <span className="set-cap-pill is-on" title={t('tkbSnippetMemoryHit')}>
              {t('tkbSnippetMemoryHit')}
            </span>
          )}
          {snippetResult && snippetResult.elapsedMs > 0 && (
            <span className="set-cap-pill is-fallback">
              {t('tkbSnippetElapsed', { ms: snippetResult.elapsedMs })}
            </span>
          )}
          {snippetResult && snippetResult.matchedTerms.length > 0 && (
            <span
              className="set-cap-pill is-on"
              title={t('tkbSnippetTerms', { terms: snippetResult.matchedTerms.join(', ') })}
            >
              KB · {snippetResult.matchedTerms.length}
            </span>
          )}
        </div>
        {snippetResult && (
          <div className="set-tkb-snippet-out">
            <span className="set-tkb-snippet-out-label">{t('tkbSnippetResult')}</span>
            <pre className="set-tkb-snippet-pre">{snippetResult.translation}</pre>
            <button
              type="button"
              className="set-btn"
              onClick={() => void copySnippet()}
            >
              {snippetCopied ? t('tkbSaved') : t('tkbSnippetCopy')}
            </button>
          </div>
        )}
        {snippetError && <div className="set-tkb-flash is-err">{snippetError}</div>}
      </div>
    </>
  )
}
function AiStatusPill({ status }: { status: AiStatus | null }) {
  if (!status) return null
  return (
    <span
      className={`set-ai-status ${status.kind}`}
      role="status"
      // error text (HTTP body, network message) can be long — full text via native tooltip
      title={status.kind === 'err' ? status.text : undefined}
    >
      {status.kind === 'testing' ? (
        <span className="set-ai-spin" aria-hidden="true" />
      ) : status.kind === 'ok' ? (
        <svg
          className="set-ai-status-icon"
          width="14"
          height="14"
          viewBox="0 0 14 14"
          aria-hidden="true"
        >
          <circle cx="7" cy="7" r="6.3" fill="currentColor" opacity="0.16" />
          <path
            d="M4.2 7.3l1.9 1.9 3.7-4.3"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
            fill="none"
          />
        </svg>
      ) : (
        <svg
          className="set-ai-status-icon"
          width="14"
          height="14"
          viewBox="0 0 14 14"
          aria-hidden="true"
        >
          <circle cx="7" cy="7" r="6.3" fill="currentColor" opacity="0.16" />
          <path d="M7 3.8v3.9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
          <circle cx="7" cy="10.1" r="1" fill="currentColor" />
        </svg>
      )}
      <span className="set-ai-status-text">{status.text}</span>
    </span>
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
  /** an installed skill is older than the bundled one: dot on the Integrations entry */
  skillUpdateDue?: boolean
  onSkillUpdateDue?: (due: boolean) => void
}

/** Skills & Plugins management pane — control GenOffice's agent extensions
 *  backed by @genoffice/agent-skills (11 built-in extensions: 8 skills + 3 plugins).
 *  Each entry can be enabled/disabled, hot-reloaded, or installed from marketplace.
 */
type MpCategory = MarketplaceCategory
type MpEntry = MarketplaceEntry

/** 'a, b , c' → ['a','b','c'] — shared by the publish form */
const splitList = (v: string): string[] =>
  v.split(',').map((s) => s.trim()).filter(Boolean)

function SkillsPluginsPane({ t }: { t: TFunc }) {
  const [skills, setSkills] = useState<SkillEntry[]>([])
  const [plugins, setPlugins] = useState<PluginEntry[]>([])
  const [loading, setLoading] = useState(true)
  // Marketplace v2 — search / filter / sort
  const [mpSkills, setMpSkills] = useState<MpEntry[]>([])
  const [mpPlugins, setMpPlugins] = useState<MpEntry[]>([])
  const [mpCategories, setMpCategories] = useState<MarketplaceCategoryInfo[]>([])
  const [mpTotal, setMpTotal] = useState(0)
  const [mpQ, setMpQ] = useState('')
  const [mpCategory, setMpCategory] = useState<MpCategory | ''>('')
  const [mpType, setMpType] = useState<'all' | 'skill' | 'plugin'>('all')
  const [mpSort, setMpSort] = useState<'popular' | 'rating' | 'newest' | 'name'>('popular')
  const [mpInstalled, setMpInstalled] = useState<'all' | boolean>('all')
  const [minRating, setMinRating] = useState(0)
  const [marketMsg, setMarketMsg] = useState<string | null>(null)
  const [detailEntry, setDetailEntry] = useState<{
    entry: MpEntry
    kind: 'skill' | 'plugin'
    pi?: { skillPath?: string | null; installed?: boolean; enabled?: boolean; packageDir?: string; hasCode?: boolean; artifact?: { filename: string; size?: string } | null; piPackage?: string; mode?: string }
    installed?: unknown
  } | null>(null)
  const [showUpload, setShowUpload] = useState(false)
  const [uploadKind, setUploadKind] = useState<'skill' | 'plugin'>('skill')
  const [uploadUploads, setUploadUploads] = useState<MarketplaceUploadEntry[]>([])
  const [uploadArtifact, setUploadArtifact] = useState<{ filename: string; content: string; size: number } | null>(null)
  const [uploadStatus, setUploadStatus] = useState<string | null>(null)
  const [uploadForm, setUploadForm] = useState({
    id: '',
    name: '',
    description: '',
    longDescription: '',
    version: '1.0.0',
    tools: '',
    scopes: '',
    category: 'productivity' as MpCategory,
    tags: '',
    author: '',
    icon: '',
    homepage: '',
  })
  const [uploadMsg, setUploadMsg] = useState<{ kind: 'ok' | 'err'; text: string } | null>(null)
  const [uploading, setUploading] = useState(false)
  const [rateDraft, setRateDraft] = useState<number>(5)
  const [rateMsg, setRateMsg] = useState<{ kind: 'ok' | 'err'; text: string } | null>(null)
  const [rateBusy, setRateBusy] = useState(false)
  // pi runtime + search UX
  const [piResources, setPiResources] = useState<PiResourceReport | null>(null)
  const [piLoading, setPiLoading] = useState(false)
  const [recentSearches, setRecentSearches] = useState<string[]>([])
  const [activeIndex, setActiveIndex] = useState(-1)
  const [piSectionOpen, setPiSectionOpen] = useState(false)
  const [unpublishingId, setUnpublishingId] = useState<string | null>(null)
  const searchInputRef = useRef<HTMLInputElement | null>(null)
  const mpSectionRef = useRef<HTMLDivElement | null>(null)
  const artifactInputRef = useRef<HTMLInputElement | null>(null)

  const readableBytes = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
  }

  const handleArtifactFile = async (file: File) => {
    // 128KB cap so we don't bloat the upload body.
    if (file.size > 128 * 1024) {
      setUploadMsg({ kind: 'err', text: 'Artifact too large (max 128 KB).' })
      return
    }
    const text = await file.text()
    setUploadArtifact({ filename: file.name, content: text, size: file.size })
    setUploadMsg(null)
  }

  /** Build the flattened result index used by keyboard navigation. */
  const mpResults = useMemo(() => {
    const list: Array<{ id: string; kind: 'skill' | 'plugin'; entry: MpEntry }> = []
    for (const m of mpSkills) list.push({ id: m.id, kind: 'skill', entry: m })
    for (const m of mpPlugins) list.push({ id: m.id, kind: 'plugin', entry: m })
    return list
  }, [mpSkills, mpPlugins])

  const handleSearchKey = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Escape') {
      if (mpQ) {
        setMpQ('')
      } else if (detailEntry) {
        setDetailEntry(null)
      } else {
        searchInputRef.current?.blur()
      }
      e.preventDefault()
      return
    }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      const max = mpResults.length - 1
      if (max < 0) return
      const dir = e.key === 'ArrowDown' ? 1 : -1
      const next = activeIndex < 0 ? (dir > 0 ? 0 : max) : activeIndex + dir
      const wrapped = ((next % (max + 1)) + (max + 1)) % (max + 1)
      setActiveIndex(wrapped)
      e.preventDefault()
      return
    }
    if (e.key === 'Enter') {
      if (activeIndex >= 0 && mpResults[activeIndex]) {
        const target = mpResults[activeIndex]
        void openDetail(target.entry, target.kind)
        e.preventDefault()
      } else if (mpQ.trim()) {
        rememberSearch(mpQ)
      }
    }
  }

  /** Global keyboard shortcuts: '/' or Cmd/Ctrl+K focus the search input. */
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null
      const inEditable = !!target && (
        target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable
      )
      if (!inEditable && e.key === '/') {
        e.preventDefault()
        searchInputRef.current?.focus()
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        searchInputRef.current?.focus()
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [])

  const unpublishUpload = async (kind: 'skill' | 'plugin', id: string) => {
    setUnpublishingId(id)
    try {
      const res = await window.aiOffice.marketplaceDeleteUpload?.(kind, id)
      const listed = await window.aiOffice.marketplaceListUploads?.()
      setUploadUploads((listed as { uploads?: MarketplaceUploadEntry[] })?.uploads ?? [])
      const note = res && typeof res === 'object' && 'uninstalled' in res
        ? (res as { uninstalled?: boolean }).uninstalled
        : undefined
      setUploadMsg({
        kind: 'ok',
        text: note
          ? t('mpUploadUnpublishedWithUninstall')
          : t('mpUploadUnpublished'),
      })
      await refresh()
      await refreshPi()
    } catch (err) {
      setUploadMsg({ kind: 'err', text: err instanceof Error ? err.message : String(err) })
    } finally {
      setUnpublishingId(null)
    }
  }

  // recent searches persisted under localStorage
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem('genoffice.mp.recent')
      if (raw) {
        const arr = JSON.parse(raw)
        if (Array.isArray(arr)) {
          setRecentSearches(arr.filter((x): x is string => typeof x === 'string').slice(0, 8))
        }
      }
    } catch {
      // localStorage may be unavailable in some test environments — ignore.
    }
  }, [])

  const rememberSearch = useCallback((q: string) => {
    const trimmed = q.trim()
    if (!trimmed) return
    setRecentSearches((prev) => {
      const next = [trimmed, ...prev.filter((x) => x.toLowerCase() !== trimmed.toLowerCase())].slice(0, 5)
      try { window.localStorage.setItem('genoffice.mp.recent', JSON.stringify(next)) } catch {}
      return next
    })
  }, [])

  const refreshPi = useCallback(async () => {
    setPiLoading(true)
    try {
      const r = await window.aiOffice.listPiResources?.()
      if (r) setPiResources(r as PiResourceReport)
    } finally {
      setPiLoading(false)
    }
  }, [])

  const refresh = useCallback(async () => {
    const [s, p, mp, cats] = await Promise.all([
      window.aiOffice.listSkills?.() ?? Promise.resolve([]),
      window.aiOffice.listPlugins?.() ?? Promise.resolve([]),
      window.aiOffice.marketplaceSearch?.({
        q: mpQ,
        category: mpCategory || undefined,
        type: mpType === 'all' ? undefined : mpType,
        minRating: minRating || undefined,
        installed: mpInstalled,
        sort: mpSort,
      }) ?? Promise.resolve({ skills: [], plugins: [], total: 0 }),
      window.aiOffice.marketplaceCategories?.() ?? Promise.resolve({ categories: [] }),
    ])
    setSkills(s)
    setPlugins(p)
    setMpSkills((mp as { skills?: MpEntry[] }).skills ?? [])
    setMpPlugins((mp as { plugins?: MpEntry[] }).plugins ?? [])
    setMpTotal((mp as { total?: number }).total ?? 0)
    setMpCategories((cats as { categories?: MarketplaceCategoryInfo[] }).categories ?? [])
    setLoading(false)
  }, [mpQ, mpCategory, mpType, mpSort, mpInstalled, minRating])

  useEffect(() => {
    const t = setTimeout(() => { void refresh() }, 120)
    return () => clearTimeout(t)
  }, [refresh])

  useEffect(() => {
    if (!showUpload) return
    void window.aiOffice.marketplaceListUploads?.().then((r) => {
      setUploadUploads((r as { uploads?: MarketplaceUploadEntry[] }).uploads ?? [])
    })
  }, [showUpload])

  useEffect(() => {
    // Fetch pi runtime report on first mount so the panel can render even
    // before the user opens the marketplace section.
    void refreshPi()
  }, [refreshPi])

  useEffect(() => {
    void refresh()
  }, [refresh])

  if (loading) return null

  const toggleSkill = async (id: SkillKind, enabled: boolean) => {
    const next = await window.aiOffice.toggleSkill?.(id, enabled)
    if (Array.isArray(next)) setSkills(next)
  }

  const reloadSkill = async (id: SkillKind) => {
    const next = await window.aiOffice.reloadSkill?.(id)
    if (Array.isArray(next)) setSkills(next)
  }

  const togglePlugin = async (id: PluginKind, enabled: boolean) => {
    const next = await window.aiOffice.togglePlugin?.(id, enabled)
    if (Array.isArray(next)) setPlugins(next)
  }

  const reloadPlugin = async (id: PluginKind) => {
    const next = await window.aiOffice.reloadPlugin?.(id)
    if (Array.isArray(next)) setPlugins(next)
  }

  const resetSkills = async () => {
    const next = await window.aiOffice.resetSkills?.()
    if (Array.isArray(next)) setSkills(next)
  }

  const resetPlugins = async () => {
    const next = await window.aiOffice.resetPlugins?.()
    if (Array.isArray(next)) setPlugins(next)
  }

  const installMarketSkill = async (id: string) => {
    setMarketMsg(null)
    const res = await window.aiOffice.installSkill?.(id)
    if (res?.ok) {
      setMarketMsg(`✓ Installed skill "${id}" — see Skills list above and reload to activate`)
      await refresh()
    } else {
      setMarketMsg(`✗ ${res?.error || 'Install failed'}`)
    }
  }

  const installMarketPlugin = async (id: string) => {
    setMarketMsg(null)
    const res = await window.aiOffice.installPlugin?.(id)
    if (res?.ok) {
      setMarketMsg(`✓ Installed plugin "${id}" — see Plugins list above and reload to activate`)
      await refresh()
    } else {
      setMarketMsg(`✗ ${res?.error || 'Install failed'}`)
    }
  }

  const uninstallMarketPlugin = async (id: string) => {
    setMarketMsg(null)
    const res = await window.aiOffice.uninstallPlugin?.(id as PluginKind)
    if (res?.ok) {
      setMarketMsg(`✓ Uninstalled plugin "${id}"`)
      await refresh()
    } else {
      setMarketMsg(`✗ ${res?.error || 'Uninstall failed'}`)
    }
  }

  const uninstallMarketSkill = async (id: SkillKind) => {
    setMarketMsg(null)
    const res = await window.aiOffice.uninstallSkill?.(id)
    if (res?.ok) {
      setMarketMsg(`✓ Uninstalled skill "${id}"`)
      await refresh()
    } else {
      setMarketMsg(`✗ ${res?.error || 'Uninstall failed'}`)
    }
  }

  const openDetail = async (entry: MpEntry, kind: 'skill' | 'plugin') => {
    // Open optimistically with the entry from the grid, then hydrate from
    // marketplace-detail (the canonical source). We track `kind` alongside
    // the entry so submitRating / uninstall calls don't have to re-derive
    // it from a fragile heuristic (e.g. checking tool prefixes).
    setDetailEntry({ entry, kind })
    const res = (await window.aiOffice.marketplaceDetail?.(entry.id, kind)) as
      | { ok: boolean; type?: 'skill' | 'plugin'; entry?: MpEntry; pi?: Record<string, unknown>; installed?: unknown; error?: string }
      | undefined
    if (res?.ok && res.entry) {
      setDetailEntry({
        entry: res.entry,
        kind: res.type || kind,
        pi: res.pi,
        installed: res.installed,
      })
    }
  }

  const submitUpload = async () => {
    setUploadMsg(null)
    setUploading(true)
    try {
      const payload: MarketplaceUploadPayload = {
        id: uploadForm.id.trim(),
        name: uploadForm.name.trim(),
        description: uploadForm.description.trim(),
        longDescription: uploadForm.longDescription.trim() || undefined,
        version: uploadForm.version.trim(),
        tools: splitList(uploadForm.tools),
        scopes: splitList(uploadForm.scopes),
        category: uploadForm.category,
        tags: splitList(uploadForm.tags),
        author: uploadForm.author.trim() || undefined,
        icon: uploadForm.icon.trim() || undefined,
        homepage: uploadForm.homepage.trim() || undefined,
        artifact: uploadArtifact
          ? { filename: uploadArtifact.filename, content: uploadArtifact.content }
          : undefined,
      }
      // Detect a likely overwrite by probing the marketplace before submit.
      // If the catalog already contains an entry with this id, confirm with
      // the user before sending `force: true`.
      const probe = await window.aiOffice.marketplaceSearch?.({
        q: payload.id,
        type: uploadKind === 'plugin' ? 'plugin' : 'skill',
      })
      const probeEntries = (probe as { plugins?: MpEntry[]; skills?: MpEntry[] } | undefined)
      const matches = uploadKind === 'plugin'
        ? probeEntries?.plugins ?? []
        : probeEntries?.skills ?? []
      const conflict = matches.some((m) => m.id === payload.id)
      if (conflict) {
        const ok = window.confirm(t('mpOverwriteConfirm'))
        if (!ok) {
          setUploading(false)
          return
        }
        ;(payload as { force?: boolean }).force = true
      }
      const res = await window.aiOffice.marketplaceUpload?.(uploadKind, payload)
      if (res?.ok) {
        setUploadMsg({ kind: 'ok', text: res.message ?? '✓' })
        setUploadForm({
          id: '',
          name: '',
          description: '',
          longDescription: '',
          version: '1.0.0',
          tools: '',
          scopes: '',
          category: 'productivity',
          tags: '',
          author: '',
          icon: '',
          homepage: '',
        })
        setUploadArtifact(null)
        const listed = await window.aiOffice.marketplaceListUploads?.()
        setUploadUploads((listed as { uploads?: MarketplaceUploadEntry[] })?.uploads ?? [])
        // a publish changes the catalog: re-run the search so the new entry
        // shows up in the grid without the user having to touch a filter
        await refresh()
      } else {
        setUploadMsg({ kind: 'err', text: res?.error ?? 'Upload failed' })
      }
    } catch (err) {
      setUploadMsg({ kind: 'err', text: err instanceof Error ? err.message : String(err) })
    } finally {
      setUploading(false)
    }
  }

  const submitRating = async () => {
    if (!detailEntry) return
    setRateMsg(null)
    setRateBusy(true)
    try {
      // detailEntry carries the authoritative `kind` captured when the
      // drawer opened — no more brittle heuristic over tool prefixes or
      // author names. Server's marketplaceRate only tracks user uploads, so
      // a curated skill will report ok=false with a clear error and the UI
      // surfaces it.
      const finalKind = detailEntry.kind
      const res = await window.aiOffice.marketplaceRate?.(detailEntry.entry.id, finalKind, rateDraft)
      if (res?.ok) {
        setRateMsg({
          kind: 'ok',
          text: `${t('mpRateThanks')} 平均 ${res.averageRating?.toFixed(2)} (${res.ratingCount} 次评分)`,
        })
        await refresh()
        const refreshed = await window.aiOffice.marketplaceDetail?.(detailEntry.entry.id, finalKind)
        if (refreshed?.ok && refreshed.entry) setDetailEntry({ entry: refreshed.entry, kind: finalKind })
      } else {
        setRateMsg({ kind: 'err', text: res?.error ?? 'Rating failed' })
      }
    } catch (err) {
      setRateMsg({ kind: 'err', text: err instanceof Error ? err.message : String(err) })
    } finally {
      setRateBusy(false)
    }
  }

  /** Highlight matches of `q` tokens inside a string. Returns ReactNodes. */
  const highlightTokens = (text: string, q: string): React.ReactNode => {
    if (!q.trim()) return text
    const tokens = q.trim().split(/\s+/).filter(Boolean)
    if (tokens.length === 0) return text
    const escaped = tokens.map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    const re = new RegExp(`(${escaped.join('|')})`, 'gi')
    const parts = text.split(re)
    return parts.map((part, i) =>
      i % 2 === 1 ? <mark key={i} className="set-mp-mark">{part}</mark> : <span key={i}>{part}</span>,
    )
  }

  /** marketplace grid card — icon, rating, tags, install/uninstall + detail */
  const renderMpCard = (m: MpEntry, kind: 'skill' | 'plugin') => {
    const catKey = ('mpCat' + m.category[0].toUpperCase() + m.category.slice(1)) as StringKey
    const cardIndex = mpResults.findIndex((r) => r.id === m.id && r.kind === kind)
    const isActive = activeIndex >= 0 && cardIndex === activeIndex
    const isExtension = m.artifact?.kind === 'extension'
    return (
      <article
        key={`${kind}:${m.id}`}
        className={`set-mp-card${isActive ? ' is-active' : ''}`}
        data-mp-id={m.id}
        data-mp-kind={kind}
        data-installed={m.installed ? '1' : '0'}
        data-active={isActive ? '1' : '0'}
        data-mp-index={cardIndex}
      >
        <header className="set-mp-card-head">
          <span className="set-mp-card-icon">{m.icon || m.name[0]}</span>
          <div className="set-mp-card-title">
            <span className="set-mp-card-name">{highlightTokens(m.name, mpQ)}</span>
            <span className="set-mp-card-sub">v{m.version} · {m.author}</span>
          </div>
          {m.featured && <span className="set-mp-featured">{t('mpFeatured')}</span>}
        </header>

        <div className="set-mp-card-badges">
          {m.piPackage && <span className="set-mp-badge">{t('mpCardPi')}</span>}
          {isExtension
            ? <span className="set-mp-badge set-mp-badge-code">{t('mpCardHasCode')}</span>
            : <span className="set-mp-badge set-mp-badge-soft">{t('mpCardGuidance')}</span>}
        </div>

        <p className="set-mp-card-desc">{highlightTokens(m.description, mpQ)}</p>

        <div className="set-mp-card-tags">
          <span className="set-mp-tag set-mp-tag-cat">{t(catKey)}</span>
          {m.tags.slice(0, 3).map((tag) => (
            <span key={tag} className="set-mp-tag">{tag}</span>
          ))}
        </div>

        <footer className="set-mp-card-foot">
          <span className="set-mp-card-rating" title={t('mpRating')}>★ {m.rating.toFixed(1)}</span>
          <span className="set-mp-card-downloads">↓ {m.downloads.toLocaleString()}</span>
          <span className="set-mp-spacer" />
          <button
            type="button"
            className="set-btn-mini"
            onClick={() => void openDetail(m, kind)}
            title={t('mpDetails')}
          >
            {t('mpDetails')}
          </button>
          {m.installed ? (
            <button
              type="button"
              className="set-btn-mini set-btn-uninstall"
              onClick={() =>
                void (kind === 'skill'
                  ? uninstallMarketSkill(m.id as SkillKind)
                  : uninstallMarketPlugin(m.id))
              }
              title={t('uninstallBtn')}
            >
              − {t('uninstallBtn')}
            </button>
          ) : (
            <button
              type="button"
              className="set-btn-mini set-btn-install"
              onClick={() =>
                void (kind === 'skill' ? installMarketSkill(m.id) : installMarketPlugin(m.id))
              }
              title={t('install')}
            >
              + {t('install')}
            </button>
          )}
        </footer>
      </article>
    )
  }

  const renderSkillRow = (s: SkillEntry) => (
    <li key={s.id} className="set-skill-row" data-skill-id={s.id}>
      <div className="set-skill-head">
        <label className="set-toggle">
          <input
            type="checkbox"
            checked={s.status === 'enabled'}
            onChange={(e) => void toggleSkill(s.id, e.target.checked)}
          />
          <span className="set-skill-name">{s.name}</span>
          <span className={`set-skill-badge set-skill-badge-${s.status}`}>{s.status}</span>
        </label>
        <button
          type="button"
          className="set-btn-mini"
          onClick={() => void reloadSkill(s.id)}
          title={t('reloadSkill')}
        >
          ↻
        </button>
      </div>
      <div className="set-skill-desc">{s.description}</div>
      <div className="set-skill-meta">
        <span>v{s.version}</span>
        <span>· {s.author}</span>
        <span>· {s.tools.length} {t('tools')}</span>
        <span>· {s.package}/{s.source.split('/').pop()}</span>
      </div>
      <div className="set-skill-scopes">
        {s.scopes.map((sc) => (
          <span key={sc} className="set-scope-tag">{sc}</span>
        ))}
      </div>
    </li>
  )

  const renderPluginRow = (p: PluginEntry) => (
    <li key={p.id} className="set-skill-row" data-plugin-id={p.id}>
      <div className="set-skill-head">
        <label className="set-toggle">
          <input
            type="checkbox"
            checked={p.status === 'enabled'}
            onChange={(e) => void togglePlugin(p.id, e.target.checked)}
          />
          <span className="set-skill-name">{p.name}</span>
          <span className={`set-skill-badge set-skill-badge-${p.status}`}>{p.status}</span>
        </label>
        <button
          type="button"
          className="set-btn-mini"
          onClick={() => void reloadPlugin(p.id)}
          title={t('reloadPlugin')}
        >
          ↻
        </button>
      </div>
      <div className="set-skill-desc">{p.description}</div>
      <div className="set-skill-meta">
        <span>v{p.version}</span>
        <span>· {p.author}</span>
        <span>· {p.tools.length} {t('tools')}</span>
        <span>· {p.package}/{p.source.split('/').pop()}</span>
      </div>
      {p.requirements.length > 0 && (
        <div className="set-skill-scopes">
          <span className="set-req-tag">{t('requirements')}: </span>
          {p.requirements.map((r) => (
            <span key={r} className="set-scope-tag set-scope-tag-req">{r}</span>
          ))}
        </div>
      )}
    </li>
  )

  return (
    <>
      <h3 className="set-pane-title">{t('setSecSkillsPlugins')}</h3>
      <div className="set-field-stack" style={{ marginBottom: 14 }}>
        <div className="set-field-label">{t('skillsPluginsIntro')}</div>
        <div className="set-field-desc">
          {t('skillsPluginsDesc', { count: skills.length + plugins.length })}
        </div>
      </div>

      <div className="set-skill-section">
        <div className="set-skill-section-head">
          <h4>{t('skillsTitle')}</h4>
          <span className="set-skill-count">{skills.length}</span>
          <button type="button" className="set-btn-mini" onClick={() => void resetSkills()}>
            {t('reset')}
          </button>
        </div>
        <div className="set-field-desc" style={{ marginBottom: 8 }}>{t('skillsDesc')}</div>
        <ul className="set-skill-list" role="list">{skills.map(renderSkillRow)}</ul>
      </div>

      <div className="set-skill-section">
        <div className="set-skill-section-head">
          <h4>{t('pluginsTitle')}</h4>
          <span className="set-skill-count">{plugins.length}</span>
          <button type="button" className="set-btn-mini" onClick={() => void resetPlugins()}>
            {t('reset')}
          </button>
        </div>
        <div className="set-field-desc" style={{ marginBottom: 8 }}>{t('pluginsDesc')}</div>
        <ul className="set-skill-list" role="list">{plugins.map(renderPluginRow)}</ul>
      </div>

      <div className="set-skill-section set-mp" data-marketplace="v2" ref={mpSectionRef}>
        <div className="set-skill-section-head">
          <h4>{t('marketplace')}</h4>
          <span className="set-skill-count">{mpTotal}</span>
          <button
            type="button"
            className={`set-btn-mini set-mp-pi-toggle ${piSectionOpen ? 'is-open' : ''}`}
            aria-expanded={piSectionOpen}
            onClick={() => setPiSectionOpen((v) => !v)}
            title={t('mpPiTitle')}
          >
            <span aria-hidden="true">π</span>
            <span className="set-mp-pi-counts">
              {piResources
                ? `${piResources.extensions.length}/${piResources.skills.length}`
                : '…'}
            </span>
          </button>
          <span className="set-mp-spacer" />
          <button
            type="button"
            className="set-btn set-btn-primary set-mp-publish"
            onClick={() => { setShowUpload((v) => !v); setUploadMsg(null) }}
          >
            {showUpload ? t('mpClose') : t('mpUpload')}
          </button>
        </div>

        {piSectionOpen && (
          <div className="set-mp-pi" data-pi-runtime="1">
            <div className="set-mp-pi-head">
              <span className="set-mp-pi-title">{t('mpPiTitle')}</span>
              <button
                type="button"
                className="set-btn-mini"
                disabled={piLoading}
                onClick={() => void refreshPi()}
                title={t('mpPiRefresh')}
              >
                {piLoading ? '…' : t('mpPiRefresh')}
              </button>
            </div>
            {piResources ? (
              <>
                <div className="set-mp-pi-paths">
                  <span className="set-mp-pi-path">
                    {t('mpPiAgentDir')}: <code>{piResources.agentDir || '—'}</code>
                  </span>
                  <span className="set-mp-pi-path">
                    {t('mpPiSettings')}: <code>{piResources.settingsPath || '—'}</code>
                  </span>
                </div>
                <div className="set-mp-pi-counts-row">
                  <span className="set-mp-pi-pill">
                    {t('mpPiExtensions')} <strong>{piResources.extensions.length}</strong>
                    {piResources.external.extensions > 0 && (
                      <em> · {piResources.external.extensions} {t('mpPiExternal')}</em>
                    )}
                  </span>
                  <span className="set-mp-pi-pill">
                    {t('mpPiSkills')} <strong>{piResources.skills.length}</strong>
                    {piResources.external.skills > 0 && (
                      <em> · {piResources.external.skills} {t('mpPiExternal')}</em>
                    )}
                  </span>
                  <span className="set-mp-pi-pill">
                    {t('mpPiPackages')} <strong>{piResources.packages.length}</strong>
                  </span>
                </div>
                {piResources.extensions.length === 0 ? (
                  <div className="set-mp-pi-empty">{t('mpPiEmpty')}</div>
                ) : (
                  <ul className="set-mp-pi-list" role="list">
                    {piResources.extensions.slice(0, 8).map((ext) => (
                      <li key={ext.path} className="set-mp-pi-resource-row" data-managed={ext.managed ? '1' : '0'}>
                        <code>{ext.name}</code>
                        {ext.managed && <span className="set-mp-pi-tag">genoffice</span>}
                        {!ext.enabled && <span className="set-mp-pi-tag set-mp-pi-tag-off">off</span>}
                      </li>
                    ))}
                    {piResources.extensions.length > 8 && (
                      <li className="set-mp-pi-resource-row set-mp-pi-more">
                        +{piResources.extensions.length - 8} more
                      </li>
                    )}
                  </ul>
                )}
                {piResources.skills.length > 0 && (
                  <ul className="set-mp-pi-list" role="list">
                    {piResources.skills.slice(0, 8).map((sk) => (
                      <li key={sk.path} className="set-mp-pi-resource-row" data-managed={sk.managed ? '1' : '0'}>
                        <code>{sk.name}</code>
                        {sk.managed && <span className="set-mp-pi-tag">genoffice</span>}
                        {!sk.enabled && <span className="set-mp-pi-tag set-mp-pi-tag-off">off</span>}
                      </li>
                    ))}
                    {piResources.skills.length > 8 && (
                      <li className="set-mp-pi-resource-row set-mp-pi-more">
                        +{piResources.skills.length - 8} more
                      </li>
                    )}
                  </ul>
                )}
                {piResources.diagnostics.length > 0 && (
                  <ul className="set-mp-pi-diag" role="list">
                    {piResources.diagnostics.map((d, i) => (
                      <li key={i} data-diag-type={d.type}>
                        {d.message}
                      </li>
                    ))}
                  </ul>
                )}
              </>
            ) : (
              <div className="set-mp-pi-empty">{piLoading ? '…' : t('mpPiEmpty')}</div>
            )}
          </div>
        )}

        {/* ── Search bar ─────────────────────────────────────────── */}
        <div className="set-mp-searchrow">
          <span className="set-mp-searchicon" aria-hidden="true">
            <svg width="14" height="14" viewBox="0 0 14 14">
              <circle cx="6" cy="6" r="4.4" stroke="currentColor" strokeWidth="1.5" fill="none" />
              <path d="M9.4 9.4l3 3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
          </span>
          <input
            ref={searchInputRef}
            type="text"
            className="set-input set-mp-search"
            placeholder={t('mpSearchPlaceholder')}
            value={mpQ}
            onChange={(e) => setMpQ(e.target.value)}
            onKeyDown={(e) => handleSearchKey(e)}
            onFocus={() => setActiveIndex(-1)}
            aria-label={t('mpSearchPlaceholder')}
            data-mp-search="1"
          />
          {(mpQ || mpCategory || (mpInstalled !== 'all') || minRating > 0) && (
            <button
              type="button"
              className="set-btn-mini set-mp-reset"
              onClick={() => {
                setMpQ('')
                setMpCategory('')
                setMpInstalled('all')
                setMinRating(0)
                setActiveIndex(-1)
              }}
            >
              {t('mpResetFilters')}
            </button>
          )}
        </div>

        {/* ── Recent searches (only when the box is empty) ───────── */}
        {!mpQ && recentSearches.length > 0 && (
          <div className="set-mp-recent" data-mp-recent="1">
            <span className="set-mp-recent-label">{t('mpRecentSearches')}</span>
            {recentSearches.map((s) => (
              <button
                key={s}
                type="button"
                className="set-mp-recent-chip"
                onClick={() => {
                  setMpQ(s)
                  setActiveIndex(-1)
                  searchInputRef.current?.focus()
                }}
              >
                {s}
              </button>
            ))}
            <button
              type="button"
              className="set-mp-recent-clear"
              onClick={() => {
                setRecentSearches([])
                try { window.localStorage.removeItem('genoffice.mp.recent') } catch {}
              }}
              aria-label="Clear recent searches"
            >
              ×
            </button>
          </div>
        )}

        {/* ── Category chips ─────────────────────────────────────── */}
        <div className="set-mp-chips" role="tablist" aria-label={t('mpCategory')}>
          <button
            type="button"
            className={`set-mp-chip ${mpCategory === '' ? 'set-mp-chip-active' : ''}`}
            onClick={() => setMpCategory('')}
          >
            {t('mpAll')}
          </button>
          {mpCategories.map((c) => {
            const key = ('mpCat' + c.id[0].toUpperCase() + c.id.slice(1)) as StringKey
            return (
              <button
                key={c.id}
                type="button"
                className={`set-mp-chip ${mpCategory === c.id ? 'set-mp-chip-active' : ''}`}
                onClick={() => setMpCategory(mpCategory === c.id ? '' : c.id)}
              >
                {t(key)}
                <span className="set-mp-chip-count">{c.count}</span>
              </button>
            )
          })}
        </div>

        {/* ── Filter row: type / installed / rating / sort ────────── */}
        <div className="set-mp-filters">
          <div className="set-mp-seg" role="group">
            {([['all', t('mpAll')], ['skill', t('mpTypeSkill')], ['plugin', t('mpTypePlugin')]] as const).map(
              ([v, label]) => (
                <button
                  key={v}
                  type="button"
                  className={`set-mp-segbtn ${mpType === v ? 'set-mp-segbtn-active' : ''}`}
                  onClick={() => setMpType(v)}
                >
                  {label}
                </button>
              ),
            )}
          </div>

          <div className="set-mp-seg" role="group">
            {(
              [
                ['all', t('mpAll')],
                [true, t('mpOnlyInstalled')],
                [false, t('mpOnlyAvailable')],
              ] as const
            ).map(([v, label]) => (
              <button
                key={String(v)}
                type="button"
                className={`set-mp-segbtn ${mpInstalled === v ? 'set-mp-segbtn-active' : ''}`}
                onClick={() => setMpInstalled(v)}
              >
                {label}
              </button>
            ))}
          </div>

          <div className="set-mp-seg set-mp-stars" role="group" aria-label={t('mpRating')}>
            {[0, 3, 4, 4.5].map((r) => (
              <button
                key={r}
                type="button"
                className={`set-mp-segbtn ${minRating === r ? 'set-mp-segbtn-active' : ''}`}
                onClick={() => setMinRating(r)}
              >
                {r === 0 ? t('mpAll') : `★ ${r}+`}
              </button>
            ))}
          </div>

          <select
            className="set-mp-sort"
            value={mpSort}
            onChange={(e) => setMpSort(e.target.value as typeof mpSort)}
            aria-label={t('mpSortPopular')}
          >
            <option value="popular">{t('mpSortPopular')}</option>
            <option value="rating">{t('mpSortRating')}</option>
            <option value="newest">{t('mpSortNewest')}</option>
            <option value="name">{t('mpSortName')}</option>
          </select>
        </div>

        {/* ── publish form ───────────────────────────────────────── */}
        {showUpload && (
          <div className="set-mp-upload" data-upload-form="1">
            <h5>{t('mpUploadTitle')}</h5>
            <div className="set-mp-upload-grid">
              <label className="set-mp-field">
                <span>{t('mpUploadKind')}</span>
                <select
                  value={uploadKind}
                  onChange={(e) => setUploadKind(e.target.value as 'skill' | 'plugin')}
                >
                  <option value="skill">{t('mpTypeSkill')}</option>
                  <option value="plugin">{t('mpTypePlugin')}</option>
                </select>
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadId')}</span>
                <input
                  value={uploadForm.id}
                  placeholder="my-extension"
                  onChange={(e) => setUploadForm({ ...uploadForm, id: e.target.value })}
                />
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadName')}</span>
                <input
                  value={uploadForm.name}
                  placeholder="My Extension"
                  onChange={(e) => setUploadForm({ ...uploadForm, name: e.target.value })}
                />
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadVersion')}</span>
                <input
                  value={uploadForm.version}
                  placeholder="1.0.0"
                  onChange={(e) => setUploadForm({ ...uploadForm, version: e.target.value })}
                />
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadAuthor')}</span>
                <input
                  value={uploadForm.author}
                  placeholder="Community"
                  onChange={(e) => setUploadForm({ ...uploadForm, author: e.target.value })}
                />
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadCategory')}</span>
                <select
                  value={uploadForm.category}
                  onChange={(e) =>
                    setUploadForm({ ...uploadForm, category: e.target.value as MpCategory })
                  }
                >
                  {mpCategories.map((c) => {
                    const key = ('mpCat' + c.id[0].toUpperCase() + c.id.slice(1)) as StringKey
                    return (
                      <option key={c.id} value={c.id}>
                        {t(key)}
                      </option>
                    )
                  })}
                </select>
              </label>
              <label className="set-mp-field">
                <span>{t('mpUploadIcon')}</span>
                <input
                  value={uploadForm.icon}
                  maxLength={2}
                  placeholder="🚀"
                  onChange={(e) => setUploadForm({ ...uploadForm, icon: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadDesc')}</span>
                <input
                  value={uploadForm.description}
                  placeholder={t('mpUploadDesc')}
                  onChange={(e) => setUploadForm({ ...uploadForm, description: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadTools')}</span>
                <input
                  value={uploadForm.tools}
                  placeholder="tool_a, tool_b"
                  onChange={(e) => setUploadForm({ ...uploadForm, tools: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadScopes')}</span>
                <input
                  value={uploadForm.scopes}
                  placeholder="files:read, network:out"
                  onChange={(e) => setUploadForm({ ...uploadForm, scopes: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadTags')}</span>
                <input
                  value={uploadForm.tags}
                  placeholder="tag1, tag2"
                  onChange={(e) => setUploadForm({ ...uploadForm, tags: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadLongDesc')}</span>
                <textarea
                  value={uploadForm.longDescription}
                  placeholder={t('mpUploadLongDescHint')}
                  maxLength={1024}
                  rows={3}
                  onChange={(e) => setUploadForm({ ...uploadForm, longDescription: e.target.value })}
                />
              </label>
              <label className="set-mp-field set-mp-field-wide">
                <span>{t('mpUploadHomepage')}</span>
                <input
                  value={uploadForm.homepage}
                  placeholder="https://github.com/you/your-extension"
                  onChange={(e) => setUploadForm({ ...uploadForm, homepage: e.target.value })}
                />
              </label>
              <label
                className="set-mp-field set-mp-field-wide set-mp-artifact"
                data-artifact={uploadArtifact ? '1' : '0'}
                onDragOver={(e) => { e.preventDefault() }}
                onDrop={(e) => {
                  e.preventDefault()
                  const file = e.dataTransfer.files?.[0]
                  if (file) void handleArtifactFile(file)
                }}
              >
                <span>{t('mpUploadArtifact')}</span>
                <div className="set-mp-artifact-row">
                  <input
                    ref={(el) => { artifactInputRef.current = el }}
                    type="file"
                    accept=".md,.ts,.js,.mts,.mjs,text/markdown,application/typescript,text/javascript"
                    style={{ display: 'none' }}
                    onChange={(e) => {
                      const file = e.target.files?.[0]
                      if (file) void handleArtifactFile(file)
                      e.target.value = ''
                    }}
                  />
                  <button
                    type="button"
                    className="set-btn set-btn-mini"
                    onClick={() => artifactInputRef.current?.click()}
                  >
                    {t('mpUploadPickFile')}
                  </button>
                  {uploadArtifact ? (
                    <>
                      <span className="set-mp-artifact-chip" title={uploadArtifact.filename}>
                        <code>{uploadArtifact.filename}</code>
                        <span className="set-mp-artifact-size">{readableBytes(uploadArtifact.size)}</span>
                      </span>
                      <button
                        type="button"
                        className="set-btn-mini"
                        onClick={() => setUploadArtifact(null)}
                      >
                        {t('mpUploadClearFile')}
                      </button>
                      <span className="set-mp-artifact-badge">
                        {uploadKind === 'skill' ? t('mpUploadGuidanceOnly') : t('mpUploadExtensionCode')}
                      </span>
                    </>
                  ) : (
                    <span className="set-mp-artifact-empty">{t('mpUploadArtifactHint')}</span>
                  )}
                </div>
              </label>
            </div>
            <div className="set-mp-upload-actions">
              <button
                type="button"
                className="set-btn set-btn-primary"
                disabled={uploading}
                onClick={() => void submitUpload()}
              >
                {uploading ? '…' : t('mpUploadSubmit')}
              </button>
              <button
                type="button"
                className="set-btn"
                onClick={() => { setShowUpload(false); setUploadMsg(null) }}
              >
                {t('mpUploadCancel')}
              </button>
            </div>
            {uploadMsg && (
              <div className="set-install-msg" data-upload-msg={uploadMsg.kind}>
                {uploadMsg.kind === 'ok' ? `✓ ${uploadMsg.text}` : `✗ ${uploadMsg.text}`}
              </div>
            )}

            <div className="set-mp-upload-history">
              <h6>{t('mpUploadHistory')} · {uploadUploads.length}</h6>
              {uploadUploads.length === 0 ? (
                <div className="set-field-desc">{t('mpUploadEmpty')}</div>
              ) : (
                <ul className="set-skill-list" role="list">
                  {uploadUploads.map((u) => (
                    <li key={u.file} className="set-mp-upload-item" data-upload-id={u.id} data-upload-kind={u.kind}>
                      <code>{u.file}</code>
                      <span className="set-mp-upload-name">{u.name ?? u.id ?? ''}</span>
                      <span className="set-mp-upload-kind">{u.kind ?? ''}</span>
                      {u.artifact && (
                        <span
                          className={`set-mp-artifact-chip ${u.artifact.kind === 'extension' ? 'is-code' : 'is-md'}`}
                          title={u.artifact.filename}
                        >
                          <code>{u.artifact.filename}</code>
                          <span className="set-mp-artifact-size">{u.artifact.size ?? ''}</span>
                        </span>
                      )}
                      {u.reviewStatus && u.reviewStatus !== 'published' && (
                        <span className="set-mp-upload-pending">{t('mpUploadPending')}</span>
                      )}
                      <span className="set-mp-spacer" />
                      <button
                        type="button"
                        className="set-btn-mini set-btn-uninstall"
                        disabled={unpublishingId === u.id}
                        onClick={() => u.id && u.kind && void unpublishUpload(u.kind as 'skill' | 'plugin', u.id)}
                      >
                        {unpublishingId === u.id ? '…' : t('mpUploadUnpublish')}
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </div>
        )}

        {marketMsg && <div className="set-install-msg" data-market-msg="1">{marketMsg}</div>}

        {/* ── Results grid ───────────────────────────────────────── */}
        <div className="set-mp-layout">
          <div className="set-mp-results">
            {/* Result summary — tells the user how many entries the current
                query/filters matched and which filters are active, so an
                empty result is never a dead end. */}
            <div className="set-mp-summary" data-mp-summary="1">
              {(() => {
                const active: string[] = []
                if (mpQ) active.push(`"${mpQ}"`)
                if (mpCategory) {
                  const key = ('mpCat' + mpCategory[0].toUpperCase() + mpCategory.slice(1)) as StringKey
                  active.push(t(key))
                }
                if (mpType !== 'all') active.push(mpType === 'skill' ? t('mpTypeSkill') : t('mpTypePlugin'))
                if (mpInstalled === true) active.push(t('mpOnlyInstalled'))
                if (mpInstalled === false) active.push(t('mpOnlyAvailable'))
                if (minRating > 0) active.push(`★ ${minRating}+`)
                const shown = mpSkills.length + mpPlugins.length
                return (
                  <>
                    <span className="set-mp-summary-count">
                      {t('mpShowing', { shown, total: mpTotal })}
                    </span>
                    {active.length > 0 && (
                      <span className="set-mp-summary-filters">
                        {active.map((a) => (
                          <span key={a} className="set-mp-summary-chip">{a}</span>
                        ))}
                      </span>
                    )}
                  </>
                )
              })()}
            </div>
            {mpSkills.length + mpPlugins.length === 0 ? (
              <div className="set-mp-empty" data-mp-empty="1">
                <div className="set-mp-empty-title">{t('mpNoResults')}</div>
                <div className="set-mp-empty-hint">{t('mpNoResultsHint')}</div>
                {(mpQ || mpCategory || mpType !== 'all' || mpInstalled !== 'all' || minRating > 0) && (
                  <button
                    type="button"
                    className="set-btn set-btn-primary set-mp-empty-reset"
                    onClick={() => {
                      setMpQ('')
                      setMpCategory('')
                      setMpType('all')
                      setMpInstalled('all')
                      setMinRating(0)
                    }}
                  >
                    {t('mpResetFilters')}
                  </button>
                )}
              </div>
            ) : (
              <div className="set-mp-grid" data-mp-grid="1">
                {mpSkills.map((m) => renderMpCard(m, 'skill'))}
                {mpPlugins.map((m) => renderMpCard(m, 'plugin'))}
              </div>
            )}
          </div>

          {detailEntry && (
            <aside className="set-mp-detail" data-mp-detail={detailEntry.entry.id}>
              <div className="set-mp-detail-head">
                <span className="set-mp-card-icon">{detailEntry.entry.icon ?? detailEntry.entry.name[0]}</span>
                <div>
                  <div className="set-mp-detail-name">{detailEntry.entry.name}</div>
                  <div className="set-mp-detail-meta">
                    v{detailEntry.entry.version} · {detailEntry.entry.author}
                  </div>
                </div>
                <button
                  type="button"
                  className="set-btn-mini"
                  onClick={() => setDetailEntry(null)}
                  aria-label={t('mpClose')}
                >
                  ×
                </button>
              </div>
              <p className="set-mp-detail-desc">
                {detailEntry.entry.longDescription || detailEntry.entry.description}
              </p>
              <div className="set-mp-detail-stats">
                <span className="set-mp-card-rating">★ {detailEntry.entry.rating.toFixed(1)}</span>
                <span>
                  {detailEntry.entry.downloads.toLocaleString()} {t('mpDownloads')}
                </span>
              </div>
              <div className="set-mp-detail-block">
                <h6>{t('tools')} · {detailEntry.entry.tools.length}</h6>
                <div className="set-skill-scopes">
                  {detailEntry.entry.tools.map((tool) => (
                    <span key={tool} className="set-scope-tag">{tool}</span>
                  ))}
                </div>
              </div>
              <div className="set-mp-detail-block">
                <h6>{t('mpScopes')}</h6>
                <div className="set-skill-scopes">
                  {detailEntry.entry.scopes.map((sc) => (
                    <span key={sc} className="set-scope-tag">{sc}</span>
                  ))}
                </div>
              </div>
              {detailEntry.kind === 'plugin' && (detailEntry.entry as { requirements?: string[] }).requirements && (detailEntry.entry as { requirements?: string[] }).requirements!.length > 0 && (
                <div className="set-mp-detail-block">
                  <h6>{t('requirements')}</h6>
                  <div className="set-skill-scopes">
                    {(detailEntry.entry as { requirements?: string[] }).requirements!.map((r) => (
                      <span key={r} className="set-scope-tag set-scope-tag-req">{r}</span>
                    ))}
                  </div>
                </div>
              )}
              <div className="set-mp-detail-block">
                <h6>{t('mpCategory')}</h6>
                <div className="set-mp-card-tags">
                  {detailEntry.entry.tags.map((tag) => (
                    <span key={tag} className="set-mp-tag">{tag}</span>
                  ))}
                </div>
              </div>
              <div className="set-mp-detail-meta">
                <code>{detailEntry.entry.package}</code>
                {detailEntry.entry.homepage && (
                  <a href={detailEntry.entry.homepage} target="_blank" rel="noopener noreferrer">
                    {detailEntry.entry.homepage}
                  </a>
                )}
              </div>
              {detailEntry.pi && (
                <div className="set-mp-detail-block set-mp-pi-detail" data-pi-detail="1">
                  <h6>{t('mpCardPiDetail')}</h6>
                  <ul className="set-mp-pi-detail-list" role="list">
                    {detailEntry.kind === 'plugin' ? (
                      <>
                        {(detailEntry.pi as { packageDir?: string }).packageDir && (
                          <li>
                            <span>packageDir</span>
                            <code>{(detailEntry.pi as { packageDir?: string }).packageDir}</code>
                          </li>
                        )}
                        {(detailEntry.pi as { mode?: string }).mode && (
                          <li>
                            <span>mode</span>
                            <code>{(detailEntry.pi as { mode?: string }).mode}</code>
                          </li>
                        )}
                        <li>
                          <span>code</span>
                          <code>
                            {(detailEntry.pi as { hasCode?: boolean }).hasCode
                              ? t('mpCardHasCode')
                              : t('mpCardGuidance')}
                          </code>
                        </li>
                        <li>
                          <span>installed</span>
                          <code>
                            {(detailEntry.pi as { installed?: boolean }).installed
                              ? '✓'
                              : '—'}
                          </code>
                        </li>
                      </>
                    ) : (
                      <>
                        {(detailEntry.pi as { skillPath?: string | null }).skillPath && (
                          <li>
                            <span>SKILL.md</span>
                            <code>{(detailEntry.pi as { skillPath?: string | null }).skillPath}</code>
                          </li>
                        )}
                        <li>
                          <span>enabled</span>
                          <code>
                            {(detailEntry.pi as { enabled?: boolean }).enabled ? '✓' : '—'}
                          </code>
                        </li>
                      </>
                    )}
                  </ul>
                </div>
              )}
              {detailEntry.entry.artifact && (
                <div className="set-mp-detail-block">
                  <h6>{t('mpUploadArtifact')}</h6>
                  <div className="set-mp-artifact-chip" title={detailEntry.entry.artifact.filename}>
                    <code>{detailEntry.entry.artifact.filename}</code>
                    <span className="set-mp-artifact-size">{detailEntry.entry.artifact.size ?? ''}</span>
                  </div>
                </div>
              )}
              {detailEntry.entry.artifact && (
                <button
                  type="button"
                  className="set-btn-mini set-btn-uninstall set-mp-detail-unpublish"
                  disabled={!!unpublishingId && unpublishingId === detailEntry.entry.id}
                  onClick={() => void unpublishUpload(detailEntry.kind, detailEntry.entry.id)}
                >
                  {t('mpUploadUnpublish')}
                </button>
              )}
              {/* Rate widget — only show for community uploads (rating counts > 0
                  or the entry has 0 ratings but is a non-builtin upload) */}
              <div className="set-mp-detail-block set-mp-rate-block">
                <h6>{t('mpRateTitle')}</h6>
                <p className="set-mp-rate-hint">{t('mpRateHint')}</p>
                <div className="set-mp-rate-row">
                  {[1, 2, 3, 4, 5].map((n) => (
                    <button
                      key={n}
                      type="button"
                      className={`set-mp-rate-star${n <= rateDraft ? ' is-on' : ''}`}
                      aria-label={`${n} 星`}
                      disabled={rateBusy}
                      onClick={() => setRateDraft(n)}
                    >
                      ★
                    </button>
                  ))}
                  <button
                    type="button"
                    className="set-btn set-btn-primary set-mp-rate-submit"
                    disabled={rateBusy}
                    onClick={() => void submitRating()}
                  >
                    {rateBusy ? '…' : t('mpRateSubmit')}
                  </button>
                </div>
                {rateMsg && (
                  <div className={`set-mp-rate-msg set-mp-rate-msg-${rateMsg.kind}`}>{rateMsg.text}</div>
                )}
              </div>
            </aside>
          )}
        </div>
      </div>
    </>
  )
}



/** Module-manager pane: toggle + drag-to-reorder the home-page Quick Start cards. */
function ModulesPane({ t }: { t: TFunc }) {
  const [modules, setModules] = useState<ModuleEntry[]>([])
  const [loading, setLoading] = useState(true)
  const dragIndexRef = useRef<number | null>(null)

  useEffect(() => {
    let alive = true
    void window.aiOffice.listModules?.().then((m) => {
      if (!alive) return
      setModules(m)
      setLoading(false)
    })
    return () => {
      alive = false
    }
  }, [])

  if (loading) return null

  const toggle = async (id: string, enabled: boolean) => {
    const next = await window.aiOffice.setModuleEnabled?.(id as ModuleKind, enabled)
    if (Array.isArray(next)) {
      setModules(next)
      window.dispatchEvent(new CustomEvent('genoffice:modules-changed', { detail: next }))
    }
  }

  const move = async (from: number, to: number) => {
    if (from === to) return
    const next = modules.slice()
    const [item] = next.splice(from, 1)
    next.splice(to, 0, item)
    setModules(next)
    await window.aiOffice.reorderModules?.(next.map((m) => m.id))
    window.dispatchEvent(new CustomEvent('genoffice:modules-changed', { detail: next }))
  }

  const reset = async () => {
    const next = await window.aiOffice.resetModules?.()
    if (Array.isArray(next)) {
      setModules(next)
      window.dispatchEvent(new CustomEvent('genoffice:modules-changed', { detail: next }))
    }
  }

  const onDragStart = (index: number) => (e: React.DragEvent) => {
    dragIndexRef.current = index
    e.dataTransfer.effectAllowed = 'move'
  }
  const onDragOver = (index: number) => (e: React.DragEvent) => {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'move'
  }
  const onDrop = (index: number) => (e: React.DragEvent) => {
    e.preventDefault()
    const from = dragIndexRef.current
    dragIndexRef.current = null
    if (from === null) return
    void move(from, index)
  }

  return (
    <>
      <h3 className="set-pane-title">{t('setSecModules')}</h3>
      <div className="set-field-stack" style={{ marginBottom: 14 }}>
        <div className="set-field-label">{t('modulesTitle')}</div>
        <div className="set-field-desc">{t('modulesDesc')}</div>
      </div>
      <ul className="set-module-list" role="list">
        {modules.map((m, idx) => (
          <li
            key={m.id}
            className="set-module-row"
            draggable
            onDragStart={onDragStart(idx)}
            onDragOver={onDragOver(idx)}
            onDrop={onDrop(idx)}
          >
            <span className="set-module-handle" aria-hidden="true">
              <svg width="10" height="14" viewBox="0 0 10 14" fill="none">
                <circle cx="3" cy="3" r="1" fill="currentColor" />
                <circle cx="3" cy="7" r="1" fill="currentColor" />
                <circle cx="3" cy="11" r="1" fill="currentColor" />
                <circle cx="7" cy="3" r="1" fill="currentColor" />
                <circle cx="7" cy="7" r="1" fill="currentColor" />
                <circle cx="7" cy="11" r="1" fill="currentColor" />
              </svg>
            </span>
            <span className="set-module-label">{t(m.labelKey as StringKey)}</span>
            <span className="set-module-ext">.{m.ext}</span>
            <button
              className="set-switch"
              role="switch"
              aria-checked={m.enabled}
              aria-label={t(m.labelKey as StringKey)}
              onClick={() => void toggle(m.id, !m.enabled)}
            />
          </li>
        ))}
      </ul>
      <div className="set-field" style={{ marginTop: 12 }}>
        <div className="set-field-text" />
        <button className="set-btn" data-tip={t('modulesResetTip')} onClick={() => void reset()}>
          {t('modulesReset')}
        </button>
      </div>
    </>
  )
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
  skillUpdateDue: updateDue = false,
  onSkillUpdateDue,
}: SettingsModalProps) {
  const { lang, setLang, t } = useI18n()
  const [section, setSection] = useState<SectionId>('account')
  const [theme, setTheme] = useState<UiTheme>('system')
  const [saveDir, setSaveDir] = useState('')
  const [analyticsOn, setAnalyticsOn] = useState(true)
  const [analyticsSaving, setAnalyticsSaving] = useState(false)
  const [autoSaveOn, setAutoSaveOn] = useState(false)
  const [aiPrefs, setAiPrefs] = useState<AiPanelPrefs>(DEFAULT_AI_PANEL_PREFS)
  const [channel, setChannel] = useState<'stable' | 'beta'>('stable')
  const [appVersion, setAppVersion] = useState('')
  const [githubStars, setGithubStars] = useState<number | null>(null)

  useEffect(() => {
    let alive = true
    void window.aiOffice.getTheme?.().then((th) => {
      if (alive) setTheme(th)
    })
    void window.aiOffice.getDefaultSaveDir?.().then((dir) => {
      if (alive && dir) setSaveDir(dir)
    })
    void window.aiOffice.getAnalyticsEnabled?.().then((on) => {
      if (alive) setAnalyticsOn(on !== false)
    })
    void window.aiOffice.getAutoSaveDefault?.().then((v) => {
      if (alive) setAutoSaveOn(v.on)
    })
    void window.aiOffice.getAiPanelPrefs?.().then((prefs) => {
      if (alive) setAiPrefs(prefs)
    })
    void window.aiOffice.getUpdateChannel?.().then((ch) => {
      if (alive) setChannel(ch)
    })
    void window.aiOffice.getAppVersion?.().then((v) => {
      if (alive && v) setAppVersion(v)
    })
    void window.aiOffice.githubStars?.().then((n) => {
      if (alive && n !== null) setGithubStars(n)
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

  const updateAiPrefs = (patch: Partial<AiPanelPrefs>) => {
    setAiPrefs((prev) => ({ ...prev, ...patch }))
    void window.aiOffice.setAiPanelPrefs(patch).then(setAiPrefs)
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
                {s.id === 'integrations' && updateDue && (
                  <span className="set-nav-dot" role="img" aria-label={t('intgUpdateDue')} />
                )}
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
            {section === 'aiModel' && <AiModelPane t={t} />}
            {section === 'aiMedia' && <AiMediaPane t={t} />}
            {section === 'translationKb' && <TranslationKbPane t={t} />}
            {section === 'modules' && <ModulesPane t={t} />}
            {section === 'skillsPlugins' && <SkillsPluginsPane t={t} />}
            {section === 'general' && (
              <>
                <h3 className="set-pane-title">{t('setSecGeneral')}</h3>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('language')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={lang}
                    ariaLabel={t('language')}
                    options={LANG_OPTIONS.map((opt) => ({ value: opt.value, label: opt.label }))}
                    onPick={(v) => setLang(v as typeof lang)}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('theme')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={theme}
                    ariaLabel={t('theme')}
                    options={THEME_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    onPick={(v) => applyTheme(v as UiTheme)}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('setAiFontSize')}</label>
                  </div>
                  {aiPrefs.fontSize === 'custom' && (
                    <CustomFontSizeInput
                      value={aiPrefs.customFontSize}
                      label={t('aiFontSizeCustom')}
                      onCommit={(px) => updateAiPrefs({ customFontSize: px })}
                    />
                  )}
                  <Dropdown
                    className="set-dd"
                    value={aiPrefs.fontSize}
                    ariaLabel={t('setAiFontSize')}
                    options={AI_FONT_SIZE_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    onPick={(v) => {
                      const fontSize = v as AiFontSize
                      // start the custom size from the preset being left so nothing jumps
                      updateAiPrefs(
                        fontSize === 'custom' && aiPrefs.fontSize !== 'custom'
                          ? { fontSize, customFontSize: aiPanelFontPx(aiPrefs) }
                          : { fontSize },
                      )
                    }}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAiSpellcheck')}</div>
                      <div className="set-field-desc">{t('setAiSpellcheckDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={aiPrefs.spellcheck}
                    aria-label={t('setAiSpellcheck')}
                    onClick={() => updateAiPrefs({ spellcheck: !aiPrefs.spellcheck })}
                  />
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
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAutoSave')}</div>
                      <div className="set-field-desc">{t('setAutoSaveDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={autoSaveOn}
                    aria-label={t('setAutoSave')}
                    onClick={() => {
                      const next = !autoSaveOn
                      setAutoSaveOn(next)
                      void window.aiOffice.setAutoSaveDefault?.(next).catch(() => {})
                    }}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAnalytics')}</div>
                      <div className="set-field-desc">{t('setAnalyticsDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={analyticsOn}
                    aria-label={t('setAnalytics')}
                    disabled={analyticsSaving}
                    onClick={() => {
                      const next = !analyticsOn
                      setAnalyticsSaving(true)
                      void window.aiOffice
                        .setAnalyticsEnabled(next)
                        .then((persisted) => {
                          if (persisted) setAnalyticsOn(next)
                        })
                        .catch(() => {})
                        .finally(() => setAnalyticsSaving(false))
                    }}
                  />
                </div>
              </>
            )}
            {section === 'integrations' && (
              <IntegrationsPane t={t} onStatus={(st) => onSkillUpdateDue?.(skillUpdateDue(st))} />
            )}
            {section === 'about' && (
              <>
                <h3 className="set-pane-title">{t('setSecAbout')}</h3>
                <Field label={t('versionLabel')} value={appVersion || '—'} />
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('updateChannel')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={channel}
                    ariaLabel={t('updateChannel')}
                    options={CHANNEL_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    onPick={(v) => {
                      const next = v === 'beta' ? 'beta' : 'stable'
                      setChannel(next)
                      void window.aiOffice.setUpdateChannel(next)
                    }}
                  />
                </div>
                <Field
                  label={t('setGithub')}
                  value={
                    githubStars === null
                      ? 'github.com/genspark-ai/genoffice'
                      : `github.com/genspark-ai/genoffice · ★ ${formatStars(githubStars)}`
                  }
                  action={
                    <button
                      className="set-btn"
                      onClick={() => void window.aiOffice.openGitHubRepo?.()}
                    >
                      {t('starOnGitHub')}
                    </button>
                  }
                />
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
