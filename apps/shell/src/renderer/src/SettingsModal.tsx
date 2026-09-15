import { useCallback, useEffect, useRef, useState } from 'react'
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
  MarketplaceUploadPayload,
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
  | 'general'
  | 'integrations'
  | 'modules'
  | 'skillsPlugins'
  | 'about'

const SECTIONS: readonly { id: SectionId; labelKey: StringKey }[] = [
  { id: 'account', labelKey: 'setSecAccount' },
  { id: 'aiModel', labelKey: 'setSecAiModel' },
  { id: 'aiMedia', labelKey: 'setSecAiMedia' },
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
    const meta = options.find((m) => m.id === current) ?? options[0]!
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

  return (
    <>
      <h3 className="set-pane-title">{t('setSecAiMedia')}</h3>
      <div className="set-field-desc set-ai-note">{t('setAiSharedKeyHint')}</div>
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
  const [detailEntry, setDetailEntry] = useState<MpEntry | null>(null)
  const [showUpload, setShowUpload] = useState(false)
  const [uploadKind, setUploadKind] = useState<'skill' | 'plugin'>('skill')
  const [uploadUploads, setUploadUploads] = useState<{ file: string; id?: string; name?: string; kind?: string; uploadedAt?: string }[]>([])
  const [uploadStatus, setUploadStatus] = useState<string | null>(null)
  const [uploadForm, setUploadForm] = useState({
    id: '',
    name: '',
    description: '',
    version: '1.0.0',
    tools: '',
    scopes: '',
    category: 'productivity' as MpCategory,
    tags: '',
    author: '',
    icon: '',
  })
  const [uploadMsg, setUploadMsg] = useState<{ kind: 'ok' | 'err'; text: string } | null>(null)
  const [uploading, setUploading] = useState(false)

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
      setUploadUploads((r as { uploads?: typeof uploadUploads }).uploads ?? [])
    })
  }, [showUpload])

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
    setDetailEntry(entry)
    const res = await window.aiOffice.marketplaceDetail?.(entry.id, kind)
    if (res?.ok && res.entry) setDetailEntry(res.entry)
  }

  const submitUpload = async () => {
    setUploadMsg(null)
    setUploading(true)
    try {
      const payload: MarketplaceUploadPayload = {
        id: uploadForm.id.trim(),
        name: uploadForm.name.trim(),
        description: uploadForm.description.trim(),
        version: uploadForm.version.trim(),
        tools: splitList(uploadForm.tools),
        scopes: splitList(uploadForm.scopes),
        category: uploadForm.category,
        tags: splitList(uploadForm.tags),
        author: uploadForm.author.trim() || undefined,
        icon: uploadForm.icon.trim() || undefined,
      }
      const res = await window.aiOffice.marketplaceUpload?.(uploadKind, payload)
      if (res?.ok) {
        setUploadMsg({ kind: 'ok', text: res.message ?? '✓' })
        setUploadForm({
          id: '',
          name: '',
          description: '',
          version: '1.0.0',
          tools: '',
          scopes: '',
          category: 'productivity',
          tags: '',
          author: '',
          icon: '',
        })
        const listed = await window.aiOffice.marketplaceListUploads?.()
        setUploadUploads((listed as { uploads?: typeof uploadUploads })?.uploads ?? [])
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

  /** marketplace grid card — icon, rating, tags, install/uninstall + detail */
  const renderMpCard = (m: MpEntry, kind: 'skill' | 'plugin') => {
    const catKey = ('mpCat' + m.category[0].toUpperCase() + m.category.slice(1)) as StringKey
    return (
      <article
        key={`${kind}:${m.id}`}
        className="set-mp-card"
        data-mp-id={m.id}
        data-mp-kind={kind}
        data-installed={m.installed ? '1' : '0'}
      >
        <header className="set-mp-card-head">
          <span className="set-mp-card-icon">{m.icon || m.name[0]}</span>
          <div className="set-mp-card-title">
            <span className="set-mp-card-name">{m.name}</span>
            <span className="set-mp-card-sub">v{m.version} · {m.author}</span>
          </div>
          {m.featured && <span className="set-mp-featured">{t('mpFeatured')}</span>}
        </header>

        <p className="set-mp-card-desc">{m.description}</p>

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

      <div className="set-skill-section set-mp" data-marketplace="v2">
        <div className="set-skill-section-head">
          <h4>{t('marketplace')}</h4>
          <span className="set-skill-count">{mpTotal}</span>
          <span className="set-mp-spacer" />
          <button
            type="button"
            className="set-btn set-btn-primary set-mp-publish"
            onClick={() => { setShowUpload((v) => !v); setUploadMsg(null) }}
          >
            {showUpload ? t('mpClose') : t('mpUpload')}
          </button>
        </div>

        {/* ── Search bar ─────────────────────────────────────────── */}
        <div className="set-mp-searchrow">
          <span className="set-mp-searchicon" aria-hidden="true">
            <svg width="14" height="14" viewBox="0 0 14 14">
              <circle cx="6" cy="6" r="4.4" stroke="currentColor" strokeWidth="1.5" fill="none" />
              <path d="M9.4 9.4l3 3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
          </span>
          <input
            type="text"
            className="set-input set-mp-search"
            placeholder={t('mpSearchPlaceholder')}
            value={mpQ}
            onChange={(e) => setMpQ(e.target.value)}
            aria-label={t('mpSearchPlaceholder')}
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
              }}
            >
              {t('mpResetFilters')}
            </button>
          )}
        </div>

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
                    <li key={u.file} className="set-mp-upload-item">
                      <code>{u.file}</code>
                      <span>{u.name ?? u.id ?? ''}</span>
                      <span className="set-mp-upload-kind">{u.kind ?? ''}</span>
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
            {mpSkills.length + mpPlugins.length === 0 ? (
              <div className="set-mp-empty">{t('mpNoResults')}</div>
            ) : (
              <div className="set-mp-grid" data-mp-grid="1">
                {mpSkills.map((m) => renderMpCard(m, 'skill'))}
                {mpPlugins.map((m) => renderMpCard(m, 'plugin'))}
              </div>
            )}
          </div>

          {detailEntry && (
            <aside className="set-mp-detail" data-mp-detail={detailEntry.id}>
              <div className="set-mp-detail-head">
                <span className="set-mp-card-icon">{detailEntry.icon ?? detailEntry.name[0]}</span>
                <div>
                  <div className="set-mp-detail-name">{detailEntry.name}</div>
                  <div className="set-mp-detail-meta">
                    v{detailEntry.version} · {detailEntry.author}
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
                {detailEntry.longDescription || detailEntry.description}
              </p>
              <div className="set-mp-detail-stats">
                <span className="set-mp-card-rating">★ {detailEntry.rating.toFixed(1)}</span>
                <span>
                  {detailEntry.downloads.toLocaleString()} {t('mpDownloads')}
                </span>
              </div>
              <div className="set-mp-detail-block">
                <h6>{t('tools')} · {detailEntry.tools.length}</h6>
                <div className="set-skill-scopes">
                  {detailEntry.tools.map((tool) => (
                    <span key={tool} className="set-scope-tag">{tool}</span>
                  ))}
                </div>
              </div>
              <div className="set-mp-detail-block">
                <h6>{t('mpScopes')}</h6>
                <div className="set-skill-scopes">
                  {detailEntry.scopes.map((sc) => (
                    <span key={sc} className="set-scope-tag">{sc}</span>
                  ))}
                </div>
              </div>
              {detailEntry.requirements && detailEntry.requirements.length > 0 && (
                <div className="set-mp-detail-block">
                  <h6>{t('requirements')}</h6>
                  <div className="set-skill-scopes">
                    {detailEntry.requirements.map((r) => (
                      <span key={r} className="set-scope-tag set-scope-tag-req">{r}</span>
                    ))}
                  </div>
                </div>
              )}
              <div className="set-mp-detail-block">
                <h6>{t('mpCategory')}</h6>
                <div className="set-mp-card-tags">
                  {detailEntry.tags.map((tag) => (
                    <span key={tag} className="set-mp-tag">{tag}</span>
                  ))}
                </div>
              </div>
              <div className="set-mp-detail-meta">
                <code>{detailEntry.package}</code>
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
