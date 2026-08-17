import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import {
  AI_PROVIDERS,
  type AiProviderId,
  type AiSettings,
  type CodexAccountStatus,
  type CodexCapabilities,
  type CodexErrorCode,
  type CodexModelCapability,
  type CodexReasoningEffort,
  type CodexServiceTier,
  type CodexUiLabels,
} from '@genoffice/ai-provider'

export interface CodexCapabilitiesResult extends CodexCapabilities {
  errorCode?: CodexErrorCode
}

export interface AiProviderControlApi {
  aiCodexStatus(): Promise<CodexAccountStatus>
  aiCodexLogin(): Promise<CodexAccountStatus>
  aiCodexCancelLogin?(): Promise<void> | void
  aiCodexLogout?(): Promise<CodexAccountStatus>
  aiCodexCapabilities(): Promise<CodexCapabilitiesResult>
}

export type AiProviderControlLabels = CodexUiLabels

export interface UseAiProviderControlsOptions {
  settings: AiSettings
  onSettingsChange?: (settings: AiSettings) => void
  api: AiProviderControlApi
  positionKey?: unknown
}

export type CodexPopover = 'model' | 'effort' | 'speed' | null

const PROVIDERS = AI_PROVIDERS.filter(
  (provider) => provider.id === 'genspark' || provider.id === 'openai-codex',
)
const DEFAULT_CODEX_SERVICE_TIER: CodexServiceTier = { id: 'default', name: 'Standard' }

function codexServiceTiers(model: CodexModelCapability): CodexServiceTier[] {
  const tiers = model.serviceTiers?.length ? model.serviceTiers : [DEFAULT_CODEX_SERVICE_TIER]
  return tiers.some((tier) => tier.id === 'default')
    ? tiers
    : [DEFAULT_CODEX_SERVICE_TIER, ...tiers]
}

function selectedCodexServiceTier(model: CodexModelCapability, value?: string): string {
  const tiers = codexServiceTiers(model)
  if (value && tiers.some((tier) => tier.id === value)) return value
  if (model.defaultServiceTier && tiers.some((tier) => tier.id === model.defaultServiceTier)) {
    return model.defaultServiceTier
  }
  return tiers[0]?.id ?? 'default'
}

function Caret({ up = false }: { up?: boolean }) {
  return (
    <svg viewBox="0 0 12 12" width="12" height="12" fill="none" aria-hidden>
      <path
        d={up ? 'M2.5 7.5 6 4l3.5 3.5' : 'm2.5 4.5 3.5 3 3.5-3'}
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function Check() {
  return (
    <svg viewBox="0 0 20 20" width="16" height="16" fill="none" aria-hidden>
      <path d="m4 10 3.5 3.5L16 5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    </svg>
  )
}

function Reset() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" fill="none" aria-hidden>
      <path
        d="M3.7 9.9A8.5 8.5 0 1 1 4.9 17M2 8.5l1.4 2.3 2.3-1.4"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

export function useAiProviderControls({
  settings,
  onSettingsChange,
  api,
  positionKey,
}: UseAiProviderControlsOptions) {
  const [codexAccount, setCodexAccount] = useState<CodexAccountStatus | null>(null)
  const [codexLoginPending, setCodexLoginPending] = useState(false)
  const [codexCapabilities, setCodexCapabilities] = useState<CodexCapabilitiesResult | null>(null)
  const [activePopover, setActivePopover] = useState<CodexPopover>(null)
  const settingsRef = useRef(settings)
  const onSettingsChangeRef = useRef(onSettingsChange)
  const apiRef = useRef(api)
  settingsRef.current = settings
  onSettingsChangeRef.current = onSettingsChange
  apiRef.current = api

  const accountRef = useRef<CodexAccountStatus | null>(null)
  const loginPendingRef = useRef(false)
  const selectionRef = useRef(0)
  const statusRequestRef = useRef(0)
  const controlRef = useRef<HTMLDivElement>(null)
  const triggerRef = useRef<HTMLButtonElement>(null)
  const pickerRef = useRef<HTMLDivElement>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const submenuRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const selection = ++selectionRef.current
    const statusRequest = ++statusRequestRef.current
    accountRef.current = null
    setCodexAccount(null)
    loginPendingRef.current = false
    setCodexLoginPending(false)
    setActivePopover(null)
    if (settings.provider !== 'openai-codex') {
      setCodexCapabilities(null)
      return
    }
    let active = true
    setCodexCapabilities(null)
    void apiRef.current
      .aiCodexStatus()
      .then((account) => {
        if (
          !active ||
          selection !== selectionRef.current ||
          statusRequest !== statusRequestRef.current
        )
          return
        accountRef.current = account
        setCodexAccount(account)
      })
      .catch(() => {
        if (
          !active ||
          selection !== selectionRef.current ||
          statusRequest !== statusRequestRef.current
        )
          return
        const account = { loggedIn: false, errorCode: 'provider-failure' as const }
        accountRef.current = account
        setCodexAccount(account)
      })
    return () => {
      active = false
    }
  }, [api, settings.provider])

  useEffect(() => {
    if (settings.provider !== 'openai-codex' || codexAccount?.loggedIn !== true) {
      setCodexCapabilities(null)
      return
    }
    let active = true
    void apiRef.current
      .aiCodexCapabilities()
      .then((capabilities) => {
        if (!active) return
        setCodexCapabilities((previous) =>
          capabilities.errorCode && previous?.models.length
            ? { models: previous.models, errorCode: capabilities.errorCode }
            : capabilities,
        )
        const currentSettings = settingsRef.current
        const selected = capabilities.models.find(
          (candidate) => candidate.id === currentSettings.providers['openai-codex'].model,
        )
        const model = selected ?? capabilities.models[0]
        const change = onSettingsChangeRef.current
        if (!model || !change) return
        const current = currentSettings.providers['openai-codex']
        const reasoningEffort = model.reasoningEfforts.includes(current.reasoningEffort ?? 'none')
          ? (current.reasoningEffort ?? 'none')
          : (model.reasoningEfforts[0] ?? 'none')
        const serviceTier = selectedCodexServiceTier(model, current.serviceTier)
        if (
          model.id !== current.model ||
          reasoningEffort !== (current.reasoningEffort ?? 'none') ||
          serviceTier !== (current.serviceTier ?? 'default')
        ) {
          change({
            ...currentSettings,
            providers: {
              ...currentSettings.providers,
              'openai-codex': { ...current, model: model.id, reasoningEffort, serviceTier },
            },
          })
        }
      })
      .catch(() => {
        if (active) setCodexCapabilities({ models: [], errorCode: 'provider-failure' })
      })
    return () => {
      active = false
    }
  }, [api, codexAccount?.loggedIn, settings.provider])

  useEffect(() => {
    if (!activePopover) return
    const onPointerDown = (event: PointerEvent) => {
      if (!controlRef.current?.contains(event.target as Node)) setActivePopover(null)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setActivePopover(null)
    }
    document.addEventListener('pointerdown', onPointerDown)
    document.addEventListener('keydown', onKeyDown)
    return () => {
      document.removeEventListener('pointerdown', onPointerDown)
      document.removeEventListener('keydown', onKeyDown)
    }
  }, [activePopover])

  useLayoutEffect(() => {
    if (!activePopover) return
    const trigger = triggerRef.current
    const picker = pickerRef.current
    const root = rootRef.current
    const submenu = submenuRef.current
    if (!trigger || !picker || !root || !submenu) return

    const position = () => {
      const triggerRect = trigger.getBoundingClientRect()
      submenu.style.removeProperty('max-width')
      submenu.style.removeProperty('top')
      submenu.style.removeProperty('bottom')
      const rootRect = root.getBoundingClientRect()
      const activeRow = root.querySelector<HTMLElement>(`[data-codex-menu-item="${activePopover}"]`)
      const rowOffset = activeRow ? activeRow.getBoundingClientRect().top - rootRect.top : 0
      const submenuRect = submenu.getBoundingClientRect()
      const margin = 8
      const gap = 10
      const maxLeft = Math.max(margin, window.innerWidth - rootRect.width - margin)
      const left = Math.min(Math.max(margin, triggerRect.right - rootRect.width), maxLeft)
      const maxTop = Math.max(margin, window.innerHeight - rootRect.height - margin)
      const top = Math.min(Math.max(margin, triggerRect.top - rootRect.height - gap), maxTop)
      const rightSpace = window.innerWidth - margin - left - rootRect.width - gap
      const leftSpace = left - margin - gap
      const flipped = rightSpace < submenuRect.width && leftSpace > rightSpace
      const submenuSpace = flipped ? leftSpace : rightSpace

      picker.dataset.flipped = String(flipped)
      picker.style.left = `${left}px`
      picker.style.top = `${top}px`
      if (submenuSpace > 0 && submenuSpace < submenuRect.width) {
        submenu.style.maxWidth = `${submenuSpace}px`
      }
      const positionedSubmenuRect = submenu.getBoundingClientRect()
      const submenuTop = Math.min(
        Math.max(margin, top + rowOffset),
        Math.max(margin, window.innerHeight - positionedSubmenuRect.height - margin),
      )
      submenu.style.top = `${submenuTop - top}px`
      picker.dataset.positioned = 'true'
    }

    position()
    window.addEventListener('resize', position)
    return () => window.removeEventListener('resize', position)
  }, [activePopover, positionKey])

  const changeProvider = (provider: AiProviderId) => {
    const meta = PROVIDERS.find((candidate) => candidate.id === provider)
    const change = onSettingsChangeRef.current
    if (!meta || !change) return
    change({
      provider,
      providers: {
        ...settingsRef.current.providers,
        [provider]: {
          ...settingsRef.current.providers[provider],
          apiKey:
            meta.requiresApiKey === false
              ? ''
              : (settingsRef.current.providers[provider]?.apiKey ?? ''),
          model: settingsRef.current.providers[provider]?.model || meta.defaultModel,
          baseUrl: meta.needsBaseUrl
            ? (settingsRef.current.providers[provider]?.baseUrl ?? '')
            : undefined,
        },
      },
    })
  }

  const changeCodexModel = (model: string) => {
    const capability = codexCapabilities?.models.find((candidate) => candidate.id === model)
    const change = onSettingsChangeRef.current
    if (!capability || !change) return
    const currentSettings = settingsRef.current
    const current = currentSettings.providers['openai-codex']
    change({
      ...currentSettings,
      providers: {
        ...currentSettings.providers,
        'openai-codex': {
          ...current,
          model,
          reasoningEffort: capability.reasoningEfforts.includes(current.reasoningEffort ?? 'none')
            ? (current.reasoningEffort ?? 'none')
            : (capability.reasoningEfforts[0] ?? 'none'),
          serviceTier: selectedCodexServiceTier(capability, current.serviceTier),
        },
      },
    })
  }

  const changeCodexReasoning = (reasoningEffort: string) => {
    const change = onSettingsChangeRef.current
    if (!change || !codexCapabilities) return
    const currentSettings = settingsRef.current
    const model = codexCapabilities.models.find(
      (candidate) => candidate.id === currentSettings.providers['openai-codex'].model,
    )
    const effort = reasoningEffort as CodexReasoningEffort
    if (!model || !model.reasoningEfforts.includes(effort)) return
    change({
      ...currentSettings,
      providers: {
        ...currentSettings.providers,
        'openai-codex': { ...currentSettings.providers['openai-codex'], reasoningEffort: effort },
      },
    })
  }

  const changeCodexServiceTier = (serviceTier: string) => {
    const change = onSettingsChangeRef.current
    if (!change || !codexCapabilities) return
    const currentSettings = settingsRef.current
    const model = codexCapabilities.models.find(
      (candidate) => candidate.id === currentSettings.providers['openai-codex'].model,
    )
    if (!model || !codexServiceTiers(model).some((tier) => tier.id === serviceTier)) return
    change({
      ...currentSettings,
      providers: {
        ...currentSettings.providers,
        'openai-codex': { ...currentSettings.providers['openai-codex'], serviceTier },
      },
    })
  }

  const resetCodexSettings = () => {
    const change = onSettingsChangeRef.current
    if (!change || !codexCapabilities?.models.length) return
    const defaultModel =
      codexCapabilities.models.find(
        (model) =>
          model.id ===
          AI_PROVIDERS.find((provider) => provider.id === 'openai-codex')?.defaultModel,
      ) ?? codexCapabilities.models[0]
    if (!defaultModel) return
    const currentSettings = settingsRef.current
    const current = currentSettings.providers['openai-codex']
    const reasoningEffort = defaultModel.reasoningEfforts.includes('none')
      ? 'none'
      : (defaultModel.reasoningEfforts[0] ?? 'none')
    const serviceTier = selectedCodexServiceTier(
      defaultModel,
      defaultModel.defaultServiceTier ?? 'default',
    )
    change({
      ...currentSettings,
      providers: {
        ...currentSettings.providers,
        'openai-codex': {
          ...current,
          model: defaultModel.id,
          reasoningEffort,
          serviceTier,
        },
      },
    })
  }

  const togglePicker = () => {
    const trigger = triggerRef.current
    if (trigger) {
      trigger.style.width = `${trigger.getBoundingClientRect().width}px`
      void trigger.offsetWidth
    }
    setActivePopover((current) => (current ? null : 'model'))
    const schedule =
      typeof window.requestAnimationFrame === 'function'
        ? window.requestAnimationFrame.bind(window)
        : (callback: () => void) => window.setTimeout(callback, 0)
    schedule(() => triggerRef.current?.style.removeProperty('width'))
  }

  const startLogin = async () => {
    if (settingsRef.current.provider !== 'openai-codex' || loginPendingRef.current) return
    const selection = selectionRef.current
    loginPendingRef.current = true
    statusRequestRef.current += 1
    setCodexLoginPending(true)
    try {
      const result = await apiRef.current.aiCodexLogin()
      if (settingsRef.current.provider === 'openai-codex' && selection === selectionRef.current) {
        accountRef.current = result
        setCodexAccount(result)
      }
    } catch {
      if (settingsRef.current.provider === 'openai-codex' && selection === selectionRef.current) {
        const account = { loggedIn: false, errorCode: 'provider-failure' as const }
        accountRef.current = account
        setCodexAccount(account)
      }
    } finally {
      if (selection === selectionRef.current) {
        loginPendingRef.current = false
        setCodexLoginPending(false)
      }
    }
  }

  const refreshStatus = () => {
    if (settingsRef.current.provider !== 'openai-codex') return
    const selection = selectionRef.current
    const statusRequest = ++statusRequestRef.current
    void apiRef.current
      .aiCodexStatus()
      .then((account) => {
        if (selection !== selectionRef.current || statusRequest !== statusRequestRef.current) return
        accountRef.current = account
        setCodexAccount(account)
      })
      .catch(() => {
        if (selection !== selectionRef.current || statusRequest !== statusRequestRef.current) return
        const account = { loggedIn: false, errorCode: 'provider-failure' as const }
        accountRef.current = account
        setCodexAccount(account)
      })
  }

  const cancelLogin = () => {
    void apiRef.current.aiCodexCancelLogin?.()
    loginPendingRef.current = false
    setCodexLoginPending(false)
  }

  const activeCodexModel = codexCapabilities?.models.find(
    (candidate) => candidate.id === settings.providers['openai-codex'].model,
  )
  const activeCodexReasoning = settings.providers['openai-codex'].reasoningEffort ?? 'none'
  const activeCodexServiceTier = activeCodexModel
    ? selectedCodexServiceTier(activeCodexModel, settings.providers['openai-codex'].serviceTier)
    : 'default'
  const activeCodexSpeed = activeCodexModel
    ? (codexServiceTiers(activeCodexModel).find((tier) => tier.id === activeCodexServiceTier) ??
      DEFAULT_CODEX_SERVICE_TIER)
    : DEFAULT_CODEX_SERVICE_TIER
  const codexBannerError = codexAccount?.errorCode ?? codexCapabilities?.errorCode
  const showCodexBanner =
    settings.provider === 'openai-codex' &&
    Boolean(codexAccount) &&
    (codexAccount?.loggedIn === false || Boolean(codexBannerError))

  return {
    account: codexAccount,
    accountRef,
    loginPending: codexLoginPending,
    capabilities: codexCapabilities,
    activeCodexModel,
    activeCodexReasoning,
    activeCodexServiceTier,
    activeCodexSpeed,
    activePopover,
    setActivePopover,
    showCodexBanner,
    sendDisabled: settings.provider === 'openai-codex' && codexAccount?.loggedIn !== true,
    changeProvider,
    changeCodexModel,
    changeCodexReasoning,
    changeCodexServiceTier,
    resetCodexSettings,
    togglePicker,
    startLogin,
    refreshStatus,
    cancelLogin,
    controlRef,
    triggerRef,
    pickerRef,
    rootRef,
    submenuRef,
  }
}

export type AiProviderControlsState = ReturnType<typeof useAiProviderControls>

export function AiProviderSelect({
  settings,
  labels,
  controls,
}: {
  settings: AiSettings
  labels: AiProviderControlLabels
  controls: AiProviderControlsState
}) {
  return (
    <span className="ai-provider-select ai-provider-select-wrap">
      <span className="ai-provider-select-text" aria-hidden="true">
        {PROVIDERS.find((provider) => provider.id === settings.provider)?.label ??
          settings.provider}
      </span>
      <select
        className="ai-provider-select-input"
        aria-label={labels.providerSelect}
        value={settings.provider}
        onChange={(event) => controls.changeProvider(event.target.value as AiProviderId)}
      >
        {PROVIDERS.map((provider) => (
          <option key={provider.id} value={provider.id}>
            {provider.label}
          </option>
        ))}
      </select>
    </span>
  )
}

export function AiProviderAuthBanner({
  settings,
  labels,
  controls,
  notice,
}: {
  settings: AiSettings
  labels: AiProviderControlLabels
  controls: AiProviderControlsState
  notice?: string | null
}) {
  if (!controls.showCodexBanner && !notice) return null
  const errorCode = controls.account?.errorCode ?? controls.capabilities?.errorCode
  return (
    <div className="ai-codex-auth-banner" role="status">
      <span
        className={`ai-codex-auth-dot${errorCode || notice ? ' error' : ''}`}
        aria-hidden="true"
      />
      <div className="ai-codex-auth-copy">
        {settings.provider === 'openai-codex' && controls.account && (
          <>
            <strong>{labels.codexBrand}</strong>
            {controls.account.loggedIn === false && <span>{labels.signInRequired}</span>}
          </>
        )}
        {errorCode && <span className="ai-codex-auth-error">{labels.resolveError(errorCode)}</span>}
        {notice && <span className="ai-codex-auth-error">{notice}</span>}
      </div>
      {settings.provider === 'openai-codex' && controls.account?.loggedIn === false && (
        <button
          type="button"
          className="ai-codex-auth-login"
          disabled={controls.loginPending}
          aria-busy={controls.loginPending}
          onClick={() => void controls.startLogin()}
        >
          {labels.login}
        </button>
      )}
    </div>
  )
}

export function AiCodexModelControl({
  settings,
  labels,
  controls,
}: {
  settings: AiSettings
  labels: AiProviderControlLabels
  controls: AiProviderControlsState
}) {
  const model = controls.activeCodexModel
  if (settings.provider !== 'openai-codex' || !model) return null
  const reasoningLabel = labels.reasoningLabel(controls.activeCodexReasoning)
  const activePopover = controls.activePopover
  const items = [
    { id: 'model' as const, label: labels.model, value: model.name ?? model.id },
    { id: 'effort' as const, label: labels.reasoning, value: reasoningLabel },
    { id: 'speed' as const, label: labels.speed, value: controls.activeCodexSpeed.name },
  ]
  const submenuLabel =
    activePopover === 'model'
      ? labels.model
      : activePopover === 'effort'
        ? labels.reasoning
        : labels.speed

  return (
    <div className="ai-model-control" ref={controls.controlRef}>
      <button
        type="button"
        ref={controls.triggerRef}
        className={`ai-codex-model-trigger${activePopover ? ' expanded' : ''}`}
        onClick={controls.togglePicker}
        aria-label={`${labels.model}: ${model.name ?? model.id}; ${labels.reasoning}: ${reasoningLabel}; ${labels.speed}: ${controls.activeCodexSpeed.name}`}
        aria-expanded={Boolean(activePopover)}
        aria-haspopup="menu"
      >
        <span className="ai-codex-model-label">
          <span className="ai-codex-model-text">
            <span className="ai-codex-model-name">{model.name ?? model.id}</span>
            <span className="ai-codex-model-effort">{reasoningLabel}</span>
          </span>
        </span>
        <Caret />
      </button>
      {activePopover && (
        <div ref={controls.pickerRef} className="ai-codex-model-popover" data-codex-picker>
          <div
            ref={controls.rootRef}
            className="ai-codex-menu-root"
            role="menu"
            aria-label={`${labels.model}, ${labels.reasoning}, ${labels.speed}`}
          >
            {items.map((item) => (
              <button
                type="button"
                key={item.id}
                className={`ai-codex-menu-item${activePopover === item.id ? ' active' : ''}`}
                data-codex-menu-item={item.id}
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={activePopover === item.id}
                aria-controls="ai-codex-submenu"
                onClick={() => controls.setActivePopover(item.id)}
                onMouseEnter={() => controls.setActivePopover(item.id)}
                onFocus={() => controls.setActivePopover(item.id)}
              >
                <span>{item.label}</span>
                <span className="ai-codex-menu-value">{item.value}</span>
                <Caret up={false} />
              </button>
            ))}
            <div className="ai-popover-divider" />
            <button
              type="button"
              className="ai-codex-reset"
              role="menuitem"
              onClick={() => {
                controls.resetCodexSettings()
                controls.setActivePopover(null)
              }}
            >
              <span>{labels.reset}</span>
              <Reset />
            </button>
          </div>
          <div
            ref={controls.submenuRef}
            id="ai-codex-submenu"
            className="ai-codex-submenu"
            role="menu"
            aria-label={submenuLabel}
          >
            <div className="ai-popover-heading">{submenuLabel}</div>
            {activePopover === 'model' &&
              controls.capabilities?.models.map((candidate) => (
                <button
                  type="button"
                  key={candidate.id}
                  className={`ai-codex-option${candidate.id === model.id ? ' selected' : ''}`}
                  data-codex-model-option={candidate.id}
                  role="menuitemradio"
                  aria-checked={candidate.id === model.id}
                  onClick={() => {
                    controls.changeCodexModel(candidate.id)
                    controls.setActivePopover(null)
                  }}
                >
                  <span>{candidate.name ?? candidate.id}</span>
                  {candidate.id === model.id && <Check />}
                </button>
              ))}
            {activePopover === 'effort' &&
              model.reasoningEfforts.map((effort) => (
                <button
                  type="button"
                  key={effort}
                  className={`ai-codex-option${effort === controls.activeCodexReasoning ? ' selected' : ''}`}
                  data-codex-reasoning-option={effort}
                  role="menuitemradio"
                  aria-checked={effort === controls.activeCodexReasoning}
                  onClick={() => {
                    controls.changeCodexReasoning(effort)
                    controls.setActivePopover(null)
                  }}
                >
                  <span>{labels.reasoningLabel(effort)}</span>
                  {effort === controls.activeCodexReasoning && <Check />}
                </button>
              ))}
            {activePopover === 'speed' &&
              codexServiceTiers(model).map((tier) => (
                <button
                  type="button"
                  key={tier.id}
                  className={`ai-codex-option${tier.id === controls.activeCodexServiceTier ? ' selected' : ''}`}
                  data-codex-service-tier-option={tier.id}
                  role="menuitemradio"
                  aria-checked={tier.id === controls.activeCodexServiceTier}
                  onClick={() => {
                    controls.changeCodexServiceTier(tier.id)
                    controls.setActivePopover(null)
                  }}
                >
                  <span>
                    <span>{tier.name}</span>
                    {tier.description && (
                      <small className="ai-codex-option-description">{tier.description}</small>
                    )}
                  </span>
                  {tier.id === controls.activeCodexServiceTier && <Check />}
                </button>
              ))}
          </div>
        </div>
      )}
    </div>
  )
}
