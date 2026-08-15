import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { fetchCodexCapabilities } from './codex'
import type { CodexAccountStatus } from './auth'
import { CodexError } from './types'
import type {
  AiProviderConfig,
  AiProviderId,
  AiSettings,
  AiStreamChunk,
  CodexAuthContext,
  CodexCapabilities,
  CodexErrorCode,
  CodexReasoningEffort,
  LegacyAiSettings,
} from './types'
import { AiCreditsError, streamForProvider } from './stream'
import { AiTimeoutError } from './watchdog'
import { AI_PROVIDERS, defaultAiSettings, resolveAiSettings } from './providers'

export interface AiRuntimeAuth {
  status(): Promise<CodexAccountStatus>
  login(): Promise<CodexAccountStatus>
  cancelLogin(): void
  logout(): Promise<void>
  getContext(): Promise<CodexAuthContext>
}

export interface AiRuntimeMessages {
  noGensparkApiKey: string
  noApiKey(provider: AiProviderId): string
  noModel: string
}

export interface AiMainRuntimeOptions {
  auth: AiRuntimeAuth
  getGensparkApiKey(): string
  fetchCapabilities?: (auth: CodexAuthContext) => Promise<CodexCapabilities>
  streamProvider?: typeof streamForProvider
}

export interface AiRuntimeRequest {
  requestId: string
  settings: AiSettings
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

export type AiRuntimeSender = (chunk: AiStreamChunk) => void

const CODEX_REASONING_EFFORTS: readonly CodexReasoningEffort[] = [
  'none',
  'minimal',
  'low',
  'medium',
  'high',
  'xhigh',
  'max',
  'ultra',
]

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isReasoningEffort(value: unknown): value is CodexReasoningEffort {
  return (
    typeof value === 'string' && CODEX_REASONING_EFFORTS.includes(value as CodexReasoningEffort)
  )
}

function isServiceTier(value: unknown): value is string {
  return typeof value === 'string' && /^[a-zA-Z0-9._-]{1,64}$/.test(value)
}

function providerConfig(value: unknown): Record<string, unknown> {
  return isRecord(value) ? value : {}
}

/** Keep only the two account-backed providers exposed by the sidebar. */
export function sanitizeAiSettings(value: unknown): AiSettings {
  const raw = isRecord(value) ? value : {}
  const stored = resolveAiSettings(
    raw as Partial<AiSettings> & LegacyAiSettings,
    defaultAiSettings(),
  )
  const provider: AiProviderId = stored.provider === 'openai-codex' ? 'openai-codex' : 'genspark'
  const providers = {} as AiSettings['providers']
  for (const meta of AI_PROVIDERS) {
    const rawConfig = providerConfig(stored.providers[meta.id])
    providers[meta.id] = {
      // Provider keys are runtime credentials, never renderer-visible settings.
      apiKey: '',
      model:
        typeof rawConfig.model === 'string' && rawConfig.model
          ? rawConfig.model
          : meta.defaultModel,
      ...(meta.needsBaseUrl
        ? { baseUrl: typeof rawConfig.baseUrl === 'string' ? rawConfig.baseUrl : '' }
        : {}),
      ...(meta.id === 'openai-codex'
        ? {
            reasoningEffort: isReasoningEffort(rawConfig.reasoningEffort)
              ? rawConfig.reasoningEffort
              : 'none',
            serviceTier: isServiceTier(rawConfig.serviceTier) ? rawConfig.serviceTier : 'default',
          }
        : {}),
    }
  }
  return { provider, providers }
}

function codexErrorCode(error: unknown): CodexErrorCode {
  return error instanceof CodexError ? error.code : 'provider-failure'
}

export class AiMainRuntime {
  private readonly streams = new Map<string, AbortController>()
  private readonly auth: AiRuntimeAuth
  private readonly getGensparkApiKey: () => string
  private readonly fetchCapabilities: (auth: CodexAuthContext) => Promise<CodexCapabilities>
  private readonly streamProvider: typeof streamForProvider

  constructor(options: AiMainRuntimeOptions) {
    this.auth = options.auth
    this.getGensparkApiKey = options.getGensparkApiKey
    this.fetchCapabilities = options.fetchCapabilities ?? fetchCodexCapabilities
    this.streamProvider = options.streamProvider ?? streamForProvider
  }

  async codexStatus(): Promise<CodexAccountStatus> {
    try {
      return await this.auth.status()
    } catch (error) {
      return { loggedIn: false, errorCode: codexErrorCode(error) }
    }
  }

  async codexLogin(): Promise<CodexAccountStatus> {
    try {
      return await this.auth.login()
    } catch (error) {
      return { loggedIn: false, errorCode: codexErrorCode(error) }
    }
  }

  codexCancelLogin(): void {
    this.auth.cancelLogin()
  }

  async codexLogout(): Promise<CodexAccountStatus> {
    try {
      await this.auth.logout()
      return { loggedIn: false }
    } catch (error) {
      return { loggedIn: false, errorCode: codexErrorCode(error) }
    }
  }

  async codexCapabilities(): Promise<CodexCapabilities & { errorCode?: CodexErrorCode }> {
    try {
      return await this.fetchCapabilities(await this.auth.getContext())
    } catch (error) {
      return { models: [], errorCode: codexErrorCode(error) }
    }
  }

  async stream(
    request: AiRuntimeRequest,
    send: AiRuntimeSender,
    messages: AiRuntimeMessages,
  ): Promise<void> {
    const settings = sanitizeAiSettings(request.settings)
    const provider = settings.provider
    const config = { ...settings.providers[provider] }
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    let authContext: CodexAuthContext | undefined

    if (provider === 'openai-codex') {
      try {
        authContext = await this.auth.getContext()
        const capabilities = await this.fetchCapabilities(authContext)
        const model = capabilities.models.find((candidate) => candidate.id === config.model)
        if (!model) {
          send({
            requestId: request.requestId,
            type: 'error',
            errorCode: 'capabilities-unavailable',
          })
          return
        }
        if (
          config.reasoningEffort &&
          config.reasoningEffort !== 'none' &&
          !model.reasoningEfforts.includes(config.reasoningEffort)
        ) {
          send({
            requestId: request.requestId,
            type: 'error',
            errorCode: 'capabilities-unavailable',
          })
          return
        }
        if (
          config.serviceTier &&
          config.serviceTier !== 'default' &&
          !model.serviceTiers?.some((tier) => tier.id === config.serviceTier)
        ) {
          send({
            requestId: request.requestId,
            type: 'error',
            errorCode: 'capabilities-unavailable',
          })
          return
        }
      } catch (error) {
        send({ requestId: request.requestId, type: 'error', errorCode: codexErrorCode(error) })
        return
      }
    } else {
      if (!config.apiKey && provider === 'genspark') config.apiKey = this.getGensparkApiKey()
      if (!config.apiKey) {
        send({
          requestId: request.requestId,
          type: 'error',
          error: provider === 'genspark' ? messages.noGensparkApiKey : messages.noApiKey(provider),
        })
        return
      }
    }
    if (!config.model) {
      send({ requestId: request.requestId, type: 'error', error: messages.noModel })
      return
    }

    const controller = new AbortController()
    this.streams.set(request.requestId, controller)
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId: request.requestId, type: 'ping' })
    }
    try {
      let stopReason: string | undefined
      await this.streamProvider(
        provider,
        config as AiProviderConfig,
        request.system,
        request.messages,
        tools,
        maxTokens,
        {
          signal: controller.signal,
          onDelta: (text) => send({ requestId: request.requestId, type: 'delta', text }),
          onToolCall: (toolCall: AgentToolCall) =>
            send({ requestId: request.requestId, type: 'tool-call', toolCall }),
          onActivity: ping,
          onStopReason: (reason) => {
            stopReason = reason
          },
        },
        authContext,
      )
      send({ requestId: request.requestId, type: 'done', ...(stopReason ? { stopReason } : {}) })
    } catch (error) {
      if (controller.signal.aborted) {
        send({ requestId: request.requestId, type: 'done' })
      } else if (provider === 'openai-codex') {
        send({
          requestId: request.requestId,
          type: 'error',
          errorCode: error instanceof AiTimeoutError ? 'timeout' : codexErrorCode(error),
        })
      } else {
        send({
          requestId: request.requestId,
          type: 'error',
          error: error instanceof Error ? error.message : String(error),
          ...(error instanceof AiTimeoutError
            ? { errorCode: 'timeout' as const }
            : error instanceof AiCreditsError
              ? { errorCode: 'credits' as const }
              : {}),
        })
      }
    } finally {
      if (this.streams.get(request.requestId) === controller) this.streams.delete(request.requestId)
    }
  }

  cancel(requestId: string): void {
    this.streams.get(requestId)?.abort()
  }
}
