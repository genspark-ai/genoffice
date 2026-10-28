import {
  createIpcTransport,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentStreamHandle,
  type AgentTransport,
} from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { t } from '../i18n/locale'

/** Wrap the AgentLoop's stream callbacks so reasoning deltas also reach the
 * app: the loop keeps its own copy (interleaved-thinking echo in the model
 * context) but AgentLoopEvents has no reasoning event of its own. */
export function withReasoningTap(
  callbacks: AgentStreamCallbacks,
  tap: (text: string) => void,
): AgentStreamCallbacks {
  return {
    ...callbacks,
    onReasoning: (text) => {
      tap(text)
      callbacks.onReasoning?.(text)
    },
  }
}

export interface ElectronTransportOptions {
  /** receives every reasoning delta of every turn, for the thinking block */
  onReasoning?: (text: string) => void
}

/** The shared IPC transport wired to the sheets preload bridge (window.desktopApi). */
export function createElectronTransport(
  getSettings: () => AiSettings,
  options: ElectronTransportOptions = {},
): AgentTransport {
  const base = createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktopApi.onAiStream(listener),
    start: (request) => window.desktopApi.aiStream(request),
    cancel: (requestId) => void window.desktopApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
    regionErrorText: () => t('aiRegionError'),
  })
  const tap = options.onReasoning
  if (!tap) return base
  return {
    stream: (request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle =>
      base.stream(request, withReasoningTap(callbacks, tap)),
  }
}
