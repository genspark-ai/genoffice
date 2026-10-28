import {
  createIpcTransport,
  type AgentStreamCallbacks,
  type AgentStreamHandle,
  type AgentStreamRequest,
  type AgentTransport,
} from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { t } from '../i18n/locale'

/** Wrap the AgentLoop's stream callbacks so reasoning deltas also reach the
 * app: the loop keeps its own copy (interleaved-thinking echo in the model
 * context) but AgentLoopEvents has no reasoning event of its own. */
function withReasoningTap(
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

/**
 * The shared IPC transport wired to the markdown preload bridge
 * (window.markdownApi). `onReasoning` receives every reasoning delta of every
 * turn, for the panel's thinking block.
 */
export function createElectronTransport(
  getSettings: () => AiSettings,
  onReasoning?: (text: string) => void,
): AgentTransport {
  const transport = createIpcTransport<AiSettings>({
    onStream: (listener) => window.markdownApi.onAiStream(listener),
    start: (request) => window.markdownApi.aiStream(request),
    cancel: (requestId) => void window.markdownApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
    regionErrorText: () => t('aiRegionError'),
  })
  if (!onReasoning) return transport
  return {
    ...transport,
    stream: (request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle =>
      transport.stream(request, withReasoningTap(callbacks, onReasoning)),
  }
}
