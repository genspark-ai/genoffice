import {
  createIpcTransport,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentStreamHandle,
  type AgentTransport,
} from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { t } from '../i18n/locale'

/**
 * The shared IPC transport wired to the html preload bridge (window.htmlApi).
 * `onReasoning` taps the model's reasoning deltas (the loop keeps its own
 * callback: it stores the text on the assistant message and enforces the
 * per-turn budget) so the panel can render the thinking block live.
 */
export function createElectronTransport(
  getSettings: () => AiSettings,
  onReasoning?: (text: string) => void,
): AgentTransport {
  const transport = createIpcTransport<AiSettings>({
    onStream: (listener) => window.htmlApi.onAiStream(listener),
    start: (request) => window.htmlApi.aiStream(request),
    cancel: (requestId) => void window.htmlApi.aiStreamCancel(requestId),
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
    stream: (request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle =>
      transport.stream(request, {
        ...callbacks,
        onReasoning: (text) => {
          callbacks.onReasoning?.(text)
          onReasoning(text)
        },
      }),
  }
}
