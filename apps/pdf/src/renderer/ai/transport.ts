import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { t } from '../i18n/locale'

/**
 * The shared IPC transport wired to the pdf preload bridge (window.pdfApi).
 * `onReasoning` tees reasoning deltas to the chat UI while the loop keeps its
 * own copy (interleaved-thinking echo); requestId filtering stays inside the
 * base transport.
 */
export function createElectronTransport(
  getSettings: () => AiSettings,
  onReasoning?: (text: string) => void,
): AgentTransport {
  const base = createIpcTransport<AiSettings>({
    onStream: (listener) => window.pdfApi.onAiStream(listener),
    start: (request) => window.pdfApi.aiStream(request),
    cancel: (requestId) => void window.pdfApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
    regionErrorText: () => t('aiRegionError'),
  })
  return {
    stream: (request, callbacks) =>
      base.stream(request, {
        ...callbacks,
        onReasoning: (text) => {
          onReasoning?.(text)
          callbacks.onReasoning?.(text)
        },
      }),
  }
}
