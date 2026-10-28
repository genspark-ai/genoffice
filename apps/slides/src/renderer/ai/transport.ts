import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the slides preload bridge (window.slidesApi).
 *  `onReasoning` taps model reasoning deltas for the panel's thinking block
 *  (AgentLoop keeps them for model context; its event surface has no
 *  onReasoning, so the app reads them here). */
export function createElectronTransport(
  getSettings: () => AiSettings,
  onReasoning?: (text: string) => void,
): AgentTransport {
  const transport = createIpcTransport<AiSettings>({
    onStream: (listener) => window.slidesApi.onAiStream(listener),
    start: (request) => window.slidesApi.aiStream(request),
    cancel: (requestId) => void window.slidesApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiErrUnknown'),
    timeoutErrorText: () => t('aiErrStreamTimeout'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiErrNetwork'),
    overloadedErrorText: () => t('aiErrOverloaded'),
    regionErrorText: () => t('aiErrRegion'),
  })
  if (!onReasoning) return transport
  return {
    ...transport,
    stream: (request, callbacks) =>
      transport.stream(request, {
        ...callbacks,
        onReasoning: (text) => {
          callbacks.onReasoning?.(text)
          onReasoning(text)
        },
      }),
  }
}
