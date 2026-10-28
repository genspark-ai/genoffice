import {
  createIpcTransport,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * The shared IPC transport wired to the docs preload bridge (window.desktop).
 * `onReasoning`, when given, receives the model's reasoning deltas in addition
 * to the stream callbacks' own consumer (the agent loop stores them for
 * interleaved-thinking echo but does not surface them through its events).
 */
export function createElectronTransport(
  getSettings: () => AiSettings,
  onReasoning?: (text: string) => void,
): AgentTransport {
  const base = createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktop.onAiStream(listener),
    start: (request) => window.desktop.aiStream(request),
    cancel: (requestId) => void window.desktop.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
    regionErrorText: () => t('aiRegionError'),
  })
  if (!onReasoning) return base
  return {
    stream: (request: AgentStreamRequest, cb: AgentStreamCallbacks) =>
      base.stream(request, {
        ...cb,
        onReasoning: cb.onReasoning
          ? (text: string) => {
              cb.onReasoning?.(text)
              onReasoning(text)
            }
          : undefined,
      }),
  }
}
