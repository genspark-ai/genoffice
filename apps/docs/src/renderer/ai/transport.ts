import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the docs preload bridge (window.desktop). */
export function createElectronTransport(): AgentTransport {
  return createIpcTransport({
    onStream: (listener) => window.desktop.onAiStream(listener),
    start: (request) => window.desktop.aiStream(request),
    cancel: (requestId) => void window.desktop.aiStreamCancel(requestId),
    task: 'chat',
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
