import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the slides preload bridge (window.slidesApi). */
export function createElectronTransport(): AgentTransport {
  return createIpcTransport({
    onStream: (listener) => window.slidesApi.onAiStream(listener),
    start: (request) => window.slidesApi.aiStream(request),
    cancel: (requestId) => void window.slidesApi.aiStreamCancel(requestId),
    task: 'chat',
    unknownErrorText: () => t('aiErrUnknown'),
    timeoutErrorText: () => t('aiErrStreamTimeout'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
