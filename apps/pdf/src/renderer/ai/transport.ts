import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the pdf preload bridge (window.pdfApi). */
export function createElectronTransport(): AgentTransport {
  return createIpcTransport({
    onStream: (listener) => window.pdfApi.onAiStream(listener),
    start: (request) => window.pdfApi.aiStream(request),
    cancel: (requestId) => void window.pdfApi.aiStreamCancel(requestId),
    task: 'chat',
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
  })
}
