import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { t } from '../i18n/locale'

/** The shared IPC transport wired to the html preload bridge (window.htmlApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.htmlApi.onAiStream(listener),
    start: (request) => window.htmlApi.aiStream(request),
    cancel: (requestId) => void window.htmlApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
  })
}
