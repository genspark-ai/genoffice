import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import { resolveCodexError, type AiSettings } from '@genoffice/ai-provider'
import { getLang, t } from '../i18n/locale'

/** The shared IPC transport wired to the sheets preload bridge (window.desktopApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktopApi.onAiStream(listener),
    start: (request) => window.desktopApi.aiStream(request),
    cancel: (requestId) => void window.desktopApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => resolveCodexError(undefined, getLang()),
    timeoutErrorText: () => resolveCodexError('timeout', getLang()),
    creditsErrorText: () => resolveCodexError('credits', getLang()),
    resolveErrorCode: (code) =>
      code === 'network' ? undefined : resolveCodexError(code, getLang()),
    networkErrorText: () => t('aiNetworkError'),
  })
}
