import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import { resolveCodexError, type AiSettings } from '@genoffice/ai-provider'
import { getLang } from '../i18n/locale'

/** The shared IPC transport wired to the markdown preload bridge (window.markdownApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.markdownApi.onAiStream(listener),
    start: (request) => window.markdownApi.aiStream(request),
    cancel: (requestId) => void window.markdownApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => resolveCodexError(undefined, getLang()),
    timeoutErrorText: () => resolveCodexError('timeout', getLang()),
    creditsErrorText: () => resolveCodexError('credits', getLang()),
    resolveErrorCode: (code) => resolveCodexError(code, getLang()),
  })
}
