import { createIpcTransport, type AgentTransport } from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { resolveCodexError } from '@genoffice/ai-provider'
import { getLang } from '../i18n/locale'

/** The shared IPC transport wired to the slides preload bridge (window.slidesApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.slidesApi.onAiStream(listener),
    start: (request) => window.slidesApi.aiStream(request),
    cancel: (requestId) => void window.slidesApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => resolveCodexError(undefined, getLang()),
    timeoutErrorText: () => resolveCodexError('timeout', getLang()),
    creditsErrorText: () => resolveCodexError('credits', getLang()),
    resolveErrorCode: (code) => resolveCodexError(code, getLang()),
  })
}
