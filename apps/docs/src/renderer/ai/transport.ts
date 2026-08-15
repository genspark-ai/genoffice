import { createIpcTransport, type AgentTransport, type IpcErrorCode } from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t, type TFunc } from '../i18n/locale'

export function codexErrorText(code: IpcErrorCode | undefined, translate: TFunc = t): string {
  switch (code) {
    case 'auth-required':
      return translate('aiCodexSignInRequired')
    case 'auth-expired':
      return translate('aiCodexAuthExpired')
    case 'auth-temporary':
      return translate('aiCodexAuthTemporary')
    case 'timeout':
      return translate('aiTimeoutError')
    case 'capabilities-unavailable':
      return translate('aiCodexModelsUnavailable')
    case 'rate-limit':
      return translate('aiCodexRateLimited')
    case 'request-rejected':
      return translate('aiCodexRequestRejected')
    case 'invalid-stream':
      return translate('aiCodexInvalidStream')
    case 'invalid-tool-call':
      return translate('aiCodexInvalidToolCall')
    case 'credits':
      return translate('aiCreditsExhausted')
    default:
      return translate('aiUnknownError')
  }
}

/** The shared IPC transport wired to the docs preload bridge (window.desktop). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktop.onAiStream(listener),
    start: (request) => window.desktop.aiStream(request),
    cancel: (requestId) => void window.desktop.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    resolveErrorCode: (code) => codexErrorText(code),
  })
}
