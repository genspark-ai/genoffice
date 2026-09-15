export const DATAFLARE_EMBED_PROTOCOL = 'genoffice-dataflare/v1'

export interface DataflareOfficeContext {
  tenantId?: string
  userId?: string
  documentId?: string
  documentType?: 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'markdown' | 'html'
  documentSource?: 'knowledge' | 'office'
  businessObject?: { type: string; id: string; label?: string }
  readonly?: boolean
  locale?: string
  theme?: 'light' | 'dark' | 'system'
}

export type DataflareGlobalState = Partial<Pick<DataflareOfficeContext,
  | 'tenantId'
  | 'userId'
  | 'locale'
  | 'theme'
  | 'readonly'
>> & { documentRevision?: string | number }

export type DataflareEmbedCommand =
  | { type: 'init'; context: DataflareOfficeContext; sessionId: string }
  | { type: 'set-readonly'; readonly: boolean }
  | { type: 'focus-ai'; prompt?: string }
  | {
      type: 'translate'
      scope: 'selection' | 'document'
      sourceLanguage?: string
      targetLanguage: string
      preserveFormatting?: boolean
    }
  | { type: 'save' }
  | { type: 'dispose' }
  | { type: 'global-state-update'; state: DataflareGlobalState; revision?: number }

export type GenOfficeEmbedEvent =
  | { type: 'ready'; capabilities: string[] }
  | { type: 'document-dirty'; documentId?: string }
  | { type: 'document-saved'; documentId?: string; revision?: string }
  | { type: 'ai-progress'; requestId?: string; status: string; progress?: number }
  | { type: 'error'; code: string; message: string }
  | { type: 'global-state'; state: DataflareGlobalState; revision?: number }
  | { type: 'global-state-request'; revision?: number }

export type DataflareParentRequest = {
  type: 'http-request'
  requestId: string
  sessionId: string
  method: 'GET' | 'POST'
  path: string
  jsonBody?: string
  file?: { bytes: ArrayBuffer; filename: string; contentType: string }
  fields?: Record<string, string>
}

export type DataflareParentResponse = {
  type: 'http-response'
  requestId: string
  sessionId: string
  status: number
  headers: Record<string, string>
  body: ArrayBuffer
}

export type DataflareParentStreamRequest = {
  type: 'http-stream-request'
  requestId: string
  sessionId: string
  method: 'POST'
  path: string
  jsonBody?: string
}

export type DataflareParentStreamEvent = {
  type: 'http-stream-event'
  requestId: string
  sessionId: string
  eventName?: string
  eventId?: string
  data: string
}

export type DataflareParentStreamClose = {
  type: 'http-stream-close'
  requestId: string
  sessionId: string
  status: number
}

interface EmbedEnvelope<T> {
  protocol: typeof DATAFLARE_EMBED_PROTOCOL
  kind: 'command' | 'event' | 'request' | 'response' | 'stream-request' | 'stream-event' | 'stream-close'
  sessionId?: string
  payload: T
}

let activeSessionId: string | null = null

export function getDataflareEmbedSessionId(): string | null {
  return activeSessionId
}

function isCommandEnvelope(value: unknown): value is EmbedEnvelope<DataflareEmbedCommand> {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<EmbedEnvelope<unknown>>
  return candidate.protocol === DATAFLARE_EMBED_PROTOCOL && candidate.kind === 'command'
}

export function isStreamEventEnvelope(
  value: unknown,
): value is EmbedEnvelope<DataflareParentStreamEvent | DataflareParentStreamClose> {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<EmbedEnvelope<unknown>>
  return (
    candidate.protocol === DATAFLARE_EMBED_PROTOCOL &&
    (candidate.kind === 'stream-event' || candidate.kind === 'stream-close')
  )
}

function parentOrigin(): string | null {
  if (document.referrer) {
    try {
      return new URL(document.referrer).origin
    } catch {}
  }
  const ancestorOrigin = window.location.ancestorOrigins?.[0]
  return ancestorOrigin || null
}

export function isEmbeddedInHost(): boolean {
  return window.parent !== window
}

export function postToEmbedParent(payload: GenOfficeEmbedEvent): void {
  const origin = parentOrigin()
  if (!isEmbeddedInHost() || !origin) return
  const message: EmbedEnvelope<GenOfficeEmbedEvent> = {
    protocol: DATAFLARE_EMBED_PROTOCOL,
    kind: 'event',
    ...(activeSessionId ? { sessionId: activeSessionId } : {}),
    payload,
  }
  window.parent.postMessage(message, origin)
}

export function requestDataflareParent(request: DataflareParentRequest): Promise<DataflareParentResponse> {
  const origin = parentOrigin()
  if (!isEmbeddedInHost() || !origin || !activeSessionId || request.sessionId !== activeSessionId) {
    return Promise.reject(new Error('Dataflare parent bridge is unavailable'))
  }
  return new Promise((resolve, reject) => {
    const timeout = window.setTimeout(() => {
      window.removeEventListener('message', onMessage)
      reject(new Error('Dataflare parent request timed out'))
    }, 60_000)
    const onMessage = (event: MessageEvent<unknown>) => {
      if (event.source !== window.parent || event.origin !== origin) return
      const envelope = event.data as Partial<EmbedEnvelope<DataflareParentResponse>>
      const response = envelope.payload
      if (envelope.protocol !== DATAFLARE_EMBED_PROTOCOL || envelope.kind !== 'response' ||
        envelope.sessionId !== request.sessionId || !response || response.type !== 'http-response' ||
        response.sessionId !== request.sessionId || response.requestId !== request.requestId) return
      window.clearTimeout(timeout)
      window.removeEventListener('message', onMessage)
      resolve(response)
    }
    window.addEventListener('message', onMessage)
    const transfer = request.file ? [request.file.bytes] : []
    window.parent.postMessage({
      protocol: DATAFLARE_EMBED_PROTOCOL,
      kind: 'request',
      sessionId: request.sessionId,
      payload: request,
    }, origin, transfer)
  })
}

/**
 * Subscribe to an SSE stream proxied through the Dataflare host page.
 * Returns an unsubscribe function. The host opens the SSE connection
 * to Dataflare and forwards each event as a postMessage envelope.
 */
export function requestDataflareStreamParent(
  request: DataflareParentStreamRequest,
  onEvent: (event: DataflareParentStreamEvent) => void,
  onClose: (status: number) => void,
  onError: (error: Error) => void,
): () => void {
  const origin = parentOrigin()
  if (!isEmbeddedInHost() || !origin || !activeSessionId || request.sessionId !== activeSessionId) {
    onError(new Error('Dataflare stream bridge is unavailable'))
    return () => {}
  }
  const onMessage = (event: MessageEvent<unknown>) => {
    if (event.source !== window.parent || event.origin !== origin) return
    if (!isStreamEventEnvelope(event.data)) return
    const envelope = event.data as EmbedEnvelope<DataflareParentStreamEvent | DataflareParentStreamClose>
    if (envelope.sessionId !== request.sessionId) return
    const payload = envelope.payload
    if (payload.requestId !== request.requestId) return
    if (envelope.kind === 'stream-event') {
      onEvent(payload as DataflareParentStreamEvent)
    } else if (envelope.kind === 'stream-close') {
      window.removeEventListener('message', onMessage)
      onClose((payload as DataflareParentStreamClose).status)
    }
  }
  window.addEventListener('message', onMessage)
  window.parent.postMessage({
    protocol: DATAFLARE_EMBED_PROTOCOL,
    kind: 'stream-request',
    sessionId: request.sessionId,
    payload: request,
  }, origin)
  return () => {
    window.removeEventListener('message', onMessage)
  }
}

export interface DataflareEmbedBridgeHandlers {
  onCommand: (command: DataflareEmbedCommand) => void
  onGlobalState?: (state: DataflareGlobalState, revision: number | undefined) => void
}

export function installDataflareEmbedBridge(
  onCommandOrHandlers: ((command: DataflareEmbedCommand) => void) | DataflareEmbedBridgeHandlers,
): () => void {
  const handlers: DataflareEmbedBridgeHandlers =
    typeof onCommandOrHandlers === 'function'
      ? { onCommand: onCommandOrHandlers }
      : onCommandOrHandlers
  const origin = parentOrigin()
  if (!isEmbeddedInHost() || !origin) return () => {}
  let sessionId: string | null = null
  const onMessage = (event: MessageEvent<unknown>) => {
    if (event.source !== window.parent || event.origin !== origin) return
    if (!isCommandEnvelope(event.data)) return
    const command = event.data.payload
    if (command.type === 'init') {
      const normalizedSessionId = command.sessionId.trim()
      if (!normalizedSessionId) return
      activeSessionId = normalizedSessionId
      sessionId = normalizedSessionId
    } else if (!sessionId || (event.data as Partial<EmbedEnvelope<unknown>>).sessionId !== sessionId) {
      return
    }
    if (command.type === 'global-state-update') {
      handlers.onGlobalState?.(command.state, command.revision)
      return
    }
    handlers.onCommand(command)
  }
  window.addEventListener('message', onMessage)
  postToEmbedParent({
    type: 'ready',
    capabilities: ['document-context', 'ai-translation', 'ai-assistant', 'host-commands', 'global-state'],
  })
  postToEmbedParent({ type: 'global-state-request' })
  return () => {
    window.removeEventListener('message', onMessage)
    activeSessionId = null
  }
}
