/**
 * Parent-window bridge for embedding the web renderer in a business host such
 * as Dataflare Work. Authentication stays in the host/reverse proxy; this
 * channel carries only document context and UI commands.
 */
export const DATAFLARE_EMBED_PROTOCOL = 'genoffice-dataflare/v1'

export interface DataflareOfficeContext {
  tenantId?: string
  userId?: string
  documentId?: string
  documentType?: 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'markdown' | 'html'
  businessObject?: { type: string; id: string; label?: string }
  readonly?: boolean
  locale?: string
  theme?: 'light' | 'dark' | 'system'
}

export type DataflareEmbedCommand =
  | { type: 'init'; context: DataflareOfficeContext; sessionId: string }
  | { type: 'open-document'; documentId: string; documentType?: DataflareOfficeContext['documentType'] }
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

export type GenOfficeEmbedEvent =
  | { type: 'ready'; capabilities: string[] }
  | { type: 'document-dirty'; documentId?: string }
  | { type: 'document-saved'; documentId?: string; revision?: string }
  | { type: 'ai-progress'; requestId?: string; status: string; progress?: number }
  | { type: 'error'; code: string; message: string }

interface EmbedEnvelope<T> {
  protocol: typeof DATAFLARE_EMBED_PROTOCOL
  kind: 'command' | 'event'
  payload: T
}

function isEnvelope(value: unknown): value is EmbedEnvelope<DataflareEmbedCommand> {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<EmbedEnvelope<unknown>>
  return candidate.protocol === DATAFLARE_EMBED_PROTOCOL && candidate.kind === 'command'
}

function allowedParentOrigin(): string | null {
  if (document.referrer) {
    try {
      return new URL(document.referrer).origin
    } catch {
      // Fall through to the explicit development-only origin.
    }
  }
  const configured = new URLSearchParams(window.location.search).get('embedOrigin')
  if (!configured) return null
  try {
    return new URL(configured).origin
  } catch {
    return null
  }
}

export function postToEmbedParent(payload: GenOfficeEmbedEvent): void {
  if (window.parent === window) return
  const targetOrigin = allowedParentOrigin()
  if (!targetOrigin) return
  const message: EmbedEnvelope<GenOfficeEmbedEvent> = {
    protocol: DATAFLARE_EMBED_PROTOCOL,
    kind: 'event',
    payload,
  }
  window.parent.postMessage(message, targetOrigin)
}

export function installDataflareEmbedBridge(
  onCommand: (command: DataflareEmbedCommand) => void,
): () => void {
  if (window.parent === window) return () => {}
  const parentOrigin = allowedParentOrigin()
  if (!parentOrigin) return () => {}
  let sessionId: string | null = null
  const onMessage = (event: MessageEvent<unknown>) => {
    if (event.source !== window.parent) return
    if (event.origin !== parentOrigin) return
    if (!isEnvelope(event.data)) return
    const command = event.data.payload
    if (command.type === 'init') {
      if (!command.sessionId.trim()) return
      sessionId = command.sessionId
      onCommand(command)
      return
    }
    if (!sessionId) return
    onCommand(command)
  }
  window.addEventListener('message', onMessage)
  postToEmbedParent({
    type: 'ready',
    capabilities: ['document-context', 'ai-translation', 'ai-assistant', 'host-commands'],
  })
  return () => window.removeEventListener('message', onMessage)
}

export function isEmbeddedInHost(): boolean {
  return window.parent !== window
}
