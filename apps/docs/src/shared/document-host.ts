import type { DocsSaveOutcome } from './operation-result'

export type DocsDocumentOperation = 'new' | 'open' | 'save' | 'saveAs' | 'export' | 'create'
export interface DocsHostedIdentity {
  id: string
  name: string
  /** Real desktop working file, when this host is file-backed. */
  path?: string
  fileHash?: string
  readOnly?: boolean
}
export interface DocsHostedDocument extends DocsHostedIdentity {
  /** Omit for a new blank document. Opened content is complete DOCX bytes. */
  data?: ArrayBuffer
}
export interface DocsHostedSaveRequest {
  document: DocsHostedIdentity
  data: ArrayBuffer
  kind: 'save' | 'saveAs'
  auto: boolean
  /** Version loaded by the editor; file-backed hosts reject an external change. */
  expectedHash?: string
  /** A caller proposal, never authority to write a local file. */
  requestedPath?: string
  overwrite?: boolean
}
export interface DocsHostedSaveOutcome extends DocsSaveOutcome {
  document?: DocsHostedIdentity
  /** Required for Save As. true also requires document; false preserves current identity and dirty state. */
  adopt?: boolean
}
export interface DocsDocumentHost {
  capabilities?: Partial<Record<DocsDocumentOperation, boolean>>
  autoSave?: {
    /** Initial editor-session state; default false. Users can toggle it unless explicitly locked. */
    enabled: boolean
    allowToggle?: boolean
    intervalMs?: number
    onBlur?: boolean
  }
  initialDocument?(): Promise<DocsHostedDocument | null>
  newDocument?(): Promise<DocsHostedDocument | null>
  openDocument?(id?: string): Promise<DocsHostedDocument | null>
  recentDocuments?(): Promise<DocsHostedIdentity[]>
  saveDocument?(request: DocsHostedSaveRequest): Promise<DocsHostedSaveOutcome>
  createDocument?(request: { type: string; title: string; content: string }): Promise<{
    ok: boolean
    committed?: boolean | null
    pending?: boolean
    error?: string
    opened?: boolean
  }>
}

export interface DocsDocumentCommands {
  new: () => Promise<boolean>
  open(id?: string): Promise<boolean>
  save(saveAs?: boolean): Promise<boolean>
  state(): DocsDocumentState | null
  /** Immediately reports current state, then metadata/dirty/session changes. Dispose to unsubscribe. */
  subscribe(listener: (state: DocsDocumentState | null) => void): () => void
}

export interface DocsDocumentState extends Omit<DocsHostedIdentity, 'path'> {
  /** Opaque editor session identity; changes on reopening or adopting another document. */
  documentToken: string
  path: string | null
  dirty: boolean
  readOnly: boolean
}

export function assertDocumentOperation(
  host: DocsDocumentHost,
  operation: DocsDocumentOperation,
): void {
  if (host.capabilities?.[operation] === false)
    throw new Error(`The host disabled the ${operation} document operation`)
}

export function validateHostedIdentity(document: DocsHostedIdentity): void {
  if (
    !document ||
    typeof document.id !== 'string' ||
    !document.id ||
    typeof document.name !== 'string' ||
    !document.name ||
    (document.path !== undefined && (typeof document.path !== 'string' || !document.path)) ||
    (document.fileHash !== undefined && typeof document.fileHash !== 'string') ||
    (document.readOnly !== undefined && typeof document.readOnly !== 'boolean')
  )
    throw new Error('The host returned an invalid document identity')
}
