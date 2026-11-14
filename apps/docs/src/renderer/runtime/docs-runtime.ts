// Modified for ForaOffice: reliable document targeting and portable inspection.
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import type { AgentToolCall, ToolExecution } from '@genoffice/agent-core'
import { ResourceAccessError } from '@genoffice/agent-core/resource-tools'
import {
  buildDocContext,
  getSelectionScope,
  HTML_RULES,
  parseHtmlFragment,
  parseInlineFragment,
  resolveImageRefs,
  type AiTrack,
  type NumIds,
} from '../ai/protocol'
import { executeOps, opCatalog, opNames } from '../ai/ops'
import type { DocsEngineTarget } from '../../shared/document-query'
import {
  executeTool,
  markDocSeen,
  onDocSeenCarry,
  validateHtmlStructureIntent,
  type AiCommentsAccess,
  type AiDocExtras,
  type AiHeaderFooterAccess,
  type FrozenSelection,
} from '../ai/tools'
import type { AiNotesAccess } from '../ai/note-ops'
import type { AiPageSetupAccess } from '../ai/page-setup'
import type { WritePosition } from './draft-landing'
import { z } from 'zod'
import { docsToolCatalog, DOCS_READ_TOOLS } from '../../shared/tool-catalog'
import { DOCS_DRAFT_CONTROLS, DOCS_RUNTIME_TOOLS } from '../../shared/runtime-tools'
import { DocumentDraft, DraftInputError, draftPosition } from './document-draft'
import { DRAFT_META, hasActiveDocumentDraft } from './draft-landing'
import { toolUnavailable } from './capabilities'
import type {
  DocsErrorCode,
  DocsOperationResult,
  DocsSaveOutcome,
  DocsSaveResult,
} from '../../shared/operation-result'
import {
  readDocumentMedia,
  documentMediaReferences,
  prepareDocumentImageWrite,
  bindDocumentImageHtml,
} from './document-media'
import {
  DOCS_RESOURCE_TOOL_NAMES,
  DOCS_RESOURCE_TOOLS,
  DOCS_FILE_COMMANDS,
  type DocsResourceTools,
} from '../../shared/resource-tools'
import { redactLabelsOf } from '../ai/redact-view'
import { queryDocument, projectedTargetHtml, type QueryContinuation } from './document-query'

type RuntimeExecution = ToolExecution &
  Pick<DocsOperationResult, 'errorCode' | 'media' | 'imageReferenceErrors'> &
  Partial<Pick<DocsOperationResult, 'committed' | 'receipt' | 'matchCount' | 'changeCount'>>

interface DocumentGenerationState {
  documentId: string
  revision: number
  queue: Promise<unknown>
}

export interface RuntimeDocument {
  editor: Editor
  /** Changes when the host opens or replaces a document, not on each render. */
  identity: object
  /** Stable while authored stores outside the body are unchanged. */
  stateVersion?: unknown
  numIds: NumIds
  track?: AiTrack
  /** Match the host's configured review policy; detached consumers retain AI highlights by default. */
  autoAcceptUntracked?: boolean
  comments?: AiCommentsAccess
  /** Comments-only document protection permits review without body edits. */
  commentsAllowed?: boolean
  headerFooter?: AiHeaderFooterAccess
  pageSetup?: AiPageSetupAccess
  extras?: AiDocExtras
  notes?: AiNotesAccess
}

/** Detached runtimes can observe these complete authored stores without a host epoch. */
function stateVersionOf(doc: RuntimeDocument | null | undefined): unknown {
  if (!doc || doc.stateVersion !== undefined) return doc?.stateVersion
  return JSON.stringify({
    headerFooter: doc.headerFooter?.read(),
    pages: doc.pageSetup?.list(),
    comments: doc.comments?.list(),
    footnotes: doc.notes?.list('footnote'),
    endnotes: doc.notes?.list('endnote'),
  })
}

function authoredStoreUnavailable(name: string, doc: RuntimeDocument | null) {
  if (doc?.stateVersion === undefined && ['define_style', 'set_watermark'].includes(name))
    return {
      reason:
        'This authored store has no complete observable snapshot. The host must provide RuntimeDocument.stateVersion for style or watermark writes.',
      errorCode: 'unavailable' as const,
    }
  return undefined
}

export interface RuntimeCaller {
  id: string
  /** Keep the provider instance stable until its authority is replaced or revoked. */
  resources?: () => DocsResourceTools | undefined
}

export interface DocsRuntimeBindings {
  current(): RuntimeDocument | null
  /** Commit synchronous document/store updates before the next tool or save observes them. */
  apply?<T>(operation: () => T): T
  save?(options: { path: string; overwrite: boolean }): Promise<DocsSaveOutcome>
  /** Live authorized provider; changing the instance revokes pending source reads. */
  resources?(): DocsResourceTools | undefined
}

export interface DocsRuntimeClient {
  readonly signal: AbortSignal
  context(options?: { captureSelection?: boolean }): {
    documentId: string
    revision: number
    text: string
  }
  execute(call: AgentToolCall, signal?: AbortSignal): Promise<DocsOperationResult>
  save(options: { path: string; overwrite: boolean }, signal?: AbortSignal): Promise<DocsSaveResult>
  close(): void
}

export interface DocsRuntime {
  connect(caller: RuntimeCaller): DocsRuntimeClient
  capture(): DocsRuntimeScope
  dispose(): void
}

export interface DocsRuntimeScope {
  /** Aborted when the captured document authority is closed or disposed. */
  readonly signal: AbortSignal
  readonly documentId: string
  readonly text: string
  connect(caller: RuntimeCaller): DocsRuntimeClient
  close(): void
}

interface ClientSeed {
  identity: object
  seen?: ProseMirrorNode
  seenState?: unknown
  frozen: FrozenSelection | null
  editor?: Editor
}

interface InternalClient extends DocsRuntimeClient {
  fork(caller: RuntimeCaller): InternalClient
}

const failure = (
  output: string,
  errorCode: DocsErrorCode = 'execution_failure',
): RuntimeExecution => ({
  output,
  summary: 'Document operation refused',
  isError: true,
  mutated: false,
  errorCode,
})
const executionFailure = (error: unknown): RuntimeExecution =>
  failure(
    error instanceof Error ? error.message : String(error),
    error instanceof ResourceAccessError ? 'permission_denied' : 'execution_failure',
  )
const toolSchemas = new Map(
  [...docsToolCatalog(), ...DOCS_RUNTIME_TOOLS, ...DOCS_RESOURCE_TOOLS].map((tool) => [
    tool.name,
    z.fromJSONSchema(tool.inputSchema),
  ]),
)

function withCancellation(
  operation: Promise<RuntimeExecution>,
  signal: AbortSignal,
): Promise<RuntimeExecution> {
  return new Promise((resolve, reject) => {
    const onClose = () =>
      resolve(failure('The operation was canceled; late output is refused', 'canceled'))
    if (signal.aborted) onClose()
    else signal.addEventListener('abort', onClose, { once: true })
    operation.then(
      (result) => {
        signal.removeEventListener('abort', onClose)
        resolve(result)
      },
      (error: unknown) => {
        signal.removeEventListener('abort', onClose)
        reject(error)
      },
    )
  })
}

/** The live document has one write queue, while every caller owns its read baseline. */
export function createDocsRuntime(bindings: DocsRuntimeBindings): DocsRuntime {
  const apply = bindings.apply ?? (<T>(operation: () => T): T => operation())
  let disposed = false
  let identity: object | undefined
  let documentId = ''
  let documentState: DocumentGenerationState = { documentId, revision: 0, queue: Promise.resolve() }
  let knownStateVersion: unknown
  let revisionEditor: Editor | undefined
  const countRevision = ({ transaction }: { transaction: Transaction }) => {
    if (transaction.docChanged && !transaction.before.eq(transaction.doc)) documentState.revision++
  }
  const clients = new Set<DocsRuntimeClient>()
  const pendingDrafts = new Set<DocumentDraft>()
  const queryContinuations = new Map<string, QueryContinuation>()

  const current = (): RuntimeDocument => {
    const bound = bindings.current()
    const doc = bound && { ...bound, stateVersion: stateVersionOf(bound) }
    if (disposed || !doc || doc.editor.isDestroyed) throw new Error('The document is not open')
    if (doc.identity !== identity) {
      // End callers of the replaced generation, including uncooperative pending native writers.
      for (const client of clients) client.close()
      for (const draft of pendingDrafts) draft.discard()
      pendingDrafts.clear()
      queryContinuations.clear()
      identity = doc.identity
      documentId = crypto.randomUUID()
      revisionEditor?.off('transaction', countRevision)
      documentState = { documentId, revision: 0, queue: Promise.resolve() }
      knownStateVersion = doc.stateVersion
      revisionEditor = doc.editor
      revisionEditor.on('transaction', countRevision)
    } else if (doc.stateVersion !== knownStateVersion) {
      knownStateVersion = doc.stateVersion
      documentState.revision++
    }
    return doc
  }

  const connect = (caller: RuntimeCaller, seed?: ClientSeed): InternalClient => {
    let closed = false
    const lifetime = new AbortController()
    const boundIdentity = seed?.identity ?? current().identity
    const boundState = documentState
    let seen = seed?.seen
    let seenState = seed?.seenState
    let revisionsSeen = seed?.seen
    let revisionState = seed?.seenState
    let frozen = seed?.frozen ?? null
    let observedEditor = seed?.editor
    let releaseCarry: (() => void) | undefined
    const carryRead = (before: ProseMirrorNode, after: ProseMirrorNode) => {
      if (seen === before) seen = after
      if (revisionsSeen === before) revisionsSeen = after
    }
    const drafts = new Map<string, DocumentDraft>()
    const resources = () => {
      const tools = caller.resources?.() ?? bindings.resources?.()
      if (tools?.tools.some((tool) => !DOCS_RESOURCE_TOOL_NAMES.has(tool.name)))
        throw new Error('The host registered an unsupported Docs resource helper')
      return tools
    }
    const resourceRefusal = (doc: RuntimeDocument) =>
      redactLabelsOf(doc.editor.getJSON()).length > 0
        ? {
            reason:
              'Raw file helpers are disabled while this document contains AI-withheld content. Use its projected live document tools.',
            errorCode: 'permission_denied' as const,
          }
        : undefined
    const createOwnedDraft = (doc: RuntimeDocument, position: WritePosition) => {
      if (hasActiveDocumentDraft(doc.editor)) throw new Error('A document draft is already active')
      const draftId = crypto.randomUUID()
      const sourceState = stateVersionOf(bindings.current())
      const draft = new DocumentDraft(
        doc.editor,
        doc.numIds,
        position,
        () => !doc.editor.isDestroyed && bindings.current()?.identity === boundIdentity,
        () => bindings.current()?.track,
        {
          documentId: boundState.documentId,
          revision: boundState.revision,
          isCurrent: () =>
            !closed &&
            !disposed &&
            bindings.current()?.identity === boundIdentity &&
            stateVersionOf(bindings.current()) === sourceState,
        },
      )
      drafts.set(draftId, draft)
      pendingDrafts.add(draft)
      return { draftId, draft }
    }
    const draftTool = (call: AgentToolCall, doc: RuntimeDocument): ToolExecution => {
      try {
        if (call.name === 'start_document_draft') {
          const { draftId, draft } = createOwnedDraft(doc, draftPosition(doc.editor, call.input))
          return {
            output: JSON.stringify({ draftId, ...draft.inspect() }),
            summary: 'Document draft started',
            mutated: false,
          }
        }
        const draftId = String(call.input.draftId)
        const draft = drafts.get(draftId)
        if (!draft) return failure('The draft does not belong to this caller')
        let mutated = false
        if (call.name === 'update_document_draft')
          draft.update(
            Number(call.input.sequence),
            String(call.input.html),
            call.input.deleteStructures,
          )
        if (call.name === 'stop_document_draft') draft.stop()
        if (call.name === 'finish_document_draft') {
          if (call.input.decision === 'discard') draft.discard()
          else {
            const wasPending = draft.pending
            const result = draft.finish(call.input.decision === 'retain_partial', doc.track)
            if (!result.ok) return failure(result.error, result.errorCode ?? 'execution_failure')
            mutated = wasPending
          }
        }
        if (!draft.pending) pendingDrafts.delete(draft)
        return {
          output: JSON.stringify({ draftId, ...draft.inspect() }),
          summary: 'Document draft',
          mutated,
        }
      } catch (error) {
        if (error instanceof DraftInputError)
          return {
            ...failure(error.message, error.errorCode),
            ...(error.imageReferenceErrors
              ? { imageReferenceErrors: error.imageReferenceErrors }
              : {}),
          }
        return failure(error instanceof Error ? error.message : String(error))
      }
    }
    const mapScope = ({ transaction }: { transaction: Transaction }) => {
      if (
        transaction.docChanged &&
        transaction.getMeta(DRAFT_META) &&
        [...drafts.values()].some((draft) => draft.pending)
      )
        carryRead(transaction.before, transaction.doc)
      if (!transaction.docChanged || !frozen || !observedEditor) return
      const from = transaction.mapping.map(frozen.scope.from!, -1)
      const to = Math.max(from, transaction.mapping.map(frozen.scope.to!, 1))
      frozen = {
        doc: observedEditor.state.doc,
        scope: getSelectionScope(observedEditor, { from, to, empty: from === to }),
      }
    }
    observedEditor?.on('transaction', mapScope)
    if (observedEditor) releaseCarry = onDocSeenCarry(observedEditor, carryRead)
    const client: InternalClient = {
      signal: lifetime.signal,
      fork(nextCaller) {
        if (closed || disposed || current().identity !== boundIdentity)
          throw new Error('The captured document binding is no longer valid')
        return connect(nextCaller, {
          identity: boundIdentity,
          seen,
          seenState,
          frozen,
          editor: observedEditor,
        })
      },
      context(options) {
        if (closed) throw new Error('The document caller is closed')
        const doc = current()
        if (doc.identity !== boundIdentity)
          throw new Error('The document was replaced; create a new document binding')
        if (observedEditor !== doc.editor) {
          observedEditor?.off('transaction', mapScope)
          releaseCarry?.()
          observedEditor = doc.editor
          observedEditor.on('transaction', mapScope)
          releaseCarry = onDocSeenCarry(observedEditor, carryRead)
        }
        revisionsSeen = seen = doc.editor.state.doc
        revisionState = seenState = doc.stateVersion
        if (options?.captureSelection !== false || !frozen)
          frozen = { scope: getSelectionScope(doc.editor), doc: seen }
        markDocSeen(doc.editor)
        return {
          documentId,
          revision: documentState.revision,
          text:
            bindDocumentImageHtml(
              buildDocContext(
                doc.editor,
                frozen.scope,
                doc.comments?.list(),
                doc.headerFooter?.read(),
                doc.pageSetup?.list(),
              ),
              documentId,
              documentState.revision,
            ) + documentMediaReferences(doc.editor, documentId, documentState.revision),
        }
      },
      execute(call, signal) {
        const operationSignal = signal
          ? AbortSignal.any([signal, lifetime.signal])
          : lifetime.signal
        let revisionBefore = boundState.revision
        let resolvedTargets: DocsEngineTarget[] = []
        let operationResults: NonNullable<DocsOperationResult['receipt']>['operations']
        let operationChanged = false
        const ownApply = <T>(operation: () => T): T => {
          const before = bindings.current()
          let bodyChanged = false
          const observe = ({ transaction }: { transaction: Transaction }) => {
            if (
              !transaction.getMeta(DRAFT_META) &&
              transaction.docChanged &&
              !transaction.before.eq(transaction.doc)
            )
              bodyChanged = true
          }
          before?.editor.on('transaction', observe)
          const state = stateVersionOf(before)
          try {
            return apply(operation)
          } finally {
            before?.editor.off('transaction', observe)
            const after = bindings.current()
            if (
              before &&
              after &&
              before.identity === after.identity &&
              (bodyChanged || state !== stateVersionOf(after))
            )
              operationChanged = true
          }
        }
        const run = async (): Promise<RuntimeExecution> => {
          revisionBefore = boundState.revision
          if (closed || disposed) return failure('The document caller is closed', 'stale_target')
          if (operationSignal.aborted)
            return failure('stopped by the user; the operation was canceled', 'canceled')
          const doc = current()
          revisionBefore = boundState.revision
          const missingStoreEpoch = authoredStoreUnavailable(call.name, bindings.current())
          if (missingStoreEpoch)
            return failure(missingStoreEpoch.reason, missingStoreEpoch.errorCode)
          const resourceProvider = resources()
          const helper = resourceProvider?.tools.find((tool) => tool.name === call.name)
          const schema =
            toolSchemas.get(call.name) ?? (helper && z.fromJSONSchema(helper.inputSchema))
          if (!schema) return failure(`Unknown Docs tool: ${call.name}`, 'unavailable')
          const validated = schema.safeParse(call.input)
          if (!validated.success)
            return failure(
              `Invalid ${call.name} input: ${validated.error.message}`,
              'invalid_input',
            )
          if (boundIdentity !== doc.identity)
            return failure(
              'The document was replaced; create a new document binding',
              'stale_target',
            )
          if (call.name === 'query_document') {
            try {
              const result = queryDocument(
                doc.editor,
                boundState.documentId,
                boundState.revision,
                call.input,
                doc,
                queryContinuations,
              )
              seen = doc.editor.state.doc
              markDocSeen(doc.editor)
              return { output: JSON.stringify(result), summary: 'Document query', mutated: false }
            } catch (error) {
              const message = error instanceof Error ? error.message : String(error)
              return failure(
                message,
                message.startsWith('Stale') ? 'stale_target' : 'invalid_input',
              )
            }
          }
          if (call.name === 'apply_ops' && Array.isArray(call.input.ops)) {
            resolvedTargets = call.input.ops.flatMap((op) =>
              op?.target?.engine ? [op.target.engine] : [],
            )
            if (
              resolvedTargets.some(
                (target) =>
                  target.documentId !== boundState.documentId ||
                  target.revision !== boundState.revision,
              )
            )
              return failure(
                'Stale or wrong-document target; use query_document again before editing.',
                'stale_target',
              )
            for (const target of resolvedTargets) {
              if (target.expectedHtml === undefined) continue
              if (projectedTargetHtml(doc.editor, target) !== target.expectedHtml)
                return failure(
                  'Engine target content precondition failed; query_document again.',
                  'invalid_input',
                )
            }
          }
          const localIntent = (value: unknown): string[] | undefined => {
            if (value === undefined) return undefined
            if (!Array.isArray(value))
              throw new Error('deleteStructures must contain revision-bound structure references')
            return value.map((id) => {
              if (typeof id !== 'string') throw new Error('Invalid structure reference')
              const prefix = `structure:${boundState.documentId}:${boundState.revision}:`
              if (!id.startsWith(prefix))
                throw new Error('Stale or unbound structure reference; query_document again')
              return id.slice(prefix.length)
            })
          }
          try {
            if (call.input.deleteStructures !== undefined && call.name !== 'update_document_draft')
              call = {
                ...call,
                input: {
                  ...call.input,
                  deleteStructures: localIntent(call.input.deleteStructures),
                },
              }
            if (call.name === 'apply_ops' && Array.isArray(call.input.ops))
              call = {
                ...call,
                input: {
                  ...call.input,
                  ops: call.input.ops.map((op) =>
                    op && typeof op === 'object' && op.deleteStructures !== undefined
                      ? { ...op, deleteStructures: localIntent(op.deleteStructures) }
                      : op,
                  ),
                },
              }
          } catch (error) {
            return failure(error instanceof Error ? error.message : String(error), 'stale_target')
          }
          if (call.input.engineTarget !== undefined) {
            const target = call.input.engineTarget as DocsEngineTarget
            if (
              target.documentId !== boundState.documentId ||
              target.revision !== boundState.revision
            )
              return failure('Stale side-state target; query_document again.', 'stale_target')
            if (
              !target.address ||
              !Number.isSafeInteger(target.address.section) ||
              target.address.section < 0
            )
              return failure('Invalid side-state target address', 'invalid_input')
            const domain =
              call.name === 'set_page_setup'
                ? 'sections'
                : target.address.kind === 'header'
                  ? 'headers'
                  : 'footers'
            const windows = new Map<string, QueryContinuation>()
            let window = queryDocument(
              doc.editor,
              boundState.documentId,
              boundState.revision,
              { domain, limit: 50, maxItemChars: 4000 },
              doc,
              windows,
            )
            let found: boolean
            while (true) {
              found = window.items.some(
                (item) =>
                  item.target.expectedType === target.expectedType &&
                  item.target.kind === target.kind &&
                  item.supportedOperations.includes(call.name) &&
                  item.target.address?.section === target.address?.section &&
                  item.target.address?.kind === target.address?.kind &&
                  item.target.address?.view === target.address?.view &&
                  item.target.address?.textboxIndex === target.address?.textboxIndex &&
                  (target.expectedHtml === undefined ||
                    (item.complete && item.html === target.expectedHtml)),
              )
              if (found || !window.page.nextCursor) break
              window = queryDocument(
                doc.editor,
                boundState.documentId,
                boundState.revision,
                { cursor: window.page.nextCursor },
                doc,
                windows,
              )
            }
            if (!found || (call.name !== 'set_page_setup' && call.name !== 'set_header_footer'))
              return failure('Target type/address does not support this operation', 'invalid_input')
            for (const [key, value] of Object.entries(target.address))
              if (call.input[key] !== undefined && call.input[key] !== value)
                return failure(
                  'Explicit parameters disagree with the engine target',
                  'invalid_input',
                )
            resolvedTargets = [target]
            call = { ...call, input: { ...call.input, ...target.address } }
          }
          if (['insert_content', 'replace_blocks', 'replace_selection'].includes(call.name)) {
            const images = prepareDocumentImageWrite(
              doc.editor,
              String(call.input.html ?? ''),
              boundState.documentId,
              boundState.revision,
            )
            if (images.errors.length)
              return {
                ...failure(
                  'Explicit image references could not be applied; no changes were made. Use current media: references from read_blocks; use insert_picture for new images. Protected block/OLE previews are not writable image references.',
                  images.errors.some((error) => error.reason === 'expired')
                    ? 'stale_target'
                    : 'invalid_input',
                ),
                mutated: false,
                imageReferenceErrors: images.errors,
              }
            // Reuse execution's representation checks before any history metadata is dispatched.
            if (/<img\b/i.test(images.html)) {
              try {
                const nodes =
                  call.name === 'replace_selection'
                    ? parseInlineFragment(images.html)
                    : parseHtmlFragment(images.html, doc.numIds)
                resolveImageRefs(doc.editor, nodes, images.html)
              } catch (error) {
                return {
                  ...failure(
                    error instanceof Error ? error.message : String(error),
                    'invalid_input',
                  ),
                  mutated: false,
                }
              }
            }
            call = { ...call, input: { ...call.input, html: images.html } }
          }
          if (call.name === 'get_editor_capabilities') {
            const helperTools = resources()?.tools ?? []
            const tools = [
              ...docsToolCatalog(opNames()),
              ...DOCS_RUNTIME_TOOLS,
              ...DOCS_RESOURCE_TOOLS,
            ].map((tool) => {
              const unavailable = DOCS_RESOURCE_TOOL_NAMES.has(tool.name)
                ? helperTools.some((helper) => helper.name === tool.name)
                  ? resourceRefusal(doc)
                  : {
                      reason: 'This resource helper is not available in the current host context',
                      errorCode: 'unavailable' as const,
                    }
                : (authoredStoreUnavailable(tool.name, bindings.current()) ??
                  toolUnavailable(tool.name, doc))
              return { ...tool, available: !unavailable, ...(unavailable ?? {}) }
            })
            return {
              output: JSON.stringify({
                documentId,
                editable: doc.editor.isEditable,
                tools,
                resourceContext: resourceRefusal(doc) ? '' : (resources()?.buildContext?.() ?? ''),
                htmlRules: HTML_RULES,
                operations: opCatalog(),
              }),
              summary: 'Document capabilities',
              mutated: false,
            }
          }
          if (DOCS_RESOURCE_TOOL_NAMES.has(call.name)) {
            if (!helper)
              return failure(
                'This resource helper is not available in the current host context',
                'unavailable',
              )
            const denied = resourceRefusal(doc)
            if (denied) return failure(denied.reason, denied.errorCode)
            if (
              call.name === 'run_cli' &&
              !DOCS_FILE_COMMANDS.has(String((call.input.args as unknown[] | undefined)?.[0] ?? ''))
            )
              return failure(
                'This CLI command is not a file-only Office operation. Run inference and provider discovery in the external Harness.',
                'unavailable',
              )
            const independentFiles = call.name === 'run_cli'
            const operation = Promise.resolve()
              .then(() =>
                operationSignal.aborted
                  ? {
                      ...failure('The resource operation was canceled before starting', 'canceled'),
                      committed: false,
                    }
                  : resourceProvider!.executeTool(call, operationSignal),
              )
              .catch((error): RuntimeExecution => ({
                ...executionFailure(error),
                ...(independentFiles ? { committed: null } : {}),
              }))
            // File helpers must acknowledge IO; both their data and diagnostics still
            // require a current source grant when the result is delivered.
            const result: RuntimeExecution = independentFiles
              ? await operation
              : await withCancellation(operation, operationSignal)
            const refuseDelivery = (message: string, code: DocsErrorCode): RuntimeExecution => {
              const refused = failure(message, code)
              if (!independentFiles) return refused
              const committed =
                result.committed === undefined
                  ? result.mutated
                    ? true
                    : result.isError
                      ? null
                      : false
                  : result.committed
              return {
                ...refused,
                committed,
                mutated: !!result.mutated,
                output: `${message}. Independent file effect: ${committed === null ? 'unknown' : committed ? 'committed' : 'not committed'}. Source output was withheld.`,
              }
            }
            if (operationSignal.aborted)
              return refuseDelivery('The resource operation was canceled', 'canceled')
            const latest = bindings.current()
            if (closed || disposed || !latest || latest.identity !== boundIdentity)
              return refuseDelivery('The resource binding is no longer valid', 'stale_target')
            if (
              resources() !== resourceProvider ||
              !resourceProvider!.tools.some((tool) => tool.name === call.name)
            )
              return refuseDelivery(
                'Resource access was revoked during the operation',
                'permission_denied',
              )
            const revoked = resourceRefusal(latest)
            return revoked ? refuseDelivery(revoked.reason, revoked.errorCode) : result
          }
          const unavailable =
            call.name === 'finish_document_draft' && call.input.decision === 'discard'
              ? undefined
              : toolUnavailable(call.name, doc)
          if (unavailable) return failure(unavailable.reason, unavailable.errorCode)
          if (call.name === 'read_document_media') {
            const expected = doc.editor.state.doc
            let blockIndex = Number(call.input.blockIndex)
            let imageIndex =
              call.input.imageIndex === undefined ? undefined : Number(call.input.imageIndex)
            if (call.input.mediaRef !== undefined) {
              const ref = /^media:([^:]+):(\d+):(\d+)\.(\d+)$/.exec(String(call.input.mediaRef))
              if (!ref) return failure('Invalid document media reference.', 'invalid_input')
              if (ref[1] !== boundState.documentId || Number(ref[2]) !== boundState.revision)
                return failure('The media reference is stale; read fresh context.', 'stale_target')
              if (
                (call.input.blockIndex !== undefined && blockIndex !== Number(ref[3])) ||
                (imageIndex !== undefined && imageIndex !== Number(ref[4]))
              )
                return failure('mediaRef and image coordinates disagree.', 'invalid_input')
              blockIndex = Number(ref[3])
              imageIndex = Number(ref[4])
            }
            return withCancellation(
              readDocumentMedia(
                doc.editor,
                blockIndex,
                operationSignal,
                () =>
                  !closed &&
                  !disposed &&
                  bindings.current()?.identity === boundIdentity &&
                  stateVersionOf(bindings.current()) === doc.stateVersion &&
                  seen === expected &&
                  doc.editor.state.doc === expected,
                imageIndex,
              ),
              operationSignal,
            )
          }
          const reviewsRevisions = call.name === 'accept_changes' || call.name === 'reject_changes'
          const readsBlocks = call.name === 'get_document_context' || call.name === 'read_blocks'
          const blocksWereCurrent = seen === doc.editor.state.doc
          const revisionsWereCurrent = revisionsSeen === doc.editor.state.doc
          const stateWasCurrent = seenState === doc.stateVersion
          const writesDocument =
            !(call.name === 'apply_ops' && call.input.dryRun === true) &&
            !DOCS_READ_TOOLS.has(call.name) &&
            !DOCS_DRAFT_CONTROLS.has(call.name) &&
            call.name !== 'start_document_draft' &&
            call.name !== 'create_document'
          if (
            !DOCS_READ_TOOLS.has(call.name) &&
            !DOCS_DRAFT_CONTROLS.has(call.name) &&
            (reviewsRevisions
              ? !revisionsWereCurrent || revisionState !== doc.stateVersion
              : !blocksWereCurrent || !stateWasCurrent)
          )
            return failure(
              reviewsRevisions
                ? 'Document revisions or document-side state changed since this caller read them; use read_revisions before accepting or rejecting changes, or read_document or get_document_context for fresh full context.'
                : !stateWasCurrent
                  ? 'Document-side state changed since this caller read it; use read_document or get_document_context for fresh full context. read_blocks does not refresh document-side state.'
                  : 'The document body changed since this caller read it; use read_blocks, read_document or get_document_context before retrying.',
              'stale_target',
            )
          const isDraftTool =
            call.name === 'start_document_draft' || DOCS_DRAFT_CONTROLS.has(call.name)
          const structureError = validateHtmlStructureIntent(
            doc.editor,
            call,
            frozen?.doc === doc.editor.state.doc ? frozen.scope : null,
            doc.numIds,
          )
          if (structureError) return failure(structureError, 'invalid_input')
          if (hasActiveDocumentDraft(doc.editor) && !isDraftTool && !DOCS_READ_TOOLS.has(call.name))
            return failure('Settle the active document draft before editing')
          if (call.name === 'document_history') {
            const changed = ownApply(() =>
              call.input.action === 'undo'
                ? doc.editor.commands.undo()
                : doc.editor.commands.redo(),
            )
            if (!changed) return failure(`No ${String(call.input.action)} is available`)
            revisionsSeen = seen = doc.editor.state.doc
            revisionState = seenState = stateVersionOf(bindings.current())
            return {
              output: `Document ${String(call.input.action)} completed`,
              summary: 'Document history',
              mutated: true,
            }
          }
          // These stores have their own host persistence; they do not form body undo entries.
          const storeOnly = [
            'set_header_footer',
            'set_page_setup',
            'set_watermark',
            'define_style',
            'reply_comment',
            'resolve_comment',
            'delete_comment',
            'edit_note',
          ].includes(call.name)
          let changesHistory = !DOCS_READ_TOOLS.has(call.name) && !storeOnly && !isDraftTool
          if (call.name === 'apply_ops') {
            const preview = executeOps(doc.editor, call.input.ops, {
              numIds: doc.numIds,
              track: doc.track,
              selection: frozen?.doc === doc.editor.state.doc ? frozen.scope : null,
              styles: doc.extras?.styles,
              dryRun: true,
              engineEpoch: boundState,
            })
            if (!preview.ok)
              return failure(preview.error ?? 'The operation batch cannot execute', 'invalid_input')
            changesHistory = call.input.dryRun !== true && preview.wouldChange === true
          }
          if (changesHistory) doc.editor.view.dispatch(closeHistory(doc.editor.state.tr))
          let expectedDoc = doc.editor.state.doc
          let expectedState = doc.stateVersion
          markDocSeen(doc.editor)
          const stopExpectedCarry = onDocSeenCarry(doc.editor, (before, after) => {
            if (expectedDoc === before) expectedDoc = after
          })
          let result: RuntimeExecution
          try {
            const operation = Promise.resolve(
              ownApply(() =>
                isDraftTool
                  ? draftTool(call, doc)
                  : call.name === 'apply_ops'
                    ? (() => {
                        const outcome = executeOps(doc.editor, call.input.ops, {
                          numIds: doc.numIds,
                          track: doc.track,
                          selection: frozen?.doc === doc.editor.state.doc ? frozen.scope : null,
                          styles: doc.extras?.styles,
                          dryRun: call.input.dryRun === true,
                          engineEpoch: boundState,
                        })
                        if (!outcome.ok)
                          return failure(
                            outcome.error ?? 'The operation batch failed',
                            'invalid_input',
                          )
                        const results = outcome.previewResults ?? outcome.results
                        operationResults = results.map((result) => ({
                          ...result,
                          ...(result.structuralEffects
                            ? {
                                structuralEffects: result.structuralEffects.map((effect) => ({
                                  ...effect,
                                  reference: `structure:${boundState.documentId}:${revisionBefore}:${effect.reference}`,
                                })),
                              }
                            : {}),
                        }))
                        return {
                          output:
                            call.input.dryRun === true
                              ? `Dry run: ${(outcome.plan ?? []).join('\n')}`
                              : outcome.summary,
                          summary: 'Document operations',
                          mutated:
                            call.input.dryRun !== true &&
                            results.some((result) => result.changed > 0),
                          matchCount: results.reduce(
                            (sum, result) => sum + (result.matchCount ?? result.matched),
                            0,
                          ),
                          changeCount: results.reduce(
                            (sum, result) => sum + (result.changeCount ?? result.changed),
                            0,
                          ),
                        }
                      })()
                    : executeTool(
                        doc.editor,
                        call,
                        doc.numIds,
                        doc.track,
                        operationSignal,
                        frozen,
                        doc.comments,
                        doc.headerFooter,
                        doc.pageSetup,
                        doc.extras,
                        doc.notes,
                        {
                          isCurrent: () =>
                            !closed &&
                            !disposed &&
                            bindings.current()?.identity === boundIdentity &&
                            (DOCS_READ_TOOLS.has(call.name) || doc.editor.isEditable) &&
                            stateVersionOf(bindings.current()) === expectedState &&
                            expectedDoc === doc.editor.state.doc,
                          markSeen: () => {
                            expectedDoc = doc.editor.state.doc
                            expectedState = stateVersionOf(bindings.current())
                            if (blocksWereCurrent || readsBlocks) seen = expectedDoc
                          },
                          apply: ownApply,
                          selection: () => frozen?.scope ?? null,
                          imageHtml: (html) =>
                            bindDocumentImageHtml(html, boundState.documentId, boundState.revision),
                        },
                      ),
              ),
            )
            // Creation has independent host effects and no rollback contract. Await its
            // acknowledged outcome even when this document/caller has since closed.
            result = await (call.name === 'create_document'
              ? operation
              : withCancellation(operation, operationSignal))
          } finally {
            stopExpectedCarry()
          }
          for (const draft of pendingDrafts) if (!draft.pending) pendingDrafts.delete(draft)
          if (result.mutated && !doc.track && doc.autoAcceptUntracked) {
            let transaction = doc.editor.state.tr
            doc.editor.state.doc.forEach((node, offset) => {
              if (node.attrs.aiChanged)
                transaction = transaction.setNodeMarkup(offset, undefined, {
                  ...node.attrs,
                  aiChanged: false,
                })
            })
            if (transaction.docChanged)
              doc.editor.view.dispatch(transaction.setMeta('addToHistory', false))
          }
          if (!result.isError) {
            if (call.name === 'get_document_context')
              result.output = bindDocumentImageHtml(
                result.output,
                boundState.documentId,
                boundState.revision,
              )
            if (readsBlocks)
              result.output += documentMediaReferences(
                doc.editor,
                boundState.documentId,
                boundState.revision,
                call.name === 'read_blocks' ? Number(call.input.startBlockIndex) : 0,
                call.name === 'read_blocks'
                  ? Number(call.input.endBlockIndex)
                  : doc.editor.state.doc.childCount - 1,
              )
            const nextState = stateVersionOf(bindings.current())
            if (readsBlocks || (result.mutated && blocksWereCurrent)) seen = doc.editor.state.doc
            if (
              readsBlocks ||
              call.name === 'read_revisions' ||
              (result.mutated && revisionsWereCurrent)
            )
              revisionsSeen = doc.editor.state.doc
            if (
              call.name === 'get_document_context' ||
              ((result.mutated || writesDocument) && stateWasCurrent)
            )
              seenState = nextState
            if (
              call.name === 'get_document_context' ||
              call.name === 'read_revisions' ||
              ((result.mutated || writesDocument) && revisionState === doc.stateVersion)
            )
              revisionState = nextState
          }
          if (result.mutated) doc.editor.view.dispatch(closeHistory(doc.editor.state.tr))
          return result
        }
        const result = boundState.queue
          .then(run)
          .catch(executionFailure)
          .then((execution): DocsOperationResult => {
            if (!disposed && bindings.current()?.identity === boundIdentity) current()
            const helper = DOCS_RESOURCE_TOOL_NAMES.has(call.name)
            const changed = operationChanged
            const independent = helper || call.name === 'create_document'
            const committed = independent
              ? execution.committed === undefined
                ? !!execution.mutated
                : execution.committed
              : changed
            return {
              ...execution,
              documentId: boundState.documentId,
              revision: boundState.revision,
              committed,
              mutated: helper || call.name === 'create_document' ? execution.mutated : changed,
              receipt: {
                outcome:
                  helper || call.name === 'create_document'
                    ? 'independent-effect'
                    : execution.isError
                      ? changed
                        ? 'failure-after-commit'
                        : 'rejected'
                      : call.name === 'apply_ops' && call.input.dryRun === true
                        ? 'preview'
                        : changed
                          ? 'applied'
                          : DOCS_READ_TOOLS.has(call.name)
                            ? 'read'
                            : 'no-op',
                revisionBefore,
                revisionAfter: boundState.revision,
                persistence:
                  helper || call.name === 'create_document'
                    ? committed === null
                      ? 'unknown'
                      : 'unchanged'
                    : changed
                      ? 'pending'
                      : 'unchanged',
                ...(resolvedTargets.length ? { resolvedTargets } : {}),
                ...(operationResults ? { operations: operationResults } : {}),
              },
              ...(execution.isError
                ? { errorCode: execution.errorCode ?? 'execution_failure' }
                : {}),
            }
          })
        boundState.queue = result.catch(() => undefined)
        return result
      },
      save(options, signal) {
        const run = async (): Promise<DocsSaveResult> => {
          const revisionBefore = boundState.revision
          const refuse = (error: string, errorCode: DocsErrorCode): DocsSaveOutcome => ({
            ok: false,
            error,
            errorCode,
            committed: false,
            pending: true,
          })
          let outcome: DocsSaveOutcome
          if (
            closed ||
            disposed ||
            !bindings.current() ||
            bindings.current()?.editor.isDestroyed ||
            current().identity !== boundIdentity
          )
            outcome = refuse('The document binding is no longer valid', 'stale_target')
          else if (signal?.aborted)
            outcome = refuse('The save was canceled before writing', 'canceled')
          else if (!bindings.save)
            outcome = refuse('The host does not provide saving', 'unavailable')
          else if (hasActiveDocumentDraft(current().editor))
            outcome = refuse('Settle the active document draft before saving', 'invalid_input')
          else {
            try {
              outcome = await bindings.save(options)
            } catch (error) {
              outcome = {
                ok: false,
                error: error instanceof Error ? error.message : String(error),
                errorCode: 'host_failure',
                committed: null,
                pending: true,
              }
            }
            if (
              closed ||
              disposed ||
              bindings.current()?.identity !== boundIdentity ||
              bindings.current()?.editor.isDestroyed
            )
              outcome = {
                ...outcome,
                ok: false,
                pending: true,
                errorCode: 'stale_target',
                error:
                  'The document binding changed during saving; the reported file effect belongs to the previous document',
              }
            else current()
          }
          return {
            ...outcome,
            documentId: boundState.documentId,
            revision: boundState.revision,
            output: outcome.ok
              ? `Document saved to ${outcome.path ?? options.path}`
              : (outcome.error ?? 'The document could not be saved'),
            summary: outcome.ok ? 'Document saved' : 'Document save incomplete',
            receipt: {
              outcome: 'independent-effect',
              revisionBefore,
              revisionAfter: boundState.revision,
              persistence:
                outcome.ok && outcome.committed === true && !outcome.pending
                  ? 'completed'
                  : outcome.committed === null
                    ? 'unknown'
                    : 'pending',
            },
            ...(!outcome.ok
              ? { isError: true, errorCode: outcome.errorCode ?? 'host_failure' }
              : {}),
          }
        }
        const result = boundState.queue.then(run)
        boundState.queue = result.catch(() => undefined)
        return result
      },
      close() {
        closed = true
        lifetime.abort()
        for (const draft of drafts.values()) {
          draft.discard()
          pendingDrafts.delete(draft)
        }
        drafts.clear()
        observedEditor?.off('transaction', mapScope)
        releaseCarry?.()
        clients.delete(client)
      },
    }
    clients.add(client)
    return client
  }
  return {
    connect,
    capture() {
      const seed = connect({ id: crypto.randomUUID() })
      const context = seed.context()
      const scopedClients = new Set<DocsRuntimeClient>()
      return {
        ...context,
        signal: seed.signal,
        connect(caller) {
          const client = seed.fork(caller)
          scopedClients.add(client)
          return client
        },
        close() {
          seed.close()
          for (const client of scopedClients) client.close()
          scopedClients.clear()
        },
      }
    },
    dispose() {
      disposed = true
      for (const client of clients) client.close()
      revisionEditor?.off('transaction', countRevision)
    },
  }
}
