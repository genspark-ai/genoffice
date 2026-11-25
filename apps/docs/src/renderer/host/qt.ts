import { createExternalDocsHost, type DocsEditorServices } from './docs-host'
import type { McpCommandMessage } from '../../shared/ipc'
import type { DocsAiInstructionRequest, DocsEditorAction } from '../../shared/editor-action'
import type { DocsOperationResult } from '../../shared/operation-result'
import { createDocsResourceTools } from '../runtime/resource-tools'
import { ResourceAccessError } from '@genoffice/agent-core/resource-tools'
import type { AttachmentMeta } from '../../shared/ipc'
import type { DocsInspection } from '../sdk/inspection'
import type { DocsAnnotations } from '../sdk/annotations'
import type { DocsEditorUi } from '../sdk/file-menu'
import type {
  DocsDocumentCommands,
  DocsDocumentHost,
  DocsHostedDocument,
} from '../../shared/document-host'
import { executeDocsSdkCall } from '../sdk/calls'
import type { DocsRuntime, DocsRuntimeScope } from '../runtime/docs-runtime'

export interface QtSignal<Args extends unknown[]> {
  connect(listener: (...args: Args) => void): void
  disconnect(listener: (...args: Args) => void): void
}

/** The published QWebChannel object. It transports native-host IPC, never model requests. */
export interface DocsQtChannel {
  request(message: string): void
  complete(message: string): void
  reply: QtSignal<[string]>
  command: QtSignal<[string]>
  event: QtSignal<[string, string]>
}

interface EditorEndpoint {
  runtime?: Pick<DocsRuntime, 'capture'>
  inspection?: DocsInspection
  annotations?: DocsAnnotations
  documents?: DocsDocumentCommands
  ui?: DocsEditorUi
  commands: {
    bindScope?(bindingId: string, scope: DocsRuntimeScope, onRelease?: () => void): void
    handle(message: McpCommandMessage): Promise<unknown>
    releaseScope(bindingId: string): void
    disconnect(callerId: string): void
  }
}

function encode(value: unknown): unknown {
  if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) {
    const bytes =
      value instanceof ArrayBuffer
        ? new Uint8Array(value)
        : new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
    let binary = ''
    for (let offset = 0; offset < bytes.length; offset += 32_768)
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 32_768))
    return { base64: btoa(binary) }
  }
  if (Array.isArray(value)) return value.map(encode)
  if (value && typeof value === 'object')
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, encode(item)]))
  return value
}

const notifications = new Set([
  'reportCreatedDocument',
  'reportDocumentPath',
  'headlessExportDone',
  'respondToZotero',
  'spellDiag',
  'armContextMenu',
  'claimContextMenu',
  'reportMcpResult',
  'signalMcpReady',
  'reportCloseCheck',
  'reportCloseSaveResult',
  'reportViewMenuState',
])

/** Mounting uses the normal SDK; Qt owns native facilities and the separate Node process. */
export function createQtDocsHost(channel: DocsQtChannel) {
  let disposed = false
  let editor: EditorEndpoint | undefined
  let editorDisposers: Array<() => void> = []
  const pending = new Map<string, { resolve(value: unknown): void; reject(error: unknown): void }>()
  const listeners = new Map<string, Set<(...args: unknown[]) => void>>()
  let attachments: AttachmentMeta[] = []
  let knowledge: Awaited<ReturnType<DocsEditorServices['kbList']>> = []
  let resources: ReturnType<typeof createDocsResourceTools> | undefined
  const notify = (method: string, params: unknown) =>
    channel.request(JSON.stringify({ jsonrpc: '2.0', method, params: encode(params) }))
  const request = (
    method: string,
    params: unknown,
    signal?: AbortSignal,
    cancel?: () => void,
  ): Promise<unknown> =>
    new Promise((resolve, reject) => {
      if (disposed) {
        reject(new Error('The Qt editor host was disposed'))
        return
      }
      if (signal?.aborted) {
        reject(signal.reason)
        return
      }
      const id = crypto.randomUUID()
      const cleanup = () => signal?.removeEventListener('abort', abort)
      const abort = () => {
        if (!pending.delete(id)) return
        cleanup()
        try {
          cancel?.()
        } catch (error) {
          reject(error)
          return
        }
        reject(signal?.reason)
      }
      pending.set(id, {
        resolve: (value) => {
          cleanup()
          resolve(value)
        },
        reject: (error) => {
          cleanup()
          reject(error)
        },
      })
      signal?.addEventListener('abort', abort, { once: true })
      try {
        channel.request(JSON.stringify({ jsonrpc: '2.0', id, method, params: encode(params) }))
      } catch (error) {
        pending.delete(id)
        cleanup()
        reject(error)
      }
    })
  const reply = (source: string) => {
    const message = JSON.parse(source) as {
      jsonrpc: string
      id: string
      result?: unknown
      error?: { message: string; data?: { errorCode?: string } }
    }
    if (message.jsonrpc !== '2.0') throw new Error('Invalid Qt host response')
    const operation = pending.get(message.id)
    if (!operation) return
    pending.delete(message.id)
    if (message.error)
      operation.reject(
        Object.assign(new Error(message.error.message), {
          errorCode: message.error.data?.errorCode,
        }),
      )
    else operation.resolve(message.result)
  }
  const event = (name: string, source: string) => {
    if (disposed) return
    const args = JSON.parse(source) as unknown[]
    for (const listener of listeners.get(name) ?? []) listener(...args)
  }
  const command = (source: string) => {
    if (disposed) return
    const message = JSON.parse(source) as {
      id?: string
      method: string
      params: {
        message: McpCommandMessage
        bindingId: string
        callerId: string
      }
    }
    if (message.method === 'editor.bindMcp') {
      try {
        if (!editor?.runtime || !editor.commands.bindScope)
          throw new Error('The live document runtime is not ready')
        const scope = editor.runtime.capture()
        const { bindingId } = message.params
        try {
          editor.commands.bindScope(bindingId, scope, () => {
            if (!disposed) notify('editor.mcpRevoked', { bindingId })
          })
        } catch (error) {
          scope.close()
          throw error
        }
        channel.complete(
          JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: { bindingId, documentId: scope.documentId },
          }),
        )
      } catch (error) {
        channel.complete(
          JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            error: {
              code: -32000,
              message: error instanceof Error ? error.message : String(error),
            },
          }),
        )
      }
    } else if (message.method === 'sdk.request') {
      const params = message.params as unknown as {
        operation: string
        args: Record<string, unknown>
      }
      void Promise.resolve()
        .then(() => {
          if (!editor?.inspection || !editor.annotations || !editor.documents)
            throw new Error('The public document SDK is not ready')
          return executeDocsSdkCall(
            {
              inspection: editor.inspection,
              annotations: editor.annotations,
              documents: editor.documents,
              ui: editor.ui,
            },
            params.operation,
            params.args,
          )
        })
        .then(
          (result) => {
            if (!disposed)
              channel.complete(JSON.stringify({ jsonrpc: '2.0', id: message.id, result }))
          },
          (error: unknown) => {
            if (!disposed)
              channel.complete(
                JSON.stringify({
                  jsonrpc: '2.0',
                  id: message.id,
                  error: {
                    code: -32000,
                    message: error instanceof Error ? error.message : String(error),
                  },
                }),
              )
          },
        )
    } else if (message.method === 'editor.releaseScope')
      editor?.commands.releaseScope(message.params.bindingId)
    else if (message.method === 'editor.disconnect')
      editor?.commands.disconnect(message.params.callerId)
    else if (message.method === 'editor.cancel') {
      const request = message.params.message
      void editor?.commands.handle({
        ...request,
        command: 'cancel_operation',
        payload: { requestId: request.requestId },
      })
    } else if (message.method === 'editor.request') {
      const request = message.params.message
      const result = editor
        ? editor.commands.handle(request)
        : Promise.reject(new Error('The Qt document is not ready'))
      void result
        .then(
          (value) => ({ requestId: request.requestId, ok: true, result: value }),
          (error: unknown) => ({
            requestId: request.requestId,
            ok: false,
            error: error instanceof Error ? error.message : String(error),
            ...(error instanceof Error && 'result' in error
              ? { result: error.result as DocsOperationResult }
              : {}),
          }),
        )
        .then((value) => {
          if (!disposed)
            channel.complete(JSON.stringify({ jsonrpc: '2.0', id: message.id, result: value }))
        })
    }
  }
  channel.reply.connect(reply)
  channel.event.connect(event)
  channel.command.connect(command)
  const services = new Proxy({} as DocsEditorServices, {
    get(_target, key) {
      if (typeof key !== 'string' || key === 'then') return undefined
      // WebEngine deliberately does not disclose an absolute path for a browser File.
      if (key === 'getPathForFile') return () => ''
      if (key.startsWith('on'))
        return (listener: (...args: unknown[]) => void) => {
          const entries = listeners.get(key) ?? new Set()
          entries.add(listener)
          listeners.set(key, entries)
          return () => {
            entries.delete(listener)
          }
        }
      if (notifications.has(key))
        return (...args: unknown[]) => {
          if (!disposed) notify('services.notify', { method: key, args })
        }
      return (...args: unknown[]) => request('services.call', { method: key, args })
    },
  })
  const host = createExternalDocsHost(
    'qt',
    services,
    {
      resolveInstruction: async (action: DocsAiInstructionRequest, signal: AbortSignal) => {
        const result = await request('editor.resolveInstruction', action, signal, () =>
          notify('editor.cancelInstruction', { id: action.id }),
        )
        if (result === null || result === undefined || typeof result === 'string') return result
        throw new Error('The Qt instruction resolver returned an invalid result')
      },
      submit: async (action: DocsEditorAction) => {
        await request('editor.submitAction', action)
      },
    },
    () => resources,
  )
  host.supportsPrintScale = false
  const sourceRead = async <T>(operation: () => Promise<T>): Promise<T> => {
    try {
      return await operation()
    } catch (error) {
      if (error instanceof Error && 'errorCode' in error && error.errorCode === 'permission_denied')
        throw new ResourceAccessError(error.message)
      throw error
    }
  }
  let cliAvailable: boolean | undefined
  let attachmentsAvailable = false
  const updateResources = (state: {
    attachments: AttachmentMeta[]
    knowledge: typeof knowledge
    cliAvailable: boolean
    documentHost?: Pick<DocsDocumentHost, 'capabilities' | 'autoSave'> & { enabled: boolean }
    checkers?: string[]
  }) => {
    if (state.checkers)
      host.checkers = Object.fromEntries(
        state.checkers.map((source) => [
          source,
          async (snapshot: import('../sdk/inspection').DocsTextSnapshot) =>
            request('services.call', {
              method: 'checkDocument',
              args: [source, snapshot],
            }) as Promise<import('../sdk/annotations').DocsAnnotation[]>,
        ]),
      )
    if (state.documentHost?.enabled && !host.documents) {
      const document = async (
        method: string,
        args: unknown[] = [],
      ): Promise<DocsHostedDocument | null> => {
        const result = (await request('services.call', { method, args })) as
          (Omit<DocsHostedDocument, 'data'> & { data?: { base64: string } }) | null
        if (!result) return null
        return {
          ...result,
          ...(result.data
            ? {
                data: Uint8Array.from(atob(result.data.base64), (char) => char.charCodeAt(0))
                  .buffer,
              }
            : {}),
        } as DocsHostedDocument
      }
      host.documents = {
        capabilities: state.documentHost.capabilities,
        autoSave: state.documentHost.autoSave,
        initialDocument: () => document('initialDocument'),
        newDocument: () => document('newDocument'),
        openDocument: (id) => document('openDocument', [id ?? null]),
        recentDocuments: () =>
          request('services.call', { method: 'recentDocuments', args: [] }) as ReturnType<
            NonNullable<DocsDocumentHost['recentDocuments']>
          >,
        saveDocument: (value) =>
          request('services.call', { method: 'saveDocument', args: [value] }) as ReturnType<
            NonNullable<DocsDocumentHost['saveDocument']>
          >,
        createDocument: (value) =>
          request('services.call', { method: 'createDocument', args: [value] }) as ReturnType<
            NonNullable<DocsDocumentHost['createDocument']>
          >,
      }
    }
    attachments = state.attachments
    knowledge = state.knowledge
    const hasAttachments = attachments.length > 0
    if (cliAvailable === state.cliAvailable && attachmentsAvailable === hasAttachments && resources)
      return
    cliAvailable = state.cliAvailable
    attachmentsAvailable = hasAttachments
    resources = createDocsResourceTools({
      ...(attachmentsAvailable && {
        attachments: {
          list: () => attachments,
          read: (path, offset, maxChars) =>
            sourceRead(async () => {
              const result = await services.readAttachment(path, offset, maxChars)
              if (!result.ok) throw new Error(result.error ?? 'Attachment read failed')
              return {
                text: result.text ?? '',
                totalChars: result.totalChars ?? 0,
                offset: result.offset ?? 0,
              }
            }),
        },
      }),
      knowledge: {
        listFiles: () => knowledge,
        search: (q, limit) => sourceRead(() => services.kbSearch({ q, limit })),
        read: (path, offset) => sourceRead(() => services.kbRead(path, offset)),
      },
      ...(cliAvailable
        ? {
            runCli: async (args: string[], signal?: AbortSignal) => {
              if (signal?.aborted)
                return {
                  ok: false,
                  code: 1,
                  stdout: '',
                  stderr: 'Canceled before starting',
                  committed: false,
                }
              const operationId = crypto.randomUUID()
              const running = request('services.call', {
                method: 'runCli',
                args: [args, operationId],
              })
              const cancel = () => {
                void request('services.call', { method: 'cancelCli', args: [operationId] }).catch(
                  () => {},
                )
              }
              signal?.addEventListener('abort', cancel, { once: true })
              try {
                return (await running) as Awaited<ReturnType<DocsEditorServices['runCli']>>
              } finally {
                signal?.removeEventListener('abort', cancel)
              }
            },
          }
        : {}),
    })
  }
  listeners.set(
    'onResourceStateChanged',
    new Set([(...args) => updateResources(args[0] as Parameters<typeof updateResources>[0])]),
  )
  return {
    host,
    async refreshResources() {
      const state = await request('services.call', { method: 'getResourceState', args: [] })
      if (disposed) throw new Error('The Qt editor host was disposed')
      updateResources(state as Parameters<typeof updateResources>[0])
    },
    attach(endpoint: EditorEndpoint) {
      if (disposed) throw new Error('The Qt editor host was disposed')
      for (const off of editorDisposers) off()
      editorDisposers = []
      editor = endpoint
      notify('editor.ready', {})
      if (endpoint.ui)
        editorDisposers.push(
          endpoint.ui.ribbon.onAction((event) => {
            if (!disposed) notify('sdk.ribbonAction', event)
          }),
          endpoint.ui.fileMenu.onAction((event) => {
            if (!disposed) notify('sdk.menuAction', event)
          }),
        )
      if (endpoint.documents)
        editorDisposers.push(
          endpoint.documents.subscribe((state) => {
            if (!disposed) notify('sdk.documentStateChanged', { state })
          }),
        )
    },
    dispose() {
      if (disposed) return
      disposed = true
      for (const off of editorDisposers) off()
      editorDisposers = []
      editor = undefined
      notify('editor.disposed', {})
      channel.reply.disconnect(reply)
      channel.event.disconnect(event)
      channel.command.disconnect(command)
      listeners.clear()
      for (const operation of pending.values())
        operation.reject(new Error('The Qt editor host was disposed'))
      pending.clear()
    },
  }
}
