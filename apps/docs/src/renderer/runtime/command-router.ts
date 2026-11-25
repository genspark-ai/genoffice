import type { AgentToolCall } from '@genoffice/agent-core'
import type { McpCommandMessage } from '../../shared/ipc'
import type { DocsRuntime, DocsRuntimeClient, DocsRuntimeScope } from './docs-runtime'
import { DocsOperationError } from '../../shared/operation-result'

export interface DocsCommands {
  bindScope(bindingId: string, scope: DocsRuntimeScope, onRelease?: () => void): void
  releaseScope(bindingId: string): void
  /** End document authority before replacing the active generation. */
  releaseAll(): void
  handle(message: McpCommandMessage, signal?: AbortSignal): Promise<unknown>
  disconnect(callerId: string): void
  dispose(): void
}

/** Transport-neutral live-document endpoint. Identity is supplied by the trusted transport. */
export function createDocsCommandRouter(runtime: DocsRuntime): DocsCommands {
  const callers = new Map<
    string,
    { callerId: string; bindingId?: string; client: DocsRuntimeClient }
  >()
  const scopes = new Map<string, DocsRuntimeScope>()
  const releases = new Map<string, () => void>()
  const operations = new Map<
    string,
    { callerId: string; bindingId?: string; controller: AbortController }
  >()
  const disconnect = (callerId: string) => {
    for (const operation of operations.values())
      if (operation.callerId === callerId) operation.controller.abort()
    for (const [key, entry] of callers)
      if (entry.callerId === callerId) {
        entry.client.close()
        callers.delete(key)
      }
  }
  const releaseScope = (bindingId: string) => {
    for (const operation of operations.values())
      if (operation.bindingId === bindingId) operation.controller.abort()
    scopes.get(bindingId)?.close()
    scopes.delete(bindingId)
    for (const [key, entry] of callers)
      if (entry.bindingId === bindingId) {
        entry.client.close()
        callers.delete(key)
      }
    const released = releases.get(bindingId)
    releases.delete(bindingId)
    released?.()
  }
  const releaseAll = () => {
    for (const operation of operations.values()) operation.controller.abort()
    for (const { client } of callers.values()) client.close()
    callers.clear()
    for (const bindingId of [...scopes.keys()]) releaseScope(bindingId)
  }
  return {
    bindScope(bindingId: string, scope: DocsRuntimeScope, onRelease?: () => void) {
      if (scopes.has(bindingId)) throw new Error('The action scope is already registered')
      scopes.set(bindingId, scope)
      if (onRelease) releases.set(bindingId, onRelease)
    },
    releaseScope,
    releaseAll,
    async handle(message: McpCommandMessage, signal?: AbortSignal): Promise<unknown> {
      const callerId = message.callerId ?? 'legacy-mcp'
      if (message.command === 'close_caller') {
        disconnect(callerId)
        return { closed: true }
      }
      if (message.command === 'cancel_operation') {
        const requestId = (message.payload as { requestId?: unknown } | null)?.requestId
        const operation =
          typeof requestId === 'string'
            ? operations.get(JSON.stringify([callerId, requestId]))
            : undefined
        operation?.controller.abort()
        return { canceled: !!operation }
      }
      const operationKey = JSON.stringify([callerId, message.requestId])
      if (operations.has(operationKey))
        throw new Error('The document request identity is already in use')
      const controller = new AbortController()
      const operationSignal = signal
        ? AbortSignal.any([signal, controller.signal])
        : controller.signal
      operations.set(operationKey, { callerId, bindingId: message.bindingId, controller })
      try {
        const key = JSON.stringify([message.bindingId ?? null, callerId])
        let client = callers.get(key)?.client
        if (message.command === 'list_tools') {
          operationSignal.throwIfAborted()
          if (client) return client.availableTools()
          if (message.bindingId) return scopes.get(message.bindingId)?.availableTools() ?? []
          return runtime.availableTools()
        }
        if (!client) {
          if (message.bindingId) {
            const scope = scopes.get(message.bindingId)
            if (!scope) {
              const output = 'The action scope is not bound to this document'
              throw new DocsOperationError(output, {
                output,
                summary: 'Document operation refused',
                isError: true,
                errorCode: 'stale_target',
                committed: false,
                // No authorized document metadata remains; never inspect the active document.
                documentId: '',
                revision: 0,
              })
            }
            client = scope.connect({ id: callerId })
          } else {
            client = runtime.connect({ id: callerId })
            client.context()
          }
          callers.set(key, { callerId, bindingId: message.bindingId, client })
        }
        if (message.command === 'read_document') return client.context({ captureSelection: false })
        if (message.command === 'save_document') {
          const input = message.payload as { path?: unknown; overwrite?: unknown } | null
          if (!input || typeof input.path !== 'string' || !input.path)
            throw new Error('save_document requires an absolute path')
          const result = await client.save(
            { path: input.path, overwrite: input.overwrite === true },
            operationSignal,
          )
          if (!result.ok) throw new DocsOperationError(result.output, result)
          return result
        }
        const call: AgentToolCall =
          message.command === 'execute_tool'
            ? { ...(message.payload as AgentToolCall), id: message.requestId }
            : {
                id: message.requestId,
                name: message.command,
                input: message.payload as Record<string, unknown>,
              }
        const result = await client.execute(call, operationSignal)
        if (result.isError) throw new DocsOperationError(result.output, result)
        return result
      } finally {
        operations.delete(operationKey)
      }
    },
    disconnect,
    dispose() {
      releaseAll()
      operations.clear()
    },
  }
}
