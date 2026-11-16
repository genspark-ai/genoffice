import type { DocsEditorApi, OpenFileResult } from '../../shared/ipc'
import {
  assertDocumentOperation,
  validateHostedIdentity,
  type DocsDocumentHost,
  type DocsHostedDocument,
} from '../../shared/document-host'

export function hostedOpenResult(document: DocsHostedDocument | null): OpenFileResult | null {
  if (!document) return null
  validateHostedIdentity(document)
  return {
    path: document.path ?? '',
    name: document.name,
    dataUrl: '',
    hash: document.fileHash ?? '',
    data: document.data,
    hosted: { id: document.id, name: document.name, readOnly: document.readOnly,
      ...(document.path ? {path:document.path,fileHash:document.fileHash} : {}) },
  }
}

/** A configured host is authoritative, including cancellations and unsupported operations. */
export function withDocumentServices(
  services: DocsEditorApi,
  host: DocsDocumentHost,
): DocsEditorApi {
  const open = async (id?: string) => {
    assertDocumentOperation(host, 'open')
    if (!host.openDocument) throw new Error('The host does not provide document opening')
    return hostedOpenResult(await host.openDocument(id))
  }
  return new Proxy(Object.create(null) as DocsEditorApi, {
    get(_target, key) {
      if (key === 'openDocx') return () => open()
      if (key === 'openDocxPath') return open
      if (key === 'consumePendingOpenDocx')
        return async () => hostedOpenResult((await host.initialDocument?.()) ?? null)
      if (key === 'consumeNewBlankDoc') return async () => false
      if (key === 'consumeAiDocContent' || key === 'consumeHeadlessExport') return async () => null
      if (key === 'getRecentFiles')
        return async () => ((await host.recentDocuments?.()) ?? []).map((doc) => doc.id)
      if (key === 'createDocument')
        return async (request: Parameters<DocsEditorApi['createDocument']>[0]) => {
          assertDocumentOperation(host, 'create')
          if (!host.createDocument) throw new Error('The host does not provide document creation')
          const result = await host.createDocument(request)
          if (result.ok && (result.committed !== true || result.pending))
            return {
              ...result,
              ok: false,
              error: result.error ?? 'The host did not confirm document creation',
            }
          return result
        }
      if (
        typeof key === 'string' &&
        ['saveDocx', 'saveDocxNew', 'saveDocxAs', 'saveDocxTo', 'writeRecoveryCopy'].includes(key)
      )
        return () => {
          throw new Error(
            'Use the hosted document lifecycle; local file persistence is unavailable',
          )
        }
      const value: unknown = Reflect.get(services, key, services)
      if (
        typeof key === 'string' &&
        (key.startsWith('export') ||
          ['pickExportImagesTarget', 'saveMergedPdf', 'writeExportImage'].includes(key)) &&
        typeof value === 'function'
      )
        return (...args: unknown[]) => {
          assertDocumentOperation(host, 'export')
          return value.apply(services, args)
        }
      return typeof value === 'function' ? value.bind(services) : value
    },
  })
}
