/// Web-version bootstrap for the docs renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.desktop` / `window.projectApi` objects the preload exposes in the
/// desktop app, but backed by the HTTP/SSE transport against the running
/// Electron main process. Native-only channels (open/save dialogs, print,
/// clipboard, font metrics, window management) get browser equivalents so the
/// web version keeps the full feature surface. Inside Electron the preload has
/// already exposed the IPC-backed APIs and this module leaves them untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import {
  createWebFileBridge,
  downloadBytes,
  pickFileBytes,
  webCopyImage,
  webFontMetrics,
  webOpenTab,
  webPrint,
} from '@genoffice/ipc-bridge/web-native'
import { createDesktopApi, createProjectApi } from '../shared/desktop-api-factory'
import type { DesktopApi } from '../shared/ipc'
import {
  installDataflareEmbedBridge,
  getDataflareEmbedSessionId,
  postToEmbedParent,
  requestDataflareParent,
  requestDataflareStreamParent,
  type DataflareEmbedCommand,
  type DataflareOfficeContext,
} from '../shared/embed-bridge'

if (!isElectronRuntime()) {
  const embeddedPathPrefix = window.location.pathname.startsWith('/office-engine/') ? '/office-engine' : ''
  const transport = createHttpIpcTransport({ pathPrefix: embeddedPathPrefix })
  const files = createWebFileBridge(transport)
  let dataflareContext: DataflareOfficeContext | null = null
  let dataflareRevision = '0'
  const requestDataflare = (path: string, init: RequestInit = {}) => {
    if (window.parent === window) {
      const token = localStorage.getItem('Manager-Token')
      return fetch(path, {
        ...init,
        headers: { ...(init.headers || {}), ...(token ? { 'Manager-Token': token } : {}) },
      })
    }
    return requestDataflareParent({
      type: 'http-request',
      requestId: `df-http-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
      sessionId: getDataflareEmbedSessionId() || '',
      method: (init.method || 'GET').toUpperCase() as 'GET' | 'POST',
      path,
      jsonBody: typeof init.body === 'string' ? init.body : undefined,
    }).then((result) => new Response(result.body, { status: result.status, headers: result.headers }))
  }
  const uploadDataflare = (path: string, data: ArrayBuffer, fields: Record<string, string>) => {
    if (window.parent === window) {
      const token = localStorage.getItem('Manager-Token')
      const form = new FormData()
      form.append('file', new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), 'document.docx')
      Object.entries(fields).forEach(([key, value]) => form.append(key, value))
      return fetch(path, { method: 'POST', headers: token ? { 'Manager-Token': token } : undefined, body: form })
    }
    return requestDataflareParent({
      type: 'http-request',
      requestId: `df-http-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
      sessionId: getDataflareEmbedSessionId() || '',
      method: 'POST',
      path,
      file: { bytes: data, filename: 'document.docx', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
      fields,
    }).then((result) => new Response(result.body, { status: result.status, headers: result.headers }))
  }
  const bridgedWindow = window as unknown as Record<string, unknown>
  const desktopApi = (): DesktopApi => bridgedWindow.desktop as DesktopApi
  bridgedWindow.desktop = createDesktopApi(transport, {
    saveDocx: async (_path, data, _auto) => {
      const documentId = dataflareContext?.documentId
      if (!documentId || dataflareContext?.documentSource !== 'knowledge') {
        return await transport.invoke('docs:save', _path, data, _auto === true)
      }
      const response = await uploadDataflare(`/crmapi/knowledge/office/${encodeURIComponent(documentId)}`, data, { expectedRevision: dataflareRevision })
      const body = await response.json().catch(() => null) as {
        code?: number
        msg?: string
        data?: { revision?: string }
      } | null
      if (!response.ok || body?.code !== 0) {
        const error = body?.msg || `Dataflare document save failed (${response.status})`
        postToEmbedParent({ type: 'error', code: response.status === 409 || body?.code === 409 ? 'document-conflict' : 'save-failed', message: error })
        return { ok: false, reason: 'external-modified', error }
      }
      dataflareRevision = body.data?.revision || String(Number(dataflareRevision) + 1)
      postToEmbedParent({
        type: 'document-saved',
        documentId: documentId,
        revision: dataflareRevision,
      })
      return { ok: true }
    },
    aiTranslate: async (request) => {
      // Standalone web build: route through the local bridge transport so the
      // AI translation hits the web-server's `ai:translate` handler (the real
      // provider-backed translation path). The Dataflare branch only fires
      // when this window is embedded inside Dataflare with an active context.
      if (!dataflareContext) {
        return await transport.invoke('ai:translate', request)
      }
      const scope = request.range?.scope === 'document' ? 'document' : 'selection'
      const unitId = `${scope}-${Date.now().toString(36)}`
      const response = await requestDataflare('/office-engine/api/ai/translate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          requestId: unitId,
          idempotencyKey: unitId,
          documentId: dataflareContext?.documentId,
          documentType: 'docx',
          scene: scope,
          sourceLanguage: request.sourceLang || 'auto',
          targetLanguage: request.targetLang,
          preserveFormatting: request.preserveFormat !== false,
          memoryEnabled: true,
          qualityCheck: true,
          units: [{
            unitId,
            kind: scope === 'document' ? 'document' : 'paragraph',
            sourceText: request.instruction,
            order: 0,
            metadata: request.range ? { range: request.range } : undefined,
          }],
        }),
      })
      const body = (await response.json()) as {
        code?: number
        msg?: string
        data?: { requestId?: string; units?: Array<{ translatedText?: string; errorMessage?: string }> }
      }
      const unit = body.data?.units?.[0]
      if (!response.ok || body.code !== 0 || !unit?.translatedText) {
        return { ok: false, error: unit?.errorMessage || body.msg || 'Dataflare translation failed' }
      }
      return {
        ok: true,
        translated: unit.translatedText,
        planId: body.data?.requestId || unitId,
        sourceLang: request.sourceLang,
        targetLang: request.targetLang,
        preserveFormat: request.preserveFormat !== false,
      }
    },
    aiTranslateBatch: async (request: Parameters<NonNullable<import('../shared/desktop-api-factory').DesktopApiOverrides['aiTranslateBatch']>>[0]) => {
      if (!dataflareContext) {
        return await transport.invoke('ai:translate-batch', request)
      }
      const batchId = `document-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
      const response = await requestDataflare('/office-engine/api/ai/translate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requestId: batchId,
          idempotencyKey: batchId,
          documentId: dataflareContext?.documentId,
          documentType: 'docx',
          scene: request.scene || 'document',
          sourceLanguage: request.sourceLang || 'auto',
          targetLanguage: request.targetLang,
          preserveFormatting: request.preserveFormat !== false,
          memoryEnabled: true,
          qualityCheck: true,
          units: request.units.map((unit) => ({
            unitId: unit.unitId,
            kind: unit.kind,
            sourceText: unit.sourceText,
            order: unit.order,
            path: unit.path,
            metadata: { ...(unit.metadata || {}), range: unit.range || undefined },
          })),
        }),
      })
      const body = await response.json().catch(() => null) as {
        code?: number
        msg?: string
        data?: {
          units?: Array<{
            unitId?: string
            sourceText?: string
            translatedText?: string
            status?: string
            matchedTerms?: string[]
            warnings?: string[]
            errorMessage?: string
          }>
          quality?: { overallScore?: number; warnings?: string[] }
        }
      } | null
      const units = (body?.data?.units || []).map((unit) => ({
        unitId: unit.unitId || '',
        sourceText: unit.sourceText || '',
        translatedText: unit.translatedText,
        status: unit.status,
        matchedTerms: unit.matchedTerms,
        warnings: unit.warnings,
        errorMessage: unit.errorMessage,
        range: request.units.find((input) => input.unitId === unit.unitId)?.range || null,
      }))
      return {
        ok: response.ok && body?.code === 0 && units.length === request.units.length,
        units,
        quality: body?.data?.quality,
        error: body?.msg || (!response.ok ? `Dataflare translation failed (${response.status})` : undefined),
      }
    },
    aiTranslateBatchStream: async (request: Parameters<NonNullable<import('../shared/desktop-api-factory').DesktopApiOverrides['aiTranslateBatch']>>[0]) => {
      // SSE 流式批量翻译：每完成一个 unit 立即收到推送事件，
      // 通过 onUnit 回调让 GenOffice AI 面板实时追加翻译预览，
      // 替代旧的"等待整批返回"。
      if (window.parent === window || !getDataflareEmbedSessionId()) {
        // 独立模式：fallback 到同步批量接口
        return await desktopApi().aiTranslateBatch(request)
      }
      const streamUnits = new Map<string, NonNullable<NonNullable<Awaited<ReturnType<DesktopApi['aiTranslateBatch']>>>['units']>[number]>()
      let streamQuality: { overallScore?: number; warnings?: string[] } | undefined
      let streamError: string | undefined
      let streamOk = true
      const batchId = `stream-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
      await new Promise<void>((resolve) => {
        const unsubscribe = requestDataflareStreamParent(
          {
            type: 'http-stream-request',
            requestId: batchId,
            sessionId: getDataflareEmbedSessionId() || '',
            method: 'POST',
            path: '/office-engine/api/ai/translate/stream',
            jsonBody: JSON.stringify({
              requestId: batchId,
              idempotencyKey: batchId,
              documentId: dataflareContext?.documentId,
              documentType: 'docx',
              scene: request.scene || 'document',
              sourceLanguage: request.sourceLang || 'auto',
              targetLanguage: request.targetLang,
              preserveFormatting: request.preserveFormat !== false,
              memoryEnabled: true,
              qualityCheck: true,
              units: request.units.map((unit) => ({
                unitId: unit.unitId,
                kind: unit.kind,
                sourceText: unit.sourceText,
                order: unit.order,
                path: unit.path,
                metadata: { ...(unit.metadata || {}), range: unit.range || undefined },
              })),
            }),
          },
          (event) => {
            try {
              const payload = JSON.parse(event.data) as {
                type?: string
                status?: string
                unit?: {
                  unitId?: string
                  status?: string
                  sourceText?: string
                  translatedText?: string
                  matchedTerms?: string[]
                  warnings?: string[]
                  errorMessage?: string
                }
                quality?: { overallScore?: number; warnings?: string[] }
                message?: string
              }
              if (payload.type === 'unit' && payload.unit?.unitId) {
                const input = request.units.find((u) => u.unitId === payload.unit!.unitId)
                streamUnits.set(payload.unit.unitId, {
                  unitId: payload.unit.unitId,
                  sourceText: payload.unit.sourceText || '',
                  translatedText: payload.unit.translatedText,
                  status: payload.unit.status,
                  matchedTerms: payload.unit.matchedTerms,
                  warnings: payload.unit.warnings,
                  errorMessage: payload.unit.errorMessage,
                  range: input?.range || null,
                })
                postToEmbedParent({ type: 'ai-progress', status: 'running', progress: streamUnits.size / Math.max(request.units.length, 1) })
              } else if (payload.type === 'quality' && payload.quality) {
                streamQuality = payload.quality
              } else if (payload.type === 'complete') {
                if (payload.status && payload.status !== 'completed' && payload.status !== 'partial') {
                  streamOk = false
                  if (payload.status === 'failed') streamError = 'Dataflare translation completed with failed status'
                }
              } else if (payload.type === 'error') {
                streamOk = false
                streamError = payload.message || 'Dataflare stream error'
              }
            } catch (err) {
              console.warn('[web-bridge] failed to parse SSE event', err)
            }
          },
          (status) => {
            if (status >= 400) {
              streamOk = false
              streamError = `Dataflare stream failed (${status})`
            }
            unsubscribe()
            resolve()
          },
          (error) => {
            streamOk = false
            streamError = error.message
            unsubscribe()
            resolve()
          },
        )
      })
      const units = Array.from(streamUnits.values())
      return {
        ok: streamOk && units.length === request.units.length,
        units,
        quality: streamQuality,
        error: streamError,
      }
    },
    saveTranslationMemory: async (request) => {
      if (!dataflareContext) {
        return await transport.invoke('ai:save-translation-memory', request)
      }
      const requestId = `memory-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
      const response = await requestDataflare('/crmapi/ai/translation/v1/memory', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requestId,
          documentId: dataflareContext?.documentId,
          scene: request.scene,
          sourceLanguage: request.sourceLang,
          targetLanguage: request.targetLang,
          units: request.units,
        }),
      })
      const body = await response.json().catch(() => null) as {
        code?: number
        msg?: string
        data?: { savedCount?: number; skippedCount?: number }
      } | null
      if (!response.ok || body?.code !== 0) {
        return { ok: false, error: body?.msg || `Dataflare memory save failed (${response.status})` }
      }
      return { ok: true, savedCount: body.data?.savedCount, skippedCount: body.data?.skippedCount }
    },
    consumePendingOpenDocx: async () => {
      const path = new URLSearchParams(window.location.search).get('open')
      if (!path) return null
      return await transport.invoke('docs:open-path', path)
    },
    openDocx: async () => {
      const picked = await pickFileBytes(
        '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      )
      if (!picked) return null
      const { name, bytes } = picked[0]
      const path = await files.writeTempFile(name, bytes)
      return await transport.invoke('docs:open-path', path)
    },
    pickImage: async () => {
      const picked = await pickFileBytes('image/png,image/jpeg,image/gif')
      if (!picked) return null
      const { name, bytes } = picked[0]
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      const mime = IMAGE_MIME[ext]
      if (!mime) return null
      return { base64: bytesToBase64(bytes), mime, name }
    },
    saveDocxAs: async (defaultName, data) => {
      downloadBytes(defaultName, data)
      const path = await files.writeTempFile(defaultName, data)
      return { ok: true, path }
    },
    saveDocxNew: async (defaultName, data) => {
      downloadBytes(defaultName, data)
      const path = await files.writeTempFile(defaultName, data)
      return { ok: true, path }
    },
    print: async () => {
      webPrint()
      return { ok: true }
    },
    exportPdf: async () => {
      webPrint()
      return { ok: true, path: '' }
    },
    copyImageToClipboard: (dataUrl) => webCopyImage(dataUrl),
    fontMetrics: (family) => Promise.resolve(webFontMetrics(family)),
    pickAttachments: async () => {
      const picked = await pickFileBytes(undefined, true)
      if (!picked) return null
      const paths: string[] = []
      for (const file of picked) {
        paths.push(await files.writeTempFile(file.name, file.bytes))
      }
      return await transport.invoke('files:add', paths)
    },
    openNewTab: async (openPath) => {
      webOpenTab(openPath ? `#open=${encodeURIComponent(openPath)}` : window.location.href)
    },
    listDocsTabs: async () => [],
    focusDocsTab: async () => {},
    printPdfBuffer: async () => {
      // The browser cannot produce PDF bytes per print group; the merged save
      // below opens the browser print dialog (save as PDF) instead.
      return { ok: true, base64: '' }
    },
    saveMergedPdf: async () => {
      webPrint()
      return { ok: true, path: '' }
    },
  })
  bridgedWindow.projectApi = createProjectApi(transport)
  bridgedWindow.dataflareOfficeBridge = {
    postEvent: postToEmbedParent,
    isEmbedded: window.parent !== window,
    getRevision: () => dataflareRevision,
  }
  installDataflareEmbedBridge({
    onCommand: (command: DataflareEmbedCommand) => {
      if (command.type === 'init') dataflareContext = command.context
      window.dispatchEvent(new CustomEvent('dataflare:office-command', { detail: command }))
      if (command.type === 'init' && command.context.documentId && command.context.documentSource === 'knowledge') {
        void openDataflareKnowledgeDocument(command.context, transport, files, requestDataflare, (revision) => {
          dataflareRevision = revision
        })
      }
    },
    onGlobalState: (state, revision) => {
      // Update cached revision when host pushes a newer one (e.g. another tab/user saved).
      const incoming = state?.documentRevision
      if (incoming !== undefined && incoming !== null) {
        dataflareRevision = String(incoming)
      }
      window.dispatchEvent(new CustomEvent('dataflare:office-global-state', { detail: { state, revision } }))
    },
  })
}

async function openDataflareKnowledgeDocument(
  context: DataflareOfficeContext,
  transport: { invoke(channel: string, ...args: unknown[]): Promise<unknown> },
  files: { writeTempFile(name: string, bytes: ArrayBuffer): Promise<string> },
  requestDataflare: (path: string, init?: RequestInit) => Promise<Response>,
  onRevision: (revision: string) => void,
): Promise<void> {
  const documentId = context.documentId?.trim()
  if (!documentId || context.documentType !== 'docx') return
  try {
    const response = await requestDataflare(`/crmapi/knowledge/office/${encodeURIComponent(documentId)}`)
    if (!response.ok) throw new Error(`Dataflare document download failed (${response.status})`)
    onRevision(response.headers.get('X-Office-Revision') || response.headers.get('ETag')?.replace(/^"|"$/g, '') || '0')
    const bytes = await response.arrayBuffer()
    const name = `dataflare-${documentId}.docx`
    const path = await files.writeTempFile(name, bytes)
    await transport.invoke('docs:open-path', path)
  } catch (error) {
    postToEmbedParent({
      type: 'error',
      code: 'dataflare-document-open-failed',
      message: error instanceof Error ? error.message : String(error),
    })
  }
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

function bytesToBase64(bytes: ArrayBuffer): string {
  let binary = ''
  const view = new Uint8Array(bytes)
  const CHUNK = 0x8000
  for (let i = 0; i < view.length; i += CHUNK) {
    binary += String.fromCharCode(...view.subarray(i, i + CHUNK))
  }
  return btoa(binary)
}
