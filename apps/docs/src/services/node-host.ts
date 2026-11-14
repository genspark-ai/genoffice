import { createInterface } from 'node:readline'
import { Console } from 'node:console'
import { mkdir } from 'node:fs/promises'
import { randomUUID } from 'node:crypto'
import { z } from 'zod'
import { createDocsFileAccess, type DocsFileAccess } from './file-access'
import { createDocsResourceServices, type DocsResourceServices } from './resource-services'
import { createDocsDocumentCreation } from './document-creation'
import { createDocsExportServices, type DocsExportServices } from './export-services'
import { createDocsMailMergeServices, type DocsMailMergeServices } from './mail-merge-services'
import { decodeDataUrl } from '@genoffice/electron-utils/image-data'
import { MAX_COPY_IMAGE_DATA_URL, validCopyImageDataUrl } from '../main/copy-image-guard'
import { DocsFileError } from './file-errors'
import { configureMetricsCache, familyVerticalMetrics } from '@genoffice/font-metrics'
import {
  fetchRemoteImage,
  readBodyCapped,
  MAX_REMOTE_IMAGE_BYTES,
} from '@genoffice/electron-utils/remote-image'
import { createDocsFileSession, type DocsFileSession } from './document-files'
import { createDocsFileServices, type DocsFileServices } from './file-services'
import { createDocsAssetServer, type DocsAssetServer } from './asset-server'
import { createDocsHostBridge, type DocsHostBridge } from './host-bridge'
import { createDocsHttpMcpServer, type DocsHttpMcpServer } from './docs-http-mcp'
import type { DocsCommandEndpoint } from './docs-mcp'
import { readPickedImage } from './picked-image'
import { readAuthorizedImage } from './image-source'
import {
  DocsOperationError,
  isDocsOperationResult,
  isDocsSaveOutcome,
} from '../shared/operation-result'
import type { AiDocContent, CreateDocumentRequest } from '../shared/ipc'
import type { McpCommandResult } from '../shared/ipc'

// This owned server reserves stdout for JSON-RPC, including when parsers log diagnostics.
globalThis.console = new Console({ stdout: process.stderr, stderr: process.stderr })

// Stdout is native-host IPC; Agent traffic uses standard MCP transports.
const messageSchema = z
  .object({
    jsonrpc: z.literal('2.0'),
    id: z.union([z.string(), z.number(), z.null()]).optional(),
    method: z.string().optional(),
    params: z.unknown().optional(),
    result: z.unknown().optional(),
    error: z
      .object({ code: z.number().int(), message: z.string(), data: z.unknown().optional() })
      .optional(),
  })
  .refine((message) =>
    message.method !== undefined
      ? !('result' in message) && !('error' in message)
      : message.id !== undefined && 'result' in message !== 'error' in message,
  )
const initSchema = z
  .object({
    stateDirectory: z.string(),
    readRoots: z.array(z.string()),
    writeRoots: z.array(z.string()),
    defaultSaveDirectory: z.string(),
    initialPath: z.string().optional(),
    initialContent: z
      .object({
        title: z.string(),
        html: z.string(),
        targetPath: z.string(),
        creationId: z.string().uuid(),
      })
      .strict()
      .optional(),
    attachments: z.array(z.string()).max(200).optional(),
    knowledge: z.array(z.string()).max(200).optional(),
    cli: z.object({ executable: z.string(), entry: z.string() }).strict().optional(),
  })
  .strict()
const callSchema = z.object({ method: z.string(), args: z.array(z.unknown()) }).strict()
const bufferSchema = z
  .object({ base64: z.string().max(Math.ceil((512 * 1024 * 1024) / 3) * 4) })
  .strict()
const optionalTextSchema = z
  .string()
  .nullish()
  .transform((value) => value ?? undefined)
const optionalNumberSchema = z
  .number()
  .nullish()
  .transform((value) => value ?? undefined)
const exportBase64Limit = Math.ceil((128 * 1024 * 1024) / 3) * 4
const exportLinksSchema = z
  .array(
    z
      .object({
        page: z.number().int().nonnegative(),
        x: z.number(),
        y: z.number(),
        w: z.number(),
        h: z.number(),
        toPage: z.number().int().nonnegative(),
        toY: z.number(),
      })
      .strict(),
  )
  .max(20_000)
  .nullish()
  .transform((value) => value ?? undefined)
const markdownImagesSchema = z
  .array(
    z
      .object({
        name: z.string().max(255),
        src: z.string().max(Math.ceil((20 * 1024 * 1024) / 3) * 4 + 1024),
      })
      .strict(),
  )
  .max(1000)
const pending = new Map<string, { resolve(value: unknown): void; reject(error: Error): void }>()
const bridges = new Map<string, DocsHostBridge>()
const creatingBindings = new Map<string, { revoked: boolean }>()
const cliOperations = new Map<string, AbortController>()
let files: DocsFileSession | undefined
let services: DocsFileServices | undefined
let assets: DocsAssetServer | undefined
let access: DocsFileAccess | undefined
let resources: DocsResourceServices | undefined
let creation: ReturnType<typeof createDocsDocumentCreation> | undefined
let exportServices: DocsExportServices | undefined
let mailMerge: DocsMailMergeServices | undefined
let hostCli: { executable: string; entry: string } | undefined
let initialized = false
let ready = false
let closing = false
let pendingPath: string | undefined
let newBlank = false
let initialContent: AiDocContent | null = null
let activePath: string | undefined
const recent: string[] = []
let httpMcp:
  { bindingId: string; service: DocsHttpMcpServer; readOnlyService: DocsHttpMcpServer } | undefined
let startingMcp:
  { bindingId: string; controller: AbortController; promise: Promise<unknown> } | undefined
let mcpStatus: 'stopped' | 'failed' = 'stopped'

const send = (value: unknown) => {
  if (!closing) process.stdout.write(JSON.stringify(value) + '\n')
}
const notify = (method: string, params: unknown) => send({ jsonrpc: '2.0', method, params })
function nativeRequest(method: string, params: unknown, signal?: AbortSignal): Promise<unknown> {
  return new Promise((resolve, reject) => {
    if (closing || signal?.aborted) {
      reject(new Error('The native host request was canceled'))
      return
    }
    const id = randomUUID()
    const abort = () => {
      pending.delete(id)
      notify('host.cancel', { id })
      reject(new Error('The native host request was canceled'))
    }
    pending.set(id, {
      resolve: (value) => {
        signal?.removeEventListener('abort', abort)
        resolve(value)
      },
      reject: (error) => {
        signal?.removeEventListener('abort', abort)
        reject(error)
      },
    })
    signal?.addEventListener('abort', abort, { once: true })
    send({ jsonrpc: '2.0', id, method, params })
  })
}

const documentEndpoint: DocsCommandEndpoint = {
  async request(message, signal) {
    const cancel = () => notify('editor.cancel', { message })
    signal.addEventListener('abort', cancel, { once: true })
    try {
      const response = (await nativeRequest(
        'editor.request',
        { message },
        signal,
      )) as McpCommandResult
      if (!response.ok) {
        if (isDocsOperationResult(response.result))
          throw new DocsOperationError(response.error ?? response.result.output, response.result)
        throw new Error(response.error ?? 'The editor command failed')
      }
      return response.result
    } finally {
      signal.removeEventListener('abort', cancel)
    }
  },
  closeCaller: (callerId) => notify('editor.disconnect', { callerId }),
}
const mcpState = () =>
  httpMcp
    ? {
        status: 'running',
        connection: httpMcp.service.connection,
        readOnlyConnection: httpMcp.readOnlyService.connection,
      }
    : { status: mcpStatus }
const reportMcpState = () =>
  notify('host.mcpStateChanged', {
    status: httpMcp ? 'running' : mcpStatus,
    ...(httpMcp ? { documentId: httpMcp.service.connection.documentId } : {}),
  })

async function stopHttpMcp(bindingId?: string) {
  const starting = startingMcp
  const active = httpMcp
  if (bindingId && starting?.bindingId !== bindingId && active?.bindingId !== bindingId) return
  startingMcp = undefined
  httpMcp = undefined
  mcpStatus = 'stopped'
  if (starting) {
    starting.controller.abort()
    notify('editor.releaseScope', { bindingId: starting.bindingId })
  }
  const stopped = active
    ? Promise.all([active.readOnlyService.close(), active.service.close()])
    : undefined
  reportMcpState()
  await Promise.all([stopped, starting?.promise.catch(() => undefined)])
}

function startHttpMcp(): Promise<unknown> {
  if (!ready || closing) return Promise.reject(new Error('The document backend is not ready'))
  if (httpMcp) return Promise.resolve(mcpState())
  if (startingMcp) return startingMcp.promise
  const attempt = {
    bindingId: randomUUID(),
    controller: new AbortController(),
    promise: Promise.resolve<unknown>(undefined),
  }
  startingMcp = attempt
  attempt.promise = (async () => {
    const binding = z
      .object({ bindingId: z.literal(attempt.bindingId), documentId: z.string().uuid() })
      .parse(
        await nativeRequest(
          'editor.bindMcp',
          { bindingId: attempt.bindingId },
          attempt.controller.signal,
        ),
      )
    const service = await createDocsHttpMcpServer({
      version: '0.1.0',
      endpoint: documentEndpoint,
      ...binding,
      releaseScope: (bindingId) => notify('editor.releaseScope', { bindingId }),
    })
    if (closing || startingMcp !== attempt || attempt.controller.signal.aborted) {
      await service.close()
      throw new Error('MCP access was stopped while starting')
    }
    let readOnlyService: DocsHttpMcpServer
    try {
      readOnlyService = await createDocsHttpMcpServer({
        version: '0.1.0',
        endpoint: documentEndpoint,
        ...binding,
        readOnly: true,
        releaseScope: () => {},
      })
    } catch (error) {
      await service.close()
      throw error
    }
    if (closing || startingMcp !== attempt || attempt.controller.signal.aborted) {
      await Promise.all([readOnlyService.close(), service.close()])
      throw new Error('MCP access was stopped while starting')
    }
    httpMcp = { bindingId: binding.bindingId, service, readOnlyService }
    reportMcpState()
    return mcpState()
  })()
    .catch((error: unknown) => {
      notify('editor.releaseScope', { bindingId: attempt.bindingId })
      if (startingMcp === attempt) {
        mcpStatus = 'failed'
        reportMcpState()
      }
      throw error
    })
    .finally(() => {
      if (startingMcp === attempt) startingMcp = undefined
    })
  return attempt.promise
}

function bytes(value: unknown): ArrayBuffer {
  const { base64 } = bufferSchema.parse(value)
  if (base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(base64))
    throw new Error('Invalid document bytes')
  const data = Buffer.from(base64, 'base64')
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
}

async function callService(method: string, args: unknown[]): Promise<unknown> {
  if (!ready || !files || !services || !exportServices || closing)
    throw new Error('The document backend is not initialized')
  const text = (index: number) => z.string().parse(args[index])
  const source = (index: number) => z.string().nullable().optional().parse(args[index])
  switch (method) {
    case 'consumePendingOpenDocx': {
      const path = pendingPath
      pendingPath = undefined
      return path ? services.openDocxPath(path) : null
    }
    case 'consumeNewBlankDoc': {
      const blank = newBlank
      newBlank = false
      return blank
    }
    case 'consumeAiDocContent': {
      const content = initialContent
      initialContent = null
      return content
    }
    case 'consumeHeadlessExport':
      return null
    case 'openDocx':
      return services.openDocx()
    case 'openDocxPath':
      return services.openDocxPath(text(0))
    case 'openDocxDecrypt':
      return services.openDocxDecrypt(text(0), text(1))
    case 'saveDocx':
      return services.saveDocx(text(0), bytes(args[1]), z.boolean().optional().parse(args[2]))
    case 'saveDocxAs':
      return services.saveDocxAs(text(0), bytes(args[1]), source(2))
    case 'saveDocxNew':
      return services.saveDocxNew(text(0), bytes(args[1]))
    case 'saveDocxTo':
      return services.saveDocxTo(text(0), bytes(args[1]), z.boolean().parse(args[2]), source(3))
    case 'writeRecoveryCopy':
      return services.writeRecoveryCopy(text(0), bytes(args[1]))
    case 'setDocPassword':
      return services.setDocPassword(
        z.string().nullable().parse(args[0]),
        z.string().nullable().parse(args[1]),
      )
    case 'docPasswordIntentRevision':
      return services.docPasswordIntentRevision()
    case 'discardDocPasswordIntents':
      return services.discardDocPasswordIntents(z.number().int().nonnegative().parse(args[0]))
    case 'getRecentFiles':
      return services.getRecentFiles()
    case 'exportPdf':
      return exportServices.exportPdf(
        text(0),
        z.number().parse(args[1]),
        z.number().parse(args[2]),
        optionalTextSchema.parse(args[3]),
        optionalNumberSchema.parse(args[4]),
        exportLinksSchema.parse(args[5]),
      )
    case 'exportHtml':
      return exportServices.exportHtml(
        text(0),
        z
          .string()
          .max(128 * 1024 * 1024)
          .parse(args[1]),
        optionalTextSchema.parse(args[2]),
      )
    case 'exportMarkdown':
      return exportServices.exportMarkdown(
        text(0),
        z
          .string()
          .max(128 * 1024 * 1024)
          .parse(args[1]),
        markdownImagesSchema.parse(args[2]),
        optionalTextSchema.parse(args[3]),
      )
    case 'printPdfBuffer':
      return exportServices.printPdfBuffer(
        z.number().parse(args[0]),
        z.number().parse(args[1]),
        optionalNumberSchema.parse(args[2]),
      )
    case 'saveMergedPdf':
      return exportServices.saveMergedPdf(
        text(0),
        z.array(z.string().max(exportBase64Limit)).max(1000).parse(args[1]),
        optionalTextSchema.parse(args[2]),
        exportLinksSchema.parse(args[3]),
      )
    case 'pickExportImagesTarget':
      return exportServices.pickExportImagesTarget()
    case 'takeExportPdf':
      return exportServices.takeExportPdf(text(0))
    case 'writeExportImage':
      return exportServices.writeExportImage(
        text(0),
        z.string().max(255).parse(args[1]),
        z.string().max(exportBase64Limit).parse(args[2]),
      )
    case 'saveImageAs':
      return exportServices.saveImageAs(
        z
          .string()
          .max(Math.ceil(MAX_REMOTE_IMAGE_BYTES / 3) * 4 + 1024)
          .parse(args[0]),
      )
    case 'pickMergeDataSource':
      return mailMerge!.pickMergeDataSource()
    case 'writeMergeOutput': {
      const request = z
        .object({
          mode: z.enum(['single', 'labels', 'perRecord', 'envelopes']),
          files: z
            .array(
              z
                .object({
                  name: z.string().max(1000),
                  data: z.object({ base64: z.string().max(exportBase64Limit) }).strict(),
                })
                .strict(),
            )
            .max(1000),
        })
        .strict()
        .refine(
          (request) =>
            request.files.reduce((total, file) => total + file.data.base64.length, 0) <=
            Math.ceil((512 * 1024 * 1024) / 3) * 4,
          'Merge output exceeds the size limit',
        )
        .parse(args[0])
      return mailMerge!.writeMergeOutput({
        ...request,
        files: request.files.map((file) => ({ name: file.name, data: bytes(file.data) })),
      })
    }
    case 'readClipboardImage': {
      const src = z.string().max(MAX_COPY_IMAGE_DATA_URL).parse(args[0])
      if (!validCopyImageDataUrl(src))
        throw new DocsFileError('Invalid clipboard image URL', 'invalid_input')
      const data = /^data:image\//i.test(src) ? decodeDataUrl(src) : null
      const image = data ? { body: data.bytes, mime: data.mime } : await files.readAsset(src)
      if (
        !image ||
        !/^image\/(png|jpe?g|gif|webp|bmp)$/i.test(image.mime ?? '') ||
        image.body.length === 0 ||
        image.body.length > 48 * 1024 * 1024
      )
        throw new DocsFileError('Clipboard image is not an owned image asset', 'permission_denied')
      return { base64: Buffer.from(image.body).toString('base64'), mime: image.mime }
    }
    case 'createDocument':
      if (!creation) throw new Error('The document creation service is not initialized')
      return creation.create(args[0] as CreateDocumentRequest)
    case 'getResourceState':
      return {
        attachments: resources?.listAttachments() ?? [],
        knowledge: resources?.kbList() ?? [],
        cliAvailable: !!hostCli,
      }
    case 'readAttachment':
      return resources!.readAttachment(
        text(0),
        z.number().parse(args[1]),
        z.number().parse(args[2]),
      )
    case 'kbList':
      return resources!.kbList()
    case 'kbRead':
      return resources!.kbRead(text(0), z.number().parse(args[1]))
    case 'kbSearch':
      return resources!.kbSearch(
        z.object({ q: z.string(), limit: z.number().optional() }).strict().parse(args[0]),
      )
    case 'runCli': {
      const argv = z.array(z.string()).parse(args[0])
      const operationId = z.string().uuid().nullish().parse(args[1])
      if (!operationId) return resources!.runCli(argv)
      if (cliOperations.has(operationId))
        throw new DocsFileError('The CLI operation is already running', 'invalid_input')
      const controller = new AbortController()
      cliOperations.set(operationId, controller)
      try {
        return await resources!.runCli(argv, controller.signal)
      } finally {
        cliOperations.delete(operationId)
      }
    }
    case 'cancelCli': {
      const controller = cliOperations.get(z.string().uuid().parse(args[0]))
      controller?.abort()
      return { canceled: !!controller }
    }
    case 'pickImage': {
      const path = z
        .string()
        .nullable()
        .parse(await nativeRequest('host.dialog', { kind: 'image' }))
      return path ? readPickedImage(path) : null
    }
    case 'readImageSource': {
      if (!access) throw new Error('File access is not initialized')
      const source = z.object({ path: z.string() }).strict().parse(args[0])
      return readAuthorizedImage(access, source.path)
    }
    case 'fontMetrics':
      return familyVerticalMetrics(text(0))
    case 'fetchImage': {
      const url = text(0)
      if (!/^https?:\/\//.test(url))
        throw new Error('Only direct http(s) image URLs are available to this host')
      const response = await fetchRemoteImage(url)
      if (!response?.ok) return null
      const data = await readBodyCapped(response, MAX_REMOTE_IMAGE_BYTES)
      return {
        base64: Buffer.from(data).toString('base64'),
        mime: response.headers.get('content-type')?.split(';')[0] ?? 'application/octet-stream',
      }
    }
    default:
      throw new Error(`The host service is not available: ${method}`)
  }
}

async function dispatch(method: string, params: unknown): Promise<unknown> {
  if (method === 'initialize') {
    if (initialized || closing) throw new Error('The native host is already initialized or closed')
    const options = initSchema.parse(params)
    initialized = true
    const createdFiles = await createDocsFileSession(options)
    if (closing) {
      createdFiles.dispose()
      throw new Error('The native host closed during initialization')
    }
    files = createdFiles
    const createdAccess = await createDocsFileAccess(options)
    if (closing) {
      createdAccess.dispose()
      throw new Error('The native host closed during initialization')
    }
    access = createdAccess
    const createdExports = await createDocsExportServices({
      access: createdAccess,
      files: createdFiles,
      stateDirectory: options.stateDirectory,
      dialogs: {
        pickSave: async (selection) =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'saveExport', ...selection })),
        pickDirectory: async () =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'directory' })),
      },
      printPdf: async (geometry) =>
        new Uint8Array(bytes(await nativeRequest('host.printDocument', geometry))),
    })
    if (closing) {
      await createdExports.dispose()
      throw new Error('The native host closed during initialization')
    }
    exportServices = createdExports
    mailMerge = createDocsMailMergeServices({
      access: createdAccess,
      materialize: (data) => createdFiles.materializeDocx(data),
      dialogs: {
        pickSource: async () =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'mergeSource' })),
        pickSave: async (selection) =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'saveExport', ...selection })),
        pickDirectory: async () =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'directory' })),
      },
    })
    const createdResources = await createDocsResourceServices({
      ...options,
      openFiles: () => (activePath ? [activePath] : []),
    })
    if (closing) {
      createdResources.dispose()
      throw new Error('The native host closed during initialization')
    }
    resources = createdResources
    await resources.setAttachments(options.attachments ?? [])
    await resources.setKnowledgeFiles(options.knowledge ?? [])
    if (closing) throw new Error('The native host closed during initialization')
    hostCli = options.cli
    await mkdir(options.stateDirectory, { recursive: true })
    if (closing) throw new Error('The native host closed during initialization')
    configureMetricsCache(options.stateDirectory)
    pendingPath = options.initialPath
    activePath = options.initialPath
    initialContent = options.initialContent ?? null
    newBlank = !pendingPath
    services = createDocsFileServices({
      files,
      defaultSaveDirectory: options.defaultSaveDirectory,
      dialogs: {
        pickOpen: async () =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'open' })),
        pickSave: async (options) =>
          z
            .string()
            .nullable()
            .parse(await nativeRequest('host.dialog', { kind: 'save', ...options })),
        chooseRecovery: async (options) =>
          z
            .enum(['restore', 'discard'])
            .parse(await nativeRequest('host.dialog', { kind: 'recovery', ...options })),
        confirmExternalOverwrite: async (path) =>
          z.boolean().parse(await nativeRequest('host.dialog', { kind: 'overwrite', path })),
      },
      recentFiles: {
        list: async () => [...recent],
        add: async (path) => {
          const existing = recent.indexOf(path)
          if (existing >= 0) recent.splice(existing, 1)
          recent.unshift(path)
          recent.splice(30)
        },
      },
      onWarning: () => process.stderr.write('A document bookkeeping operation failed.\n'),
    })
    creation = createDocsDocumentCreation({
      access,
      directory: options.defaultSaveDirectory,
      createWord: async (content) => {
        const outcome = await nativeRequest('host.createWord', {
          ...content,
          creationId: randomUUID(),
        })
        if (!isDocsSaveOutcome(outcome)) throw new Error('Invalid new Word save acknowledgement')
        return outcome
      },
      printHtml: async (html) =>
        new Uint8Array(bytes(await nativeRequest('host.renderHtml', { html }))),
      openFile: async (path) => z.boolean().parse(await nativeRequest('host.openFile', { path })),
    })
    const createdAssets = await createDocsAssetServer(files)
    if (closing) {
      await createdAssets.close()
      throw new Error('The native host closed during initialization')
    }
    assets = createdAssets
    ready = true
    return { assets: { url: assets.url, token: assets.token }, nodeVersion: process.versions.node }
  }
  if (method === 'services.call') {
    const call = callSchema.parse(params)
    return callService(call.method, call.args)
  }
  if (method === 'files.grant') {
    if (!access || closing) throw new Error('File access is not initialized')
    const grant = z
      .object({ path: z.string(), read: z.boolean(), write: z.boolean() })
      .strict()
      .parse(params)
    await access.grantFile(grant.path, grant)
    await resources!.grantFile(grant.path, grant)
    return { ok: true }
  }
  if (method === 'files.revoke') {
    if (!access || closing) throw new Error('File access is not initialized')
    const { path } = z.object({ path: z.string() }).strict().parse(params)
    await access.revokeFile(path)
    await resources!.revokeFile(path)
    return { ok: true }
  }
  if (method === 'document.hintPath') {
    const { path } = z.object({ path: z.string().nullable() }).strict().parse(params)
    activePath = path ?? undefined
    return { ok: true }
  }
  if (method === 'resources.configure') {
    if (!resources || closing) throw new Error('Resources are not initialized')
    const selection = z
      .object({
        attachments: z.array(z.string()).max(200),
        knowledge: z.array(z.string()).max(200),
      })
      .strict()
      .parse(params)
    await resources.setAttachments(selection.attachments)
    await resources.setKnowledgeFiles(selection.knowledge)
    const state = {
      attachments: resources.listAttachments(),
      knowledge: resources.kbList(),
      cliAvailable: !!hostCli,
    }
    notify('host.resourcesChanged', state)
    return state
  }
  if (method === 'mcp.start') return startHttpMcp()
  if (method === 'mcp.state') return mcpState()
  if (method === 'mcp.stop' || method === 'mcp.revoke') {
    const bindingId =
      method === 'mcp.revoke'
        ? z.object({ bindingId: z.string().uuid() }).strict().parse(params).bindingId
        : undefined
    await stopHttpMcp(bindingId)
    return mcpState()
  }
  if (method === 'document.bindAction') {
    if (!ready || !files || closing) throw new Error('The document backend is not initialized')
    const action = z.object({ bindingId: z.string().uuid() }).passthrough().parse(params)
    if (bridges.has(action.bindingId) || creatingBindings.has(action.bindingId))
      throw new Error('The document action is already bound')
    const creation = { revoked: false }
    creatingBindings.set(action.bindingId, creation)
    try {
      const bridge = await createDocsHostBridge({
        bindingId: action.bindingId,
        releaseScope: (bindingId) => notify('editor.releaseScope', { bindingId }),
        endpoint: documentEndpoint,
      })
      if (closing || creation.revoked) {
        await bridge.close()
        throw new Error('The document action was closed or revoked during creation')
      }
      bridges.set(action.bindingId, bridge)
      return { url: bridge.url, token: bridge.token }
    } finally {
      creatingBindings.delete(action.bindingId)
    }
  }
  if (method === 'document.releaseAction') {
    const { bindingId } = z.object({ bindingId: z.string() }).strict().parse(params)
    const bridge = bridges.get(bindingId)
    const creating = creatingBindings.get(bindingId)
    if (creating) creating.revoked = true
    bridges.delete(bindingId)
    await bridge?.close()
    return { released: !!bridge || !!creating }
  }
  if (method === 'document.releaseAllActions') {
    const grants = [...bridges.values()]
    const creating = [...creatingBindings.values()].filter((binding) => !binding.revoked)
    // Revoke pending grants before asynchronous transport shutdown can publish them.
    for (const binding of creating) binding.revoked = true
    bridges.clear()
    await Promise.all([stopHttpMcp(), ...grants.map((bridge) => bridge.close())])
    return { released: grants.length + creating.length }
  }
  throw new Error('Method not found')
}

async function shutdown() {
  if (closing) return
  closing = true
  ready = false
  for (const controller of cliOperations.values()) controller.abort()
  resources?.dispose()
  const exportDisposal = exportServices?.dispose()
  for (const binding of creatingBindings.values()) binding.revoked = true
  for (const operation of pending.values())
    operation.reject(new Error('The native host disconnected'))
  pending.clear()
  try {
    await Promise.all([stopHttpMcp(), ...[...bridges.values()].map((bridge) => bridge.close())])
  } finally {
    bridges.clear()
    try {
      await exportDisposal
    } finally {
      access?.dispose()
      try {
        await assets?.close()
      } finally {
        files?.dispose()
      }
    }
  }
}

const input = createInterface({ input: process.stdin, crlfDelay: Infinity })
const rpcError = (code: number, message: string) => ({
  jsonrpc: '2.0',
  id: null,
  error: { code, message },
})
async function handleMessage(value: unknown): Promise<unknown> {
  if (Array.isArray(value)) {
    if (!value.length) return rpcError(-32600, 'Invalid Request')
    const replies = (await Promise.all(value.map(handleMessage))).filter(
      (reply) => reply !== undefined,
    )
    return replies.length ? replies : undefined
  }
  const parsed = messageSchema.safeParse(value)
  if (!parsed.success) return rpcError(-32600, 'Invalid Request')
  const message = parsed.data
  if (!message.method) {
    const request = pending.get(String(message.id))
    if (!request) return
    pending.delete(String(message.id))
    if (message.error) request.reject(new Error(message.error.message))
    else request.resolve(message.result)
    return
  }
  try {
    const result = await dispatch(message.method, message.params)
    if (message.id !== undefined) return { jsonrpc: '2.0', id: message.id, result: result ?? null }
  } catch (error) {
    if (message.id !== undefined)
      return {
        jsonrpc: '2.0',
        id: message.id,
        error: {
          code:
            error instanceof z.ZodError
              ? -32602
              : error instanceof Error && error.message === 'Method not found'
                ? -32601
                : -32000,
          message:
            error instanceof z.ZodError
              ? 'Invalid params'
              : error instanceof Error
                ? error.message
                : 'The native host operation failed',
          ...(error instanceof DocsFileError ? { data: { errorCode: error.errorCode } } : {}),
        },
      }
  }
}
input.on('line', (line) => {
  let value: unknown
  try {
    value = JSON.parse(line)
  } catch {
    send(rpcError(-32700, 'Parse error'))
    return
  }
  void handleMessage(value).then((reply) => {
    if (reply !== undefined) send(reply)
  })
})
input.once('close', () => {
  void shutdown().catch(() => {
    process.exitCode = 1
  })
})
process.once('SIGTERM', () => {
  input.close()
})
