import { Filesystem, Directory } from '@capacitor/filesystem'
import { Share } from '@capacitor/share'
import { chatForProvider, defaultAiSettings } from '@genoffice/ai-provider'
import type { AiSettings, AiChatRequest, AiStreamRequest, GenSparkAccountStatus } from '@genoffice/ai-provider'
import type { DesktopApi, OpenFileResult, PickImageResult } from '../../docs/src/shared/ipc'

const SETTINGS_KEY = 'genoffice.android.ai.settings'
const RECENT_KEY = 'genoffice.android.recent-files'
const PATH_PREFIX = 'android-cache://'

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('')
}

async function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.click()
  })
}

async function fileToOpenResult(file: File): Promise<OpenFileResult> {
  const buffer = await file.arrayBuffer()
  const bytes = new Uint8Array(buffer)
  const id = crypto.randomUUID()
  const cachePath = `${id}.docx`
  await Filesystem.writeFile({ path: cachePath, directory: Directory.Cache, data: bytesToBase64(bytes), recursive: true })
  const path = `${PATH_PREFIX}${cachePath}`
  const result: OpenFileResult = {
    path,
    name: file.name,
    data: buffer,
    hash: await sha256(bytes),
  }
  rememberRecent(path)
  return result
}

function rememberRecent(path: string): void {
  try {
    const current = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]') as string[]
    localStorage.setItem(RECENT_KEY, JSON.stringify([path, ...current.filter((p) => p !== path)].slice(0, 20)))
  } catch { /* ignore */ }
}

function loadAiSettings(): AiSettings {
  const defaults = defaultAiSettings()
  try {
    const saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? 'null') as Partial<AiSettings> | null
    if (!saved) return defaults
    return { provider: saved.provider ?? defaults.provider, providers: { ...defaults.providers, ...(saved.providers ?? {}) } }
  } catch {
    return defaults
  }
}

async function saveBytesToDocuments(defaultName: string, data: ArrayBuffer): Promise<{ ok: boolean; path?: string; error?: string }> {
  try {
    const safeName = defaultName.replace(/[\\/:*?"<>|]+/g, '_') || 'Document.docx'
    const path = `GenOffice/${safeName}`
    await Filesystem.writeFile({ path, directory: Directory.Documents, data: bytesToBase64(new Uint8Array(data)), recursive: true })
    const uri = await Filesystem.getUri({ path, directory: Directory.Documents })
    try { await Share.share({ title: safeName, text: 'GenOffice document', url: uri.uri, dialogTitle: 'Share document' }) } catch { /* sharing is optional */ }
    rememberRecent(`${PATH_PREFIX}${path}`)
    return { ok: true, path: `${PATH_PREFIX}${path}` }
  } catch (error) {
    return { ok: false, error: String(error) }
  }
}

async function writePath(path: string, data: ArrayBuffer): Promise<{ ok: boolean; error?: string }> {
  try {
    if (!path.startsWith(PATH_PREFIX)) return { ok: false, error: 'Unsupported Android document path' }
    const relative = path.slice(PATH_PREFIX.length)
    const directory = relative.startsWith('GenOffice/') ? Directory.Documents : Directory.Cache
    const actual = directory === Directory.Documents ? relative : relative
    await Filesystem.writeFile({ path: actual, directory, data: bytesToBase64(new Uint8Array(data)), recursive: true })
    return { ok: true }
  } catch (error) {
    return { ok: false, error: String(error) }
  }
}

async function pickImage(): Promise<PickImageResult | null> {
  const file = await pickFile('image/png,image/jpeg,image/gif,image/webp')
  if (!file) return null
  const bytes = new Uint8Array(await file.arrayBuffer())
  const mime = file.type === 'image/jpeg' ? 'image/jpeg' : file.type === 'image/gif' ? 'image/gif' : 'image/png'
  return { base64: bytesToBase64(bytes), mime, name: file.name }
}

const api: DesktopApi = {
  getLanguage: async () => 'en',
  onLanguageChanged: () => () => {},
  openDocx: async () => {
    const file = await pickFile('.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    return file ? fileToOpenResult(file) : null
  },
  openDocxPath: async () => null,
  consumePendingOpenDocx: async () => null,
  consumeNewBlankDoc: async () => true,
  onOpenDocx: () => () => {},
  onRenamedDocx: () => () => {},
  saveDocx: async (path, data) => writePath(path, data),
  writeRecoveryCopy: async (_path, data) => {
    try {
      await Filesystem.writeFile({ path: `recovery-${Date.now()}.docx`, directory: Directory.Cache, data: bytesToBase64(new Uint8Array(data)), recursive: true })
      return { ok: true }
    } catch { return { ok: false } }
  },
  onTeardown: () => () => {},
  saveDocxAs: (defaultName, data) => saveBytesToDocuments(defaultName, data),
  saveDocxNew: (defaultName, data) => saveBytesToDocuments(defaultName, data),
  getRecentFiles: async () => {
    try { return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]') as string[] } catch { return [] }
  },
  pickImage,
  print: async () => { window.print() },
  exportPdf: async () => ({ ok: false, error: 'PDF export on Android is handled by the system print/share flow.' }),
  printPdfBuffer: async () => ({ ok: false, error: 'Direct PDF buffer export is not available on Android yet.' }),
  saveMergedPdf: async (defaultName, parts) => {
    try {
      const { PDFDocument } = await import('pdf-lib')
      const out = await PDFDocument.create()
      for (const part of parts) {
        const source = await PDFDocument.load(base64ToBytes(part))
        const pages = await out.copyPages(source, source.getPageIndices())
        pages.forEach((page) => out.addPage(page))
      }
      const bytes = await out.save()
      return saveBytesToDocuments(defaultName.endsWith('.pdf') ? defaultName : `${defaultName}.pdf`, bytes.buffer as ArrayBuffer)
    } catch (error) { return { ok: false, error: String(error) } }
  },
  getAiSettings: async () => loadAiSettings(),
  setAiSettings: async (settings) => localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)),
  aiChat: async (request: AiChatRequest) => chatForProvider(request.settings.provider, request.settings.providers[request.settings.provider], request.system, request.user),
  aiStream: async (request: AiStreamRequest) => {
    const result = await chatForProvider(request.settings.provider, request.settings.providers[request.settings.provider], request.system, request.messages.map((m) => typeof m.content === 'string' ? m.content : JSON.stringify(m.content)).join('\n'))
    const handler = streamHandlers.get(request.requestId)
    if (!handler) return
    if (result.ok) {
      handler({ requestId: request.requestId, type: 'delta', text: result.content ?? '' })
      handler({ requestId: request.requestId, type: 'done', stopReason: 'stop' })
    } else {
      handler({ requestId: request.requestId, type: 'error', error: result.error ?? 'AI request failed' })
    }
  },
  aiStreamCancel: async (requestId) => { streamHandlers.delete(requestId) },
  aiGskStatus: async (): Promise<GenSparkAccountStatus> => ({ loggedIn: false }),
  aiGskLogin: async () => {},
  webSearch: async () => ({ results: [], method: 'error', error: 'Web search is not configured on Android.' }),
  imageSearch: async () => ({ images: [], method: 'error', error: 'Image search is not configured on Android.' }),
  fetchImage: async (url) => {
    try {
      const response = await fetch(url)
      const blob = await response.blob()
      return { base64: bytesToBase64(new Uint8Array(await blob.arrayBuffer())), mime: blob.type || 'image/png' }
    } catch { return null }
  },
  pickAttachments: async () => null,
  addAttachmentPaths: async () => ({ accepted: [], rejected: ['Android path attachments are not enabled yet.'] }),
  addPastedImage: async () => ({ accepted: [], rejected: ['Pasted image attachments are not enabled yet.'] }),
  readAttachment: async () => ({ ok: false, error: 'Attachment text extraction is not enabled on Android yet.' }),
  readAttachmentImage: async () => ({ ok: false, error: 'Attachment images are not enabled on Android yet.' }),
  getPathForFile: () => '',
  openNewTab: async () => {},
  listDocsTabs: async () => [{ id: 'android-doc', title: 'Document', focused: true }],
  focusDocsTab: async () => {},
  onAiStream: (handler) => { streamHandlers.addHandler = handler; return () => { if (streamHandlers.addHandler === handler) streamHandlers.addHandler = undefined } },
  onMenuCommand: () => () => {},
  onCloseCheck: () => () => {},
  reportCloseCheck: () => {},
  onCloseSaveRequest: () => () => {},
  reportCloseSaveResult: () => {},
}

const streamHandlers: {
  addHandler?: (chunk: Parameters<NonNullable<DesktopApi['onAiStream']>>[0]) => void
  get: (id: string) => ((chunk: Parameters<NonNullable<DesktopApi['onAiStream']>>[0]) => void) | undefined
  set: (id: string, handler: (chunk: Parameters<NonNullable<DesktopApi['onAiStream']>>[0]) => void) => void
  delete: (id: string) => void
} = {
  addHandler: undefined,
  get: () => streamHandlers.addHandler,
  set: () => {},
  delete: () => {},
}

export function installAndroidDesktopApi(): void {
  window.desktop = api
  window.projectApi = {
    resolveChat: async () => ({ ok: false }),
    appendChat: async () => ({ ok: true }),
    loadChat: async () => ({ ok: true, messages: [] }),
    rebindChat: async () => ({ ok: true }),
    listProjects: async () => [],
    createProject: async () => ({ ok: false }),
    renameProject: async () => ({ ok: false }),
    deleteProject: async () => ({ ok: false }),
    moveFile: async () => ({ ok: false }),
    getTimeline: async () => [],
  } as never
}

export { api as androidDesktopApi }
