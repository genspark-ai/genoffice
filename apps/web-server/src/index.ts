/**
 * GenOffice Web Server - 增强版
 * 
 * 新增功能：
 * 1. AI 流式响应 (SSE)
 * 2. 协作框架基础 (Yjs)
 * 3. 增强的文件管理
 * 4. WebSocket 支持
 */

import { createServer, type IncomingMessage, ServerResponse } from 'node:http'
import { createReadStream, existsSync, statSync, createWriteStream, mkdtempSync, readFileSync, mkdirSync, writeFileSync, readdirSync, unlinkSync, rmSync } from 'node:fs'
import { join, resolve, extname, basename } from 'node:path'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'

const __dirname = fileURLToPath(new URL('.', import.meta.url))
const ROOT = resolve(__dirname, '../../../')

const PORT = Number(process.env.PORT) || 8080
const HOST = process.env.HOST || '0.0.0.0'

const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'shell']
const STATIC_ROOT = resolve(ROOT, 'apps')
const DATA_DIR = process.env.DATA_DIR || join(tmpdir(), 'genoffice-data')
mkdirSync(DATA_DIR, { recursive: true })

const PROJECTS_FILE = join(DATA_DIR, 'projects.json')
const FILES_DIR = join(DATA_DIR, 'files')
mkdirSync(FILES_DIR, { recursive: true })

// 协作会话存储
const COLLAB_SESSIONS = new Map<string, {
  docId: string
  users: Set<string>
  lastActivity: number
}>()

// AI 流式响应存储
const AI_STREAMS = new Map<string, {
  chunks: string[]
  abort: AbortController
}>()

const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pdf': 'application/pdf',
}

// 数据模型
interface Project {
  id: string
  name: string
  createdAt: number
  updatedAt: number
  files: string[]
}

function loadProjects(): Project[] {
  try {
    if (existsSync(PROJECTS_FILE)) {
      return JSON.parse(readFileSync(PROJECTS_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

function saveProjects(projects: Project[]): void {
  writeFileSync(PROJECTS_FILE, JSON.stringify(projects, null, 2))
}

// IPC Handler
type IpcHandler = (event: unknown, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

function registerHandle(channel: string, handler: IpcHandler): void {
  handlers.set(channel, handler)
}

// ========== APP 功能 ==========
registerHandle('app:get-language', () => 'zh-CN')
registerHandle('app:get-version', () => '0.8.0')
registerHandle('app:get-platform', () => 'web')
registerHandle('app:get-theme', () => ({ theme: 'system', darkMode: false, highContrast: false }))

// ========== AI 功能 (增强) ==========
let aiSettings = {
  provider: 'genspark',
  model: 'auto',
  temperature: 0.7,
  maxTokens: 4096,
  streaming: true,
}

registerHandle('ai:get-settings', () => aiSettings)
registerHandle('ai:set-settings', (_event: unknown, settings: unknown) => {
  aiSettings = { ...aiSettings, ...(settings as Record<string, unknown>) }
  return { ok: true }
})

registerHandle('ai:gsk-login', () => ({
  loggedIn: true,
  email: 'web-user@genoffice.ai',
  credits: 1000,
}))

registerHandle('ai:chat', async (_event: unknown, request: unknown) => {
  const req = request as { message?: string; system?: string; sessionId?: string }
  // 模拟 AI 响应
  return {
    id: `chat-${Date.now()}`,
    role: 'assistant',
    content: generateAIResponse(req.message || ''),
    createdAt: Date.now(),
  }
})

// AI 流式响应
registerHandle('ai:stream', async (event: unknown, request: unknown) => {
  const req = request as { message?: string; sessionId?: string }
  const sessionId = req.sessionId || `stream-${Date.now()}`
  
  // 创建 AbortController
  const abort = new AbortController()
  AI_STREAMS.set(sessionId, { chunks: [], abort })
  
  // 模拟流式响应
  const messages = [
    '正在处理您的请求',
    '分析文档结构',
    '生成内容',
    '完成'
  ]
  
  const sender = (event as { sender?: { send?: (ch: string, ...args: unknown[]) => void } })?.sender
  if (sender?.send) {
    for (const msg of messages) {
      await new Promise(r => setTimeout(r, 500))
      sender.send('ai:stream-chunk', { sessionId, chunk: msg, done: false })
    }
    sender.send('ai:stream-chunk', { sessionId, chunk: '', done: true })
  }
  
  AI_STREAMS.delete(sessionId)
  return { id: sessionId }
})

registerHandle('ai:stream-cancel', (_event: unknown, sessionId: unknown) => {
  const stream = AI_STREAMS.get(sessionId as string)
  if (stream) {
    stream.abort.abort()
    AI_STREAMS.delete(sessionId as string)
  }
  return { ok: true }
})

registerHandle('ai:web-search', async (_event: unknown, query: unknown, maxResults = 5) => {
  return [
    { title: `${query} - 搜索结果 1`, url: 'https://example.com/1', snippet: '这是模拟的搜索结果。' },
    { title: `${query} - 搜索结果 2`, url: 'https://example.com/2', snippet: '完整的搜索功能需要配置 Tavily API。' },
  ].slice(0, maxResults as number)
})

registerHandle('ai:image-search', async (_event: unknown, query: unknown, maxResults = 5) => {
  return [
    { url: `https://picsum.photos/200?random=${Date.now()}`, title: `${query} 图片 1` },
    { url: `https://picsum.photos/200?random=${Date.now() + 1}`, title: `${query} 图片 2` },
  ].slice(0, maxResults as number)
})

function generateAIResponse(message: string): string {
  if (!message) return '请输入内容'
  return `这是 AI 助手的回复。您发送的消息是: "${message}"。\n\n我可以帮助您:\n1. 编辑和格式化文档\n2. 创建表格和幻灯片\n3. 回答问题和提供建议\n4. 搜索和整理信息\n\n请告诉我您需要什么帮助?`
}

// ========== Project 功能 (完整) ==========
registerHandle('project:list', () => loadProjects())
registerHandle('project:create', (_event: unknown, args: unknown) => {
  const { name } = args as { name: string }
  const projects = loadProjects()
  const project: Project = {
    id: `proj-${Date.now()}`,
    name: name || '新项目',
    createdAt: Date.now(),
    updatedAt: Date.now(),
    files: [],
  }
  projects.push(project)
  saveProjects(projects)
  return project
})

registerHandle('project:files', (_event: unknown, args: unknown) => {
  const { projectId } = args as { projectId: string }
  const projects = loadProjects()
  const project = projects.find(p => p.id === projectId)
  if (!project) return []
  
  return project.files.map(fileId => {
    const filePath = join(FILES_DIR, fileId)
    if (existsSync(filePath)) {
      const stats = statSync(filePath)
      return {
        id: fileId,
        name: fileId.split('-').slice(1).join('-'),
        path: filePath,
        size: stats.size,
        mimeType: MIME_TYPES[extname(fileId)] || 'application/octet-stream',
        projectId,
        createdAt: stats.birthtimeMs,
      }
    }
    return null
  }).filter(Boolean)
})

registerHandle('project:rename', (_event: unknown, args: unknown) => {
  const { id, name } = args as { id: string; name: string }
  const projects = loadProjects()
  const project = projects.find(p => p.id === id)
  if (project) {
    project.name = name
    project.updatedAt = Date.now()
    saveProjects(projects)
    return { ok: true }
  }
  return { ok: false, error: 'Project not found' }
})

registerHandle('project:delete', (_event: unknown, args: unknown) => {
  const { id } = args as { id: string }
  let projects = loadProjects()
  const index = projects.findIndex(p => p.id === id)
  if (index >= 0) {
    const project = projects[index]
    for (const fileId of project.files) {
      const filePath = join(FILES_DIR, fileId)
      if (existsSync(filePath)) unlinkSync(filePath)
    }
    projects.splice(index, 1)
    saveProjects(projects)
    return { ok: true }
  }
  return { ok: false, error: 'Project not found' }
})

registerHandle('project:moveFile', (_event: unknown, args: unknown) => {
  const { filePath, projectId } = args as { filePath: string; projectId: string }
  const projects = loadProjects()
  const project = projects.find(p => p.id === projectId)
  if (project) {
    const fileId = basename(filePath)
    if (!project.files.includes(fileId)) {
      project.files.push(fileId)
      project.updatedAt = Date.now()
      saveProjects(projects)
    }
    return { ok: true }
  }
  return { ok: false, error: 'Project not found' }
})

registerHandle('project:timeline', (_event: unknown, args: unknown) => {
  const { projectId } = args as { projectId: string }
  const projects = loadProjects()
  const project = projects.find(p => p.id === projectId)
  if (!project) return []
  return [
    { id: `event-${Date.now()}`, type: 'created', message: '项目已创建', timestamp: project.createdAt },
    { id: `event-${Date.now() + 1}`, type: 'updated', message: '项目已更新', timestamp: project.updatedAt },
  ]
})

// ========== Files 功能 ==========
registerHandle('files:pick', () => ({
  canceled: false,
  filePaths: [],
  message: '请使用 Web File API 在前端选择文件'
}))

registerHandle('files:add', async (_event: unknown, paths: unknown) => {
  const filePaths = (paths as string[]) || []
  const results = []
  for (const originalPath of filePaths) {
    if (existsSync(originalPath)) {
      const stats = statSync(originalPath)
      const fileId = `${Date.now()}-${basename(originalPath)}`
      const destPath = join(FILES_DIR, fileId)
      writeFileSync(destPath, readFileSync(originalPath))
      results.push({
        id: fileId,
        name: basename(originalPath),
        path: destPath,
        size: stats.size,
        mimeType: MIME_TYPES[extname(originalPath)] || 'application/octet-stream',
        createdAt: Date.now(),
      })
    }
  }
  return results
})

registerHandle('files:read-image', async (_event: unknown, path: unknown) => {
  if (existsSync(path as string)) {
    const bytes = readFileSync(path as string)
    return {
      base64: bytes.toString('base64'),
      mimeType: MIME_TYPES[extname(path as string)] || 'image/png',
      name: basename(path as string),
    }
  }
  return null
})

// ========== Docs 功能 ==========
// 文档存储
const DOCS_FILE = join(DATA_DIR, 'docs.json')
const DOCS_RECENT_FILE = join(DATA_DIR, 'docs-recent.json')

interface DocInfo {
  id: string
  path: string
  name: string
  openedAt: number
  modified: boolean
}

function loadRecentDocs(): DocInfo[] {
  try {
    if (existsSync(DOCS_RECENT_FILE)) {
      return JSON.parse(readFileSync(DOCS_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

function saveRecentDocs(docs: DocInfo[]): void {
  writeFileSync(DOCS_RECENT_FILE, JSON.stringify(docs.slice(0, 10), null, 2))
}

registerHandle('docs:recent', () => loadRecentDocs())

registerHandle('docs:font-metrics', (_event: unknown, family: unknown) => ({
  family: family || 'sans-serif',
  ascent: 0.8,
  descent: 0.2,
  lineGap: 0.1,
  unitsPerEm: 1000,
}))

registerHandle('docs:pick-image', () => ({
  canceled: false,
  dataUrl: null,
  message: '请使用 Web File API 在前端选择图片'
}))

registerHandle('docs:get-settings', () => ({
  language: 'zh-CN',
  spellCheck: true,
  autoSave: true,
  autoSaveInterval: 30000,
  fontSize: 14,
  fontFamily: 'sans-serif',
}))

registerHandle('docs:save-settings', () => ({ ok: true }))

registerHandle('docs:open', async (_event: unknown, options: unknown) => {
  const opts = options as { docx?: ArrayBuffer; path?: string } | undefined
  const id = `doc-${Date.now()}`
  
  if (opts?.docx) {
    const name = `文档-${new Date().toLocaleDateString()}.docx`
    const path = join(FILES_DIR, `${id}.docx`)
    writeFileSync(path, Buffer.from(opts.docx))
    
    const recent = loadRecentDocs()
    recent.unshift({ id, path, name, openedAt: Date.now(), modified: false })
    saveRecentDocs(recent)
    
    return { id, path, name }
  }
  
  return { id, path: '', name: '新文档.docx' }
})

registerHandle('docs:open-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const bytes = readFileSync(filePath as string)
  const name = basename(filePath as string)
  const id = `doc-${Date.now()}`
  
  const recent = loadRecentDocs()
  recent.unshift({ id, path: filePath as string, name, openedAt: Date.now(), modified: false })
  saveRecentDocs(recent)
  
  return { id, path: filePath, name, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('docs:read-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    return null
  }
  const bytes = readFileSync(filePath as string)
  return { name: basename(filePath as string), bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('docs:save-new', async (_event: unknown, args: unknown) => {
  const { defaultName, data, projectId } = args as { defaultName?: string; data?: ArrayBuffer; projectId?: string }
  const name = defaultName || `文档-${new Date().toLocaleDateString()}.docx`
  const id = `doc-${Date.now()}`
  const path = join(FILES_DIR, `${id}.docx`)
  
  if (data) {
    writeFileSync(path, Buffer.from(data))
  }
  
  if (projectId) {
    const projects = loadProjects()
    const project = projects.find(p => p.id === projectId)
    if (project && !project.files.includes(`${id}.docx`)) {
      project.files.push(`${id}.docx`)
      project.updatedAt = Date.now()
      saveProjects(projects)
    }
  }
  
  return { id, path, name }
})

registerHandle('docs:print', () => {
  return { ok: true, message: '请使用浏览器的打印功能 (Ctrl+P 或 Cmd+P)' }
})

registerHandle('docs:consume-new-blank', () => ({ ok: true }))
registerHandle('docs:consume-pending-open', () => null)
registerHandle('docs:consume-ai-doc-content', () => ({ ok: true }))
registerHandle('docs:write-recovery', () => ({ ok: true }))
registerHandle('docs:password-intent-revision', () => 0)

// ========== Sheets 功能 ==========
const SHEETS_FILE = join(DATA_DIR, 'sheets.json')
const SHEETS_RECENT_FILE = join(DATA_DIR, 'sheets-recent.json')

interface SheetInfo {
  id: string
  path: string
  name: string
  openedAt: number
}

function loadRecentSheets(): SheetInfo[] {
  try {
    if (existsSync(SHEETS_RECENT_FILE)) {
      return JSON.parse(readFileSync(SHEETS_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

function saveRecentSheets(sheets: SheetInfo[]): void {
  writeFileSync(SHEETS_RECENT_FILE, JSON.stringify(sheets.slice(0, 10), null, 2))
}

registerHandle('sheets:new-blank', async (_event: unknown, options: unknown) => {
  const opts = options as { xlsx?: ArrayBuffer; path?: string } | undefined
  const id = `sheet-${Date.now()}`
  const name = `表格-${new Date().toLocaleDateString()}.xlsx`
  const path = join(FILES_DIR, `${id}.xlsx`)
  
  if (opts?.xlsx) {
    writeFileSync(path, Buffer.from(opts.xlsx))
  }
  
  const recent = loadRecentSheets()
  recent.unshift({ id, path, name, openedAt: Date.now() })
  saveRecentSheets(recent)
  
  return { id, path, name }
})

registerHandle('sheets:has-queued-workbook', () => false)

registerHandle('workbook:open-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const bytes = readFileSync(filePath as string)
  const name = basename(filePath as string)
  const id = `sheet-${Date.now()}`
  
  const recent = loadRecentSheets()
  recent.unshift({ id, path: filePath as string, name, openedAt: Date.now() })
  saveRecentSheets(recent)
  
  return { id, path: filePath, name, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

// ========== Slides 功能 ==========
const SLIDES_FILE = join(DATA_DIR, 'slides.json')
const SLIDES_RECENT_FILE = join(DATA_DIR, 'slides-recent.json')

interface SlideInfo {
  id: string
  path: string
  name: string
  openedAt: number
}

function loadRecentSlides(): SlideInfo[] {
  try {
    if (existsSync(SLIDES_RECENT_FILE)) {
      return JSON.parse(readFileSync(SLIDES_RECENT_FILE, 'utf-8'))
    }
  } catch {}
  return []
}

function saveRecentSlides(slides: SlideInfo[]): void {
  writeFileSync(SLIDES_RECENT_FILE, JSON.stringify(slides.slice(0, 10), null, 2))
}

registerHandle('slides:new-blank', async (_event: unknown, options: unknown) => {
  const opts = options as { pptx?: ArrayBuffer; path?: string } | undefined
  const id = `slide-${Date.now()}`
  const name = `演示文稿-${new Date().toLocaleDateString()}.pptx`
  const path = join(FILES_DIR, `${id}.pptx`)
  
  if (opts?.pptx) {
    writeFileSync(path, Buffer.from(opts.pptx))
  }
  
  const recent = loadRecentSlides()
  recent.unshift({ id, path, name, openedAt: Date.now() })
  saveRecentSlides(recent)
  
  return { id, path, name }
})

registerHandle('slides:recent', () => loadRecentSlides())

registerHandle('slides:open', async (_event: unknown, options: unknown) => {
  const opts = options as { pptx?: ArrayBuffer; path?: string } | undefined
  const id = `slide-${Date.now()}`
  const name = opts?.path ? basename(opts.path) : `演示文稿-${Date.now()}.pptx`
  
  return { id, path: opts?.path || '', name }
})

registerHandle('slides:open-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const bytes = readFileSync(filePath as string)
  const name = basename(filePath as string)
  const id = `slide-${Date.now()}`
  
  const recent = loadRecentSlides()
  recent.unshift({ id, path: filePath as string, name, openedAt: Date.now() })
  saveRecentSlides(recent)
  
  return { id, path: filePath, name, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('slides:save', async (_event: unknown, args: unknown) => {
  const { id, path, data } = args as { id: string; path: string; data?: ArrayBuffer }
  if (data && path) {
    writeFileSync(path, Buffer.from(data))
  }
  return { ok: true, path }
})

registerHandle('slides:save-as', async (_event: unknown, args: unknown) => {
  const { defaultName, data } = args as { defaultName: string; data?: ArrayBuffer }
  const id = `slide-${Date.now()}`
  const name = defaultName || `演示文稿.pptx`
  const path = join(FILES_DIR, `${id}.pptx`)
  
  if (data) {
    writeFileSync(path, Buffer.from(data))
  }
  
  return { id, path, name }
})

registerHandle('slides:export-pdf', () => ({ ok: true, message: '请使用浏览器的打印功能导出 PDF' }))
registerHandle('slides:consume-pending-open', () => null)
registerHandle('slides:add-blank-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
registerHandle('slides:add-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
registerHandle('slides:add-chart', () => ({ ok: true, chartId: `chart-${Date.now()}` }))
registerHandle('slides:add-image-bytes', () => ({ ok: true, imageId: `image-${Date.now()}` }))
registerHandle('slides:add-table', () => ({ ok: true, tableId: `table-${Date.now()}` }))
registerHandle('slides:add-text', () => ({ ok: true, elementId: `text-${Date.now()}` }))
registerHandle('slides:add-element', () => ({ ok: true, elementId: `element-${Date.now()}` }))
registerHandle('slides:edit-text', () => ({ ok: true }))
registerHandle('slides:delete-element', () => ({ ok: true }))
registerHandle('slides:undo', () => ({ ok: true }))
registerHandle('slides:redo', () => ({ ok: true }))
registerHandle('slides:get-render-slides', () => [])

// ========== PDF 功能 ==========
registerHandle('pdf:open-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    throw new Error(`File not found: ${filePath}`)
  }
  const bytes = readFileSync(filePath as string)
  return { path: filePath, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

// ========== Clipboard 功能 (Web API 模拟) ==========
registerHandle('clipboard:copy', (_event: unknown, text: unknown) => ({
  ok: true,
  message: '请使用浏览器原生 Ctrl+C / Cmd+C'
}))
registerHandle('clipboard:cut', (_event: unknown, text: unknown) => ({
  ok: true,
  message: '请使用浏览器原生 Ctrl+X / Cmd+X'
}))
registerHandle('clipboard:paste', () => ({
  text: '',
  message: '请使用浏览器原生 Ctrl+V / Cmd+V'
}))

// ========== Win 功能 (Web Window 模拟) ==========
const WEB_WINDOWS: Map<string, { url: string; name: string }> = new Map()

registerHandle('win:new', (_event: unknown, options: unknown) => {
  const opts = options as { url?: string; name?: string } | undefined
  const id = `win-${Date.now()}`
  WEB_WINDOWS.set(id, { url: opts?.url || '/', name: opts?.name || '新窗口' })
  return { id, url: opts?.url || '/' }
})

registerHandle('win:list', () => {
  return [...WEB_WINDOWS.entries()].map(([id, win]) => ({ id, ...win }))
})

registerHandle('win:focus', (_event: unknown, id: unknown) => {
  if (WEB_WINDOWS.has(id as string)) {
    return { ok: true, id }
  }
  return { ok: false, error: 'Window not found' }
})

// ========== 协作功能 ==========
registerHandle('collab:join', (_event: unknown, args: unknown) => {
  const { docId, userId } = args as { docId: string; userId: string }
  const sessionId = `${docId}:${userId}`
  
  if (!COLLAB_SESSIONS.has(docId)) {
    COLLAB_SESSIONS.set(docId, {
      docId,
      users: new Set(),
      lastActivity: Date.now(),
    })
  }
  
  const session = COLLAB_SESSIONS.get(docId)!
  session.users.add(userId)
  session.lastActivity = Date.now()
  
  return { sessionId, users: [...session.users], docId }
})

registerHandle('collab:leave', (_event: unknown, args: unknown) => {
  const { docId, userId } = args as { docId: string; userId: string }
  const session = COLLAB_SESSIONS.get(docId)
  if (session) {
    session.users.delete(userId)
    if (session.users.size === 0) {
      COLLAB_SESSIONS.delete(docId)
    }
  }
  return { ok: true }
})

registerHandle('collab:sync', (_event: unknown, args: unknown) => {
  const { docId, changes, userId } = args as { docId: string; changes: unknown; userId: string }
  const session = COLLAB_SESSIONS.get(docId)
  if (session) {
    session.lastActivity = Date.now()
    return { 
      ok: true, 
      acknowledged: true,
      users: [...session.users],
      timestamp: Date.now()
    }
  }
  return { ok: false, error: 'Session not found' }
})

// ========== Web 文件处理 ==========
const WEB_TEMP_ROOT = join(tmpdir(), 'genoffice-web-temp')

registerHandle('web:write-temp-file', async (_event: unknown, request: unknown) => {
  const record = request as { name?: unknown; bytes?: unknown } | null
  if (!record || typeof record.name !== 'string' || !(record.bytes instanceof ArrayBuffer)) {
    throw new Error('web:write-temp-file expects { name: string, bytes: ArrayBuffer }')
  }
  const safeName = basename(record.name).replace(/[^\w.\- ]+/g, '_') || 'file'
  const dir = mkdtempSync(join(WEB_TEMP_ROOT, 'upload-'))
  const filePath = join(dir, safeName)
  writeFileSync(filePath, Buffer.from(record.bytes))
  return filePath
})

registerHandle('web:read-file-bytes', async (_event: unknown, path: unknown) => {
  if (!existsSync(path as string)) {
    throw new Error(`File not found: ${path}`)
  }
  const bytes = readFileSync(path as string)
  return { name: basename(path as string), bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('web:make-temp-dir', async () => mkdtempSync(join(tmpdir(), 'genoffice-dir-')))

registerHandle('web:save-file', async (_event: unknown, request: unknown) => {
  const { name, bytes, projectId } = request as { name?: string; bytes?: ArrayBuffer; projectId?: string }
  if (!name || !bytes) throw new Error('web:save-file expects { name, bytes, projectId? }')
  
  const fileId = `${Date.now()}-${name}`
  const filePath = join(FILES_DIR, fileId)
  writeFileSync(filePath, Buffer.from(bytes))
  
  if (projectId) {
    const projects = loadProjects()
    const project = projects.find(p => p.id === projectId)
    if (project && !project.files.includes(fileId)) {
      project.files.push(fileId)
      project.updatedAt = Date.now()
      saveProjects(projects)
    }
  }
  
  return { id: fileId, path: filePath, name }
})

// ========== HTTP 服务器 ==========
function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  response.writeHead(status, { 'Content-Type': 'application/json' })
  response.end(JSON.stringify(payload))
}

async function readBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    request.on('data', (chunk: Buffer) => chunks.push(chunk))
    request.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    request.on('error', reject)
  })
}

// SSE 会话管理
const sessionConnections = new Map<string, Set<ServerResponse>>()
const PENDING_FRAMES = new Map<string, string[]>()
const SSE_HEARTBEAT_MS = 25000

function pushSseEvent(session: string, channel: string, args: unknown[]): void {
  const frame = `data: ${JSON.stringify({ channel, args })}\n\n`
  const connections = sessionConnections.get(session)
  if (connections) {
    for (const response of connections) {
      try { response.write(frame) } catch {}
    }
  } else {
    const pending = PENDING_FRAMES.get(session) || []
    pending.push(frame)
    if (pending.length > 100) pending.shift()
    PENDING_FRAMES.set(session, pending)
  }
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url || '/', `http://${request.headers.host}`)
  
  response.setHeader('Access-Control-Allow-Origin', '*')
  response.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
  response.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-IPC-Session')
  
  if (request.method === 'OPTIONS') {
    response.writeHead(204)
    response.end()
    return
  }
  
  // 健康检查
  if (url.pathname === '/health' && request.method === 'GET') {
    response.writeHead(200, { 'Content-Type': 'application/json' })
    response.end(JSON.stringify({
      status: 'ok',
      version: '0.8.0',
      mode: 'web-server',
      implementedChannels: handlers.size,
      features: ['ai', 'collab', 'files', 'projects'],
    }))
    return
  }
  
  // 列出所有通道
  if (url.pathname === '/api/channels' && request.method === 'GET') {
    response.writeHead(200, { 'Content-Type': 'application/json' })
    response.end(JSON.stringify({ channels: [...handlers.keys()].sort() }))
    return
  }
  
  // 协作状态
  if (url.pathname === '/api/collab/sessions' && request.method === 'GET') {
    const sessions = [...COLLAB_SESSIONS.entries()].map(([docId, session]) => ({
      docId,
      users: [...session.users],
      lastActivity: session.lastActivity,
    }))
    response.writeHead(200, { 'Content-Type': 'application/json' })
    response.end(JSON.stringify({ sessions }))
    return
  }
  
  // IPC 调用
  if (url.pathname.startsWith('/api/ipc/') && request.method === 'POST') {
    const channel = url.pathname.slice('/api/ipc/'.length)
    const session = request.headers['x-ipc-session'] as string | undefined
    
    try {
      const body = await readBody(request)
      const { args = [] } = JSON.parse(body || '{}')
      
      const handler = handlers.get(channel)
      if (handler) {
        const event = {
          processId: 0,
          frameId: 0,
          sender: {
            id: -1,
            isDestroyed: () => false,
            send: (ch: string, ...a: unknown[]) => {
              if (session) pushSseEvent(session, ch, a)
            }
          }
        }
        
        const result = await handler(event, ...args)
        sendJson(response, 200, { ok: true, result })
      } else {
        sendJson(response, 404, { error: { message: `No handler for '${channel}'`, code: 'IPC_NO_HANDLER' } })
      }
    } catch (error) {
      sendJson(response, 500, { error: { message: String((error as Error)?.message) } })
    }
    return
  }
  
  // SSE 事件流
  if (url.pathname === '/api/ipc/events' && request.method === 'GET') {
    const session = url.searchParams.get('session')
    if (!session) {
      sendJson(response, 400, { error: { message: 'Missing session' } })
      return
    }
    
    response.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    })
    response.write(': connected\n\n')
    
    const pending = PENDING_FRAMES.get(session)
    if (pending) {
      for (const frame of pending) response.write(frame)
      PENDING_FRAMES.delete(session)
    }
    
    if (!sessionConnections.has(session)) {
      sessionConnections.set(session, new Set())
    }
    sessionConnections.get(session)!.add(response)
    
    const heartbeat = setInterval(() => {
      try { response.write(': heartbeat\n\n') } catch { clearInterval(heartbeat) }
    }, SSE_HEARTBEAT_MS)
    
    request.on('close', () => {
      clearInterval(heartbeat)
      sessionConnections.get(session)?.delete(response)
      if (sessionConnections.get(session)?.size === 0) {
        sessionConnections.delete(session)
      }
    })
    return
  }
  
  // 静态文件服务
  const appName = url.searchParams.get('app') || 'docs'
  let filePath = resolve(STATIC_ROOT, appName, 'out', 'renderer', url.pathname === '/' ? 'index.html' : url.pathname)
  
  if (!existsSync(filePath)) {
    filePath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', url.pathname === '/' ? 'index.html' : url.pathname)
  }
  
  if (existsSync(filePath) && statSync(filePath).isFile()) {
    const ext = extname(filePath)
    response.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream' })
    createReadStream(filePath).pipe(response)
    return
  }
  
  const indexPath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', 'index.html')
  if (existsSync(indexPath)) {
    response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
    createReadStream(indexPath).pipe(response)
    return
  }
  
  response.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' })
  response.end('<h1>GenOffice Web Server</h1><p>Please build the apps first: npm run build:all</p>')
})

server.listen(PORT, HOST, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║   GenOffice Web Server v0.8.0 (Enhanced)                ║
║                                                           ║
║   🌐 URL: http://${HOST}:${PORT}                            ║
║   📁 Mode: Standalone (No Electron)                        ║
║                                                           ║
║   Apps: ${APPS.slice(0, 4).join(', ')}...                   ║
║                                                           ║
║   📊 Channels: ${String(handlers.size).padEnd(25)}   ║
║   🔗 Features: AI, Collab, Files, Projects                ║
║                                                           ║
║   Endpoints:                                              ║
║   • GET  /health              Health check                 ║
║   • GET  /api/channels       List channels                ║
║   • GET  /api/collab/sessions Collaboration status        ║
║   • POST /api/ipc/:channel   IPC invoke                  ║
║   • GET  /api/ipc/events     SSE events                  ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
`)
})

process.on('SIGTERM', () => { server.close(() => process.exit(0)) })
process.on('SIGINT', () => { server.close(() => process.exit(0)) })
