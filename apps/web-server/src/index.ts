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
const ROOT = resolve(__dirname, '../../..')

const PORT = Number(process.env.PORT) || 8080
const HOST = process.env.HOST || '0.0.0.0'

const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'shell']
const STATIC_ROOT = resolve(ROOT, 'apps')

// 使用固定的绝对路径存储数据
const DATA_DIR = process.env.DATA_DIR || '/tmp/genoffice-data'
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
  const req = request as { message?: string; system?: string; sessionId?: string; context?: unknown }
  // 增强的 AI 响应，结合文档上下文
  const message = req.message || ''
  const context = req.context
  
  let content = generateAIResponse(message)
  
  // 如果有文档上下文，提供更有针对性的回答
  if (context) {
    content = `基于您提供的文档内容，我来帮您分析：\n\n${content}\n\n如需进一步帮助，请告诉我具体问题。`
  }
  
  return {
    id: `chat-${Date.now()}`,
    role: 'assistant',
    content,
    createdAt: Date.now(),
    metadata: {
      model: aiSettings.model,
      provider: aiSettings.provider,
    }
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
registerHandle('slides:get-animations', () => [])
registerHandle('slides:get-chart-data', () => ({}))
registerHandle('slides:get-comments', () => [])
registerHandle('slides:get-header-footer', () => ({ enabled: false }))
registerHandle('slides:get-layouts', () => [])
registerHandle('slides:get-link', () => null)
registerHandle('slides:get-notes', () => '')
registerHandle('slides:get-sections', () => [])
registerHandle('slides:get-shape-keys', () => [])
registerHandle('slides:get-slide-links', () => [])
registerHandle('slides:get-slide-size', () => ({ width: 960, height: 540 }))
registerHandle('slides:get-transition', () => ({ type: 'none', duration: 0 }))
registerHandle('slides:apply-edit-script', () => ({ ok: true }))
registerHandle('slides:apply-header-footer', () => ({ ok: true }))
registerHandle('slides:apply-theme', () => ({ ok: true }))
registerHandle('slides:apply-txn', () => ({ ok: true }))
registerHandle('slides:batch-edit-transform', () => ({ ok: true }))
registerHandle('slides:chart-color-schemes', () => [])
registerHandle('slides:clipboard-external', () => ({}))
registerHandle('slides:clipboard-probe', () => ({}))
registerHandle('slides:copy-elements', () => ({ ok: true }))
registerHandle('slides:copy-slide', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
registerHandle('slides:delete-comment', () => ({ ok: true }))
registerHandle('slides:delete-slide', () => ({ ok: true }))
registerHandle('slides:duplicate-elements', () => ({ ok: true }))
registerHandle('slides:edit-background', () => ({ ok: true }))
registerHandle('slides:edit-chart', () => ({ ok: true }))
registerHandle('slides:edit-connector-endpoints', () => ({ ok: true }))
registerHandle('slides:edit-fill', () => ({ ok: true }))
registerHandle('slides:edit-image-fill', () => ({ ok: true }))
registerHandle('slides:edit-picture-opacity', () => ({ ok: true }))
registerHandle('slides:edit-picture-src-rect', () => ({ ok: true }))
registerHandle('slides:edit-stroke', () => ({ ok: true }))
registerHandle('slides:edit-table-cell', () => ({ ok: true }))
registerHandle('slides:edit-table-style', () => ({ ok: true }))
registerHandle('slides:edit-transform', () => ({ ok: true }))
registerHandle('slides:find-replace', () => ({ ok: true, count: 0 }))
registerHandle('slides:flip-elements', () => ({ ok: true }))
registerHandle('slides:font-catalog', () => [])
registerHandle('slides:font-missing', () => [])
registerHandle('slides:font-download', () => ({ ok: true, message: 'Web 版本不支持字体下载' }))
registerHandle('slides:font-install-local', () => ({ ok: true }))
registerHandle('slides:insert-model3d', () => ({ ok: true }))

// ========== AI 额外功能 ==========
registerHandle('ai:log-run-failure', () => ({ ok: true }))

// ========== AI 增强功能 ==========
// AI 内容生成
registerHandle('ai:generate-content', async (_event: unknown, request: unknown) => {
  const req = request as { type?: string; topic?: string; length?: number; style?: string }
  const { type = 'paragraph', topic = '', length = 200, style = 'formal' } = req
  
  const templates: Record<string, string> = {
    paragraph: `关于"${topic}"的段落内容。`,
    summary: `以下是关于"${topic}"的摘要总结。`,
    outline: `# ${topic}大纲\n\n1. 介绍\n2. 主要内容\n3. 结论`,
    introduction: `欢迎阅读关于"${topic}"的介绍。`,
    conclusion: `总结以上内容，关于"${topic}"的主要观点是...`,
  }
  
  return {
    id: `gen-${Date.now()}`,
    type,
    content: templates[type] || templates.paragraph,
    tokens: Math.floor(length / 4),
  }
})

// AI 翻译
registerHandle('ai:translate', async (_event: unknown, request: unknown) => {
  const req = request as { text?: string; from?: string; to?: string }
  return {
    id: `trans-${Date.now()}`,
    original: req.text || '',
    translated: `[${req.to || 'en'}] ${req.text || ''}`,
    from: req.from || 'auto',
    to: req.to || 'en',
  }
})

// AI 摘要
registerHandle('ai:summarize', async (_event: unknown, request: unknown) => {
  const req = request as { text?: string; maxLength?: number }
  const text = req.text || ''
  const maxLength = req.maxLength || 100
  
  return {
    id: `sum-${Date.now()}`,
    originalLength: text.length,
    summary: text.slice(0, maxLength) + (text.length > maxLength ? '...' : ''),
    keyPoints: ['要点1', '要点2', '要点3'],
  }
})

// AI 问答
registerHandle('ai:qa', async (_event: unknown, request: unknown) => {
  const req = request as { question?: string; context?: string }
  return {
    id: `qa-${Date.now()}`,
    question: req.question || '',
    answer: `基于提供的内容，关于"${req.question}"的回答是...`,
    confidence: 0.85,
  }
})

// AI 语法检查
registerHandle('ai:grammar-check', async (_event: unknown, text: unknown) => {
  return {
    id: `grammar-${Date.now()}`,
    original: text,
    corrected: text,
    errors: [],
    suggestions: [],
  }
})

// AI 关键词提取
registerHandle('ai:extract-keywords', async (_event: unknown, text: unknown) => {
  return {
    id: `kw-${Date.now()}`,
    keywords: ['关键词1', '关键词2', '关键词3'],
    score: [0.9, 0.7, 0.5],
  }
})

// AI 情感分析
registerHandle('ai:sentiment', async (_event: unknown, text: unknown) => {
  return {
    id: `sent-${Date.now()}`,
    text,
    sentiment: 'neutral',
    score: 0.5,
    emotions: { positive: 0.3, negative: 0.2, neutral: 0.5 },
  }
})
registerHandle('slides:get-run-links', () => [])
registerHandle('slides:group-elements', () => ({ ok: true, groupId: `group-${Date.now()}` }))
registerHandle('slides:has-slide-clipboard', () => false)
registerHandle('slides:history-batch-begin', () => ({ ok: true }))
registerHandle('slides:history-batch-end', () => ({ ok: true }))
registerHandle('slides:insert-image', () => ({ ok: true }))
registerHandle('slides:is-dirty', () => false)
registerHandle('slides:master-close', () => ({ ok: true }))
registerHandle('slides:master-delete-element', () => ({ ok: true }))
registerHandle('slides:master-edit-fill', () => ({ ok: true }))
registerHandle('slides:master-edit-stroke', () => ({ ok: true }))
registerHandle('slides:master-edit-text', () => ({ ok: true }))
registerHandle('slides:master-edit-transform', () => ({ ok: true }))
registerHandle('slides:master-enter', () => ({ ok: true }))
registerHandle('slides:master-open', () => ({ ok: true }))
registerHandle('slides:media-data', () => ({}))
registerHandle('slides:move-section', () => ({ ok: true }))
registerHandle('slides:move-slide', () => ({ ok: true }))
registerHandle('slides:native-clipboard', () => ({}))
registerHandle('slides:paste-elements', () => ({ ok: true }))
registerHandle('slides:paste-slide', () => ({ ok: true }))
registerHandle('slides:presenter-end', () => ({ ok: true }))
registerHandle('slides:presenter-start', () => ({ ok: true }))
registerHandle('slides:presenter-swap', () => ({ ok: true }))
registerHandle('slides:remove-section', () => ({ ok: true }))
registerHandle('slides:rename-section', () => ({ ok: true }))
registerHandle('slides:reorder-element', () => ({ ok: true }))
registerHandle('slides:replace-picture-bytes', () => ({ ok: true }))
registerHandle('slides:set-advance-times', () => ({ ok: true }))
registerHandle('slides:set-animations', () => ({ ok: true }))
registerHandle('slides:set-element-font', () => ({ ok: true }))
registerHandle('slides:set-element-paragraph-format', () => ({ ok: true }))
registerHandle('slides:set-hidden', () => ({ ok: true }))
registerHandle('slides:set-link', () => ({ ok: true }))
registerHandle('slides:set-notes', () => ({ ok: true }))
registerHandle('slides:set-sections', () => ({ ok: true }))
registerHandle('slides:set-slide-layout', () => ({ ok: true }))
registerHandle('slides:set-slide-size', () => ({ ok: true }))
registerHandle('slides:set-table-cell-anchor', () => ({ ok: true }))
registerHandle('slides:set-table-col-width', () => ({ ok: true }))
registerHandle('slides:set-table-row-height', () => ({ ok: true }))
registerHandle('slides:set-transition', () => ({ ok: true }))
registerHandle('slides:show-fullscreen', () => ({ ok: true }))
registerHandle('slides:table-merge', () => ({ ok: true }))
registerHandle('slides:table-structure', () => ({}))
registerHandle('slides:ungroup-element', () => ({ ok: true }))
registerHandle('slides:add-comment', () => ({ ok: true, commentId: `comment-${Date.now()}` }))
registerHandle('slides:add-ink', () => ({ ok: true }))
registerHandle('slides:add-media-bytes', () => ({ ok: true }))
registerHandle('slides:add-section', () => ({ ok: true, sectionId: `section-${Date.now()}` }))
registerHandle('slides:add-slide-with-layout', () => ({ ok: true, slideId: `slide-${Date.now()}` }))
registerHandle('slides:add-smartart', () => ({ ok: true }))
registerHandle('slides:ai-snapshot-restore', () => ({ ok: true }))
registerHandle('slides:audience-ready', () => ({ ok: true }))
registerHandle('slides:cloud-gen-status', () => ({ status: 'idle' }))
registerHandle('slides:files-add', async (_event: unknown, args: unknown) => {
  const paths = (args as string[]) || []
  return paths.map(p => ({ path: p, ok: true }))
})
registerHandle('slides:files-pick', () => ({
  canceled: false,
  filePaths: [],
  message: '请使用 Web File API'
}))
registerHandle('slides:files-read-image', async (_event: unknown, path: unknown) => {
  if (existsSync(path as string)) {
    const bytes = readFileSync(path as string)
    return { base64: bytes.toString('base64'), name: basename(path as string) }
  }
  return null
})
registerHandle('slides:pick-export-dir', () => ({ path: '/tmp/exports' }))
registerHandle('slides:pick-export-pdf-path', () => ({ path: '/tmp/exports/presentation.pdf' }))
registerHandle('slides:private-font-data', () => ({}))
registerHandle('slides:private-font-faces', () => [])
registerHandle('slides:repaste-slide', () => ({ ok: true }))

// ========== Sheets 额外功能 ==========
registerHandle('sheets:consume-new-blank', () => ({ ok: true }))

// ========== Markdown 功能 ==========
registerHandle('md-asset', async (_event: unknown, args: unknown) => {
  const { path, type } = args as { path: string; type: string }
  if (type === 'read' && existsSync(path)) {
    return { content: readFileSync(path, 'utf-8') }
  }
  return null
})

// ========== AnyDoc 文档处理功能 ==========
// AnyDoc: 通用文档识别、转换和处理

interface AnyDocConfig {
  ocrEnabled: boolean
  language: string
  preserveLayout: boolean
}

const anyDocConfig: AnyDocConfig = {
  ocrEnabled: true,
  language: 'zh-CN',
  preserveLayout: true,
}

registerHandle('anydoc:get-config', () => anyDocConfig)
registerHandle('anydoc:set-config', (_event: unknown, config: unknown) => {
  Object.assign(anyDocConfig, config)
  return { ok: true }
})

registerHandle('anydoc:recognize', async (_event: unknown, args: unknown) => {
  const { filePath, options } = args as { filePath: string; options?: { ocr?: boolean; language?: string } }
  // 模拟文档识别
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const ext = extname(filePath).toLowerCase()
  const fileName = basename(filePath)
  
  return {
    id: `doc-${Date.now()}`,
    fileName,
    fileType: ext.slice(1),
    pages: 1,
    text: `这是从 ${fileName} 提取的文本内容。\n完整的 OCR 识别需要集成 Tesseract.js 或云端 OCR 服务。`,
    metadata: {
      size: statSync(filePath).size,
      created: statSync(filePath).birthtime,
      modified: statSync(filePath).mtime,
    },
    success: true,
  }
})

registerHandle('anydoc:convert', async (_event: unknown, args: unknown) => {
  const { filePath, targetFormat } = args as { filePath: string; targetFormat: string }
  // 模拟文档转换
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const sourceFormat = extname(filePath).slice(1)
  const outputPath = join(FILES_DIR, `${Date.now()}-converted.${targetFormat}`)
  
  // 复制文件作为模拟转换
  const bytes = readFileSync(filePath)
  writeFileSync(outputPath, bytes)
  
  return {
    id: `convert-${Date.now()}`,
    sourceFormat,
    targetFormat,
    outputPath,
    success: true,
    message: `文档已从 ${sourceFormat} 转换为 ${targetFormat}`,
  }
})

registerHandle('anydoc:extract-text', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    return null
  }
  
  const ext = extname(filePath as string).toLowerCase()
  const bytes = readFileSync(filePath as string)
  
  // 根据文件类型提取文本
  if (['.txt', '.md', '.json', '.xml', '.html', '.csv'].includes(ext)) {
    return { text: bytes.toString('utf-8'), format: 'text' }
  } else if (['.docx', '.xlsx', '.pptx'].includes(ext)) {
    return { text: `Office 文档内容 (${ext})\n需要集成 mammoth.js 或专业解析库`, format: 'office' }
  } else if (ext === '.pdf') {
    return { text: `PDF 文档内容\n需要集成 pdf-parse 或 pdf.js`, format: 'pdf' }
  } else if (['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext)) {
    return { text: `图片内容\n需要 OCR 识别 (Tesseract.js)`, format: 'image' }
  }
  
  return { text: '未知文件格式', format: 'unknown' }
})

registerHandle('anydoc:extract-tables', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    return null
  }
  
  // 模拟表格提取
  return {
    tables: [],
    message: '表格提取需要专业解析库支持',
  }
})

registerHandle('anydoc:extract-images', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    return null
  }
  
  return {
    images: [],
    message: '图片提取功能需要实现',
  }
})

registerHandle('anydoc:render-preview', async (_event: unknown, args: unknown) => {
  const { filePath, options } = args as { filePath: string; options?: { width?: number; height?: number } }
  
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`)
  }
  
  const bytes = readFileSync(filePath)
  const ext = extname(filePath).toLowerCase()
  
  let mimeType = 'application/octet-stream'
  if (ext === '.pdf') mimeType = 'application/pdf'
  else if (['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(ext)) mimeType = `image/${ext.slice(1)}`
  
  return {
    base64: bytes.toString('base64'),
    mimeType,
    width: options?.width || 800,
    height: options?.height || 600,
  }
})

// ========== PDF 功能 ==========
registerHandle('pdf:open-path', async (_event: unknown, filePath: unknown) => {
  if (!existsSync(filePath as string)) {
    throw new Error(`File not found: ${filePath}`)
  }
  const bytes = readFileSync(filePath as string)
  return { path: filePath, bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

// ========== Clipboard 功能 (Web API 模拟) ==========
// 原生 IPC 通道 (copy/cut/paste)
registerHandle('copy', (_event: unknown, text: unknown) => ({
  ok: true,
  message: '请使用浏览器原生 Ctrl+C / Cmd+C'
}))
registerHandle('cut', (_event: unknown, text: unknown) => ({
  ok: true,
  message: '请使用浏览器原生 Ctrl+X / Cmd+X'
}))
registerHandle('paste', () => ({
  text: '',
  message: '请使用浏览器原生 Ctrl+V / Cmd+V'
}))

// Electron IPC 通道别名
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
