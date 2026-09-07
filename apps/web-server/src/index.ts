/**
 * GenOffice Web Server
 * 
 * 独立的 Web Server 版本，不需要 Electron。
 * 提供：
 * 1. 静态文件服务
 * 2. HTTP IPC Bridge（模拟 Electron IPC）
 * 3. 应用路由
 * 
 * 功能列表（对标 Electron IPC）：
 * - 基础应用 (app:*)
 * - AI 功能 (ai:*)
 * - 文档功能 (docs:*)
 * - 项目管理 (project:*)
 * - 文件管理 (files:*)
 * - 表格功能 (sheets:*)
 * - 幻灯片功能 (slides:*)
 */

import { createServer, type IncomingMessage, ServerResponse } from 'node:http'
import { createReadStream, existsSync, statSync, createWriteStream, mkdtempSync, readFileSync, mkdirSync, writeFileSync, readdirSync, unlinkSync } from 'node:fs'
import { join, resolve, extname, basename } from 'node:path'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'

const __dirname = fileURLToPath(new URL('.', import.meta.url))
// ROOT is project root: apps/web-server/dist -> apps/web-server -> apps -> project root
const ROOT = resolve(__dirname, '../../../')

// 默认端口，可通过环境变量覆盖
const PORT = Number(process.env.PORT) || 8080
const HOST = process.env.HOST || '0.0.0.0'

// 静态文件根目录（各个应用的 renderer 输出）
const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'shell']
const STATIC_ROOT = resolve(ROOT, 'apps')

// 数据存储目录
const DATA_DIR = process.env.DATA_DIR || join(tmpdir(), 'genoffice-data')
mkdirSync(DATA_DIR, { recursive: true })

// 项目存储
const PROJECTS_FILE = join(DATA_DIR, 'projects.json')
const FILES_DIR = join(DATA_DIR, 'files')
mkdirSync(FILES_DIR, { recursive: true })

// MIME 类型映射
const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript',
  '.mjs': 'application/javascript',
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
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.wasm': 'application/wasm',
  '.map': 'application/json',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pdf': 'application/pdf',
}

// ==================== 数据存储 ====================

interface Project {
  id: string
  name: string
  createdAt: number
  updatedAt: number
  files: string[]
}

interface FileEntry {
  id: string
  name: string
  path: string
  size: number
  mimeType: string
  projectId?: string
  createdAt: number
}

// 加载/保存项目数据
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

// ==================== IPC Handler ====================

type IpcHandler = (event: unknown, ...args: unknown[]) => unknown

const handlers = new Map<string, IpcHandler>()

function registerHandle(channel: string, handler: IpcHandler): void {
  handlers.set(channel, handler)
}

// ==================== APP 功能 ====================

registerHandle('app:get-language', () => {
  const lang = process.env.LANG || process.env.LC_ALL || 'zh-CN'
  return lang.split('_')[0].toLowerCase() + '-' + lang.split('_')[1]?.split('.')[0] || 'CN'
})

registerHandle('app:get-version', () => '0.8.0')

registerHandle('app:get-platform', () => 'web')

registerHandle('app:get-theme', () => ({
  theme: 'system',
  darkMode: false,
  highContrast: false,
}))

// ==================== AI 功能 ====================

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
  // 模拟 AI 聊天响应
  const req = request as { message?: string; system?: string }
  return {
    id: `chat-${Date.now()}`,
    role: 'assistant',
    content: `这是 Web 模式的 AI 助手回复。您发送的消息是: "${req.message || 'Hello'}"。完整的 AI 功能需要后端服务支持。`,
    createdAt: Date.now(),
  }
})

registerHandle('ai:stream', async (event: unknown, request: unknown) => {
  const req = request as { message?: string }
  const chunks = [
    { content: '这是', done: false },
    { content: ' Web 模式的', done: false },
    { content: ' AI 流式响应。', done: false },
    { content: '完整的 AI 流式功能需要后端服务支持。', done: true },
  ]
  return { id: `stream-${Date.now()}`, chunks }
})

registerHandle('ai:web-search', async (_event: unknown, query: unknown, maxResults = 5) => {
  // 模拟搜索结果
  return [
    { title: `${query} - 搜索结果 1`, url: 'https://example.com/1', snippet: '这是模拟的搜索结果。' },
    { title: `${query} - 搜索结果 2`, url: 'https://example.com/2', snippet: '完整的搜索功能需要配置 Tavily API。' },
  ].slice(0, maxResults as number)
})

registerHandle('ai:image-search', async (_event: unknown, query: unknown, maxResults = 5) => {
  return [
    { url: `https://picsum.photos/200?random=1`, title: `${query} 图片 1` },
    { url: `https://picsum.photos/200?random=2`, title: `${query} 图片 2` },
  ].slice(0, maxResults as number)
})

// ==================== PROJECT 功能 ====================

registerHandle('project:list', () => {
  return loadProjects()
})

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
  
  const files: FileEntry[] = []
  for (const fileId of project.files) {
    const filePath = join(FILES_DIR, fileId)
    if (existsSync(filePath)) {
      const stats = statSync(filePath)
      files.push({
        id: fileId,
        name: fileId,
        path: filePath,
        size: stats.size,
        mimeType: MIME_TYPES[extname(fileId)] || 'application/octet-stream',
        projectId,
        createdAt: stats.birthtimeMs,
      })
    }
  }
  return files
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
  const projectIndex = projects.findIndex(p => p.id === id)
  if (projectIndex >= 0) {
    // 删除项目文件
    for (const fileId of projects[projectIndex].files) {
      const filePath = join(FILES_DIR, fileId)
      if (existsSync(filePath)) {
        unlinkSync(filePath)
      }
    }
    projects.splice(projectIndex, 1)
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

// ==================== FILES 功能 ====================

registerHandle('files:pick', async (_event: unknown, options: unknown) => {
  // Web 模式下返回模拟文件选择
  // 实际文件选择需要前端使用 Web File API
  return {
    canceled: false,
    filePaths: [],
    message: '请使用 Web File API 在前端选择文件'
  }
})

registerHandle('files:add', async (_event: unknown, paths: unknown) => {
  const filePaths = (paths as string[]) || []
  const results = []
  
  for (const originalPath of filePaths) {
    if (existsSync(originalPath)) {
      const stats = statSync(originalPath)
      const fileId = `${Date.now()}-${basename(originalPath)}`
      const destPath = join(FILES_DIR, fileId)
      
      // 复制文件
      const content = readFileSync(originalPath)
      writeFileSync(destPath, content)
      
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
  const filePath = path as string
  if (existsSync(filePath)) {
    const bytes = readFileSync(filePath)
    return {
      base64: bytes.toString('base64'),
      mimeType: MIME_TYPES[extname(filePath)] || 'image/png',
      name: basename(filePath),
    }
  }
  return null
})

// ==================== DOCS 功能 ====================

registerHandle('docs:recent', () => {
  return []
})

registerHandle('docs:font-metrics', (_event: unknown, family: unknown) => {
  // 返回模拟字体度量
  return {
    family: family || 'sans-serif',
    ascent: 0.8,
    descent: 0.2,
    lineGap: 0.1,
    unitsPerEm: 1000,
  }
})

registerHandle('docs:pick-image', () => {
  return {
    canceled: false,
    dataUrl: null,
    message: '请使用 Web File API 在前端选择图片'
  }
})

// ==================== DOCS 设置 ====================

registerHandle('docs:get-settings', () => ({
  language: 'zh-CN',
  spellCheck: true,
  autoSave: true,
  autoSaveInterval: 30000,
  fontSize: 14,
  fontFamily: 'sans-serif',
}))

registerHandle('docs:save-settings', (_event: unknown, settings: unknown) => {
  // 模拟保存设置
  return { ok: true }
})

// ==================== Web 文件处理 ====================

const WEB_TEMP_ROOT = join(tmpdir(), 'genoffice-web-temp')

registerHandle('web:write-temp-file', async (_event: unknown, request: unknown) => {
  const record = request as { name?: unknown; bytes?: unknown } | null
  if (!record || typeof record.name !== 'string' || !(record.bytes instanceof ArrayBuffer)) {
    throw new Error('web:write-temp-file expects { name: string, bytes: ArrayBuffer }')
  }
  const safeName = basename(record.name).replace(/[^\w.\- ]+/g, '_') || 'file'
  const dir = mkdtempSync(join(WEB_TEMP_ROOT, 'upload-'))
  const filePath = join(dir, safeName)
  const stream = createWriteStream(filePath)
  stream.write(Buffer.from(record.bytes))
  stream.end()
  return filePath
})

registerHandle('web:read-file-bytes', async (_event: unknown, path: unknown) => {
  if (typeof path !== 'string' || !path.startsWith(WEB_TEMP_ROOT) && !path.startsWith(DATA_DIR)) {
    throw new Error('web:read-file-bytes only reads files written by the web bridge')
  }
  if (!existsSync(path as string)) {
    throw new Error(`File not found: ${path}`)
  }
  const bytes = readFileSync(path as string)
  return { name: basename(path as string), bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('web:make-temp-dir', async () => {
  return mkdtempSync(join(tmpdir(), 'genoffice-dir-'))
})

registerHandle('web:save-file', async (_event: unknown, request: unknown) => {
  const { name, bytes, projectId } = request as { name?: string; bytes?: ArrayBuffer; projectId?: string }
  if (!name || !bytes) {
    throw new Error('web:save-file expects { name: string, bytes: ArrayBuffer, projectId?: string }')
  }
  
  const fileId = `${Date.now()}-${name}`
  const filePath = join(FILES_DIR, fileId)
  writeFileSync(filePath, Buffer.from(bytes))
  
  // 如果指定了项目，添加到项目
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

// ==================== HTTP 服务器 ====================

function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  const body = JSON.stringify(payload)
  response.writeHead(status, { 'Content-Type': 'application/json' })
  response.end(body)
}

async function readBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    request.on('data', (chunk: Buffer) => chunks.push(chunk))
    request.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    request.on('error', reject)
  })
}

// Session Hub 用于 SSE
const sessionConnections = new Map<string, Set<ServerResponse>>()
const PENDING_FRAMES = new Map<string, string[]>()
const SSE_HEARTBEAT_MS = 25000

function pushSseEvent(session: string, channel: string, args: unknown[]): void {
  const frame = `data: ${JSON.stringify({ channel, args })}\n\n`
  const connections = sessionConnections.get(session)
  if (connections) {
    for (const response of connections) {
      try {
        response.write(frame)
      } catch {
        // 连接可能已关闭
      }
    }
  } else {
    const pending = PENDING_FRAMES.get(session) || []
    pending.push(frame)
    if (pending.length > 100) pending.shift()
    PENDING_FRAMES.set(session, pending)
  }
}

// 创建 HTTP 服务器
const server = createServer(async (request, response) => {
  const url = new URL(request.url || '/', `http://${request.headers.host}`)
  
  // CORS 头
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
      dataDir: DATA_DIR,
    }))
    return
  }
  
  // 列出所有已实现的通道
  if (url.pathname === '/api/channels' && request.method === 'GET') {
    response.writeHead(200, { 'Content-Type': 'application/json' })
    response.end(JSON.stringify({ channels: [...handlers.keys()].sort() }))
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
    
    // 发送待处理事件
    const pending = PENDING_FRAMES.get(session)
    if (pending) {
      for (const frame of pending) response.write(frame)
      PENDING_FRAMES.delete(session)
    }
    
    // 添加到连接池
    if (!sessionConnections.has(session)) {
      sessionConnections.set(session, new Set())
    }
    sessionConnections.get(session)!.add(response)
    
    // 心跳
    const heartbeat = setInterval(() => {
      try {
        response.write(': heartbeat\n\n')
      } catch {
        clearInterval(heartbeat)
      }
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
  
  // 如果文件不存在，尝试默认应用
  if (!existsSync(filePath)) {
    filePath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', url.pathname === '/' ? 'index.html' : url.pathname)
  }
  
  if (existsSync(filePath) && statSync(filePath).isFile()) {
    const ext = extname(filePath)
    const mimeType = MIME_TYPES[ext] || 'application/octet-stream'
    response.writeHead(200, { 'Content-Type': mimeType })
    createReadStream(filePath).pipe(response)
    return
  }
  
  // 默认返回 docs 应用
  const indexPath = resolve(STATIC_ROOT, 'docs', 'out', 'renderer', 'index.html')
  if (existsSync(indexPath)) {
    response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
    createReadStream(indexPath).pipe(response)
    return
  }
  
  // 未找到
  response.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' })
  response.end('<h1>GenOffice Web Server</h1><p>Please build the apps first: npm run build:all</p>')
})

// 启动服务器
server.listen(PORT, HOST, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║   GenOffice Web Server v0.8.0                             ║
║                                                           ║
║   🌐 URL: http://${HOST}:${PORT}                            ║
║   📁 Mode: Standalone (No Electron)                        ║
║                                                           ║
║   Apps: ${APPS.slice(0, 4).join(', ')}...                   ║
║                                                           ║
║   📊 Implemented Channels: ${String(handlers.size).padEnd(15)}   ║
║                                                           ║
║   API: http://${HOST}:${PORT}/api/ipc/*                      ║
║   SSE: http://${HOST}:${PORT}/api/ipc/events?session=xxx     ║
║   List: http://${HOST}:${PORT}/api/channels                 ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
`)
})

// 优雅关闭
process.on('SIGTERM', () => {
  console.log('\nShutting down GenOffice Web Server...')
  server.close(() => process.exit(0))
})

process.on('SIGINT', () => {
  console.log('\nShutting down GenOffice Web Server...')
  server.close(() => process.exit(0))
})
