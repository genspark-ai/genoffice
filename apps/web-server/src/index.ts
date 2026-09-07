/**
 * GenOffice Web Server
 * 
 * 独立的 Web Server 版本，不需要 Electron。
 * 提供：
 * 1. 静态文件服务
 * 2. HTTP IPC Bridge（模拟 Electron IPC）
 * 3. 应用路由
 * 
 * 使用方式：
 *   npm run dev     # 开发模式
 *   npm run build && npm start  # 生产模式
 */

import { createServer, type IncomingMessage, ServerResponse } from 'node:http'
import { createReadStream, existsSync, statSync, createWriteStream, mkdtempSync, readFileSync } from 'node:fs'
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
}

// 简单的 IPC Handler Registry
interface IpcHandler {
  (event: unknown, ...args: unknown[]): unknown
}

const handlers = new Map<string, IpcHandler>()

function registerHandle(channel: string, handler: IpcHandler): void {
  handlers.set(channel, handler)
}

// 注册 IPC 处理程序
registerHandle('app:get-language', () => 'zh-CN')
registerHandle('app:get-version', () => '0.8.0')
registerHandle('app:get-platform', () => 'web')
registerHandle('docs:get-settings', () => ({}))
registerHandle('docs:save-settings', () => ({ ok: true }))
registerHandle('project:list', () => [])
registerHandle('project:create', (_event: unknown, name: unknown) => ({ id: Date.now().toString(), name }))
registerHandle('project:open', (_event: unknown, id: unknown) => ({ id, name: 'Sample Project' }))

// Web 文件处理
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
  if (typeof path !== 'string' || !path.startsWith(WEB_TEMP_ROOT)) {
    throw new Error('web:read-file-bytes only reads files written by the web bridge')
  }
  const bytes = readFileSync(path)
  return { name: basename(path), bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
})

registerHandle('web:make-temp-dir', async () => {
  return mkdtempSync(join(tmpdir(), 'genoffice-dir-'))
})

// HTTP IPC Bridge 辅助函数
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
    response.end(JSON.stringify({ status: 'ok', version: '0.8.0', mode: 'web-server' }))
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
        sendJson(response, 404, { error: { message: `No handler for '${channel}'` } })
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
║   API: http://${HOST}:${PORT}/api/ipc/*                      ║
║   SSE: http://${HOST}:${PORT}/api/ipc/events?session=xxx     ║
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
