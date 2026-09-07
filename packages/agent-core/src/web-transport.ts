/**
 * Web Transport for GenOffice Web Server
 * 
 * 前端使用此 Transport 与 Web Server 通信
 * 支持 IPC 调用和 SSE 流式响应
 */

import type {
  AgentStreamRequest,
  AgentTransport,
  AgentStreamHandle,
  AgentStreamCallbacks,
  AgentToolCall,
  AgentMessage,
  AgentToolDef,
} from './types'

export interface WebTransportOptions {
  /** Web Server base URL */
  baseUrl: string
  /** API key for authentication (optional) */
  apiKey?: string
  /** Request timeout in ms */
  timeout?: number
  /** SSE reconnect attempts */
  reconnectAttempts?: number
  /** SSE reconnect delay in ms */
  reconnectDelay?: number
}

const DEFAULT_TIMEOUT = 120_000
const DEFAULT_RECONNECT_ATTEMPTS = 3
const DEFAULT_RECONNECT_DELAY = 1000

/**
 * Web 环境检测
 */
export function isWebEnvironment(): boolean {
  return typeof window !== 'undefined' && !('desktop' in window)
}

/**
 * 创建 Web Transport
 * 
 * 使用 HTTP POST 进行 IPC 调用
 * 使用 SSE 进行流式响应
 */
export function createWebTransport(options: WebTransportOptions): AgentTransport {
  const {
    baseUrl,
    apiKey,
    timeout = DEFAULT_TIMEOUT,
    reconnectAttempts = DEFAULT_RECONNECT_ATTEMPTS,
    reconnectDelay = DEFAULT_RECONNECT_DELAY,
  } = options

  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  }
  if (apiKey) {
    headers['Authorization'] = `Bearer ${apiKey}`
  }

  let eventSource: EventSource | null = null
  let reconnectCount = 0

  return {
    stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle {
      const requestId = crypto.randomUUID()
      let cancelled = false
      let abortController: AbortController | undefined

      const startStream = async () => {
        abortController = new AbortController()

        const timeoutId = setTimeout(() => {
          abortController?.abort()
          callbacks.onError('Request timeout')
        }, timeout)

        try {
          // 使用 HTTP POST 发起请求，SSE 获取响应
          const response = await fetch(`${baseUrl}/api/ai/stream`, {
            method: 'POST',
            headers,
            body: JSON.stringify({
              requestId,
              settings: request.settings,
              system: request.system,
              messages: request.messages,
              tools: request.tools,
              maxTokens: request.maxTokens,
            }),
            signal: abortController.signal,
          })

          clearTimeout(timeoutId)

          if (!response.ok) {
            const errorText = await response.text()
            callbacks.onError(`HTTP ${response.status}: ${errorText}`)
            callbacks.onDone()
            return
          }

          // 处理 SSE 流
          const reader = response.body?.getReader()
          if (!reader) {
            callbacks.onError('No response body')
            callbacks.onDone()
            return
          }

          const decoder = new TextDecoder()
          let buffer = ''

          while (true) {
            if (cancelled) break

            const { done, value } = await reader.read()
            if (done) break

            buffer += decoder.decode(value, { stream: true })
            const lines = buffer.split('\n')
            buffer = lines.pop() || ''

            for (const line of lines) {
              if (!line.startsWith('data: ')) continue

              const data = line.slice(6)
              if (data === '[DONE]') {
                callbacks.onDone()
                return
              }

              try {
                const chunk = JSON.parse(data)
                handleChunk(chunk)
              } catch {
                // Ignore parse errors
              }
            }
          }

          callbacks.onDone()
        } catch (err) {
          clearTimeout(timeoutId)
          if (!cancelled) {
            callbacks.onError(err instanceof Error ? err.message : 'Unknown error')
            callbacks.onDone()
          }
        }
      }

      const handleChunk = (chunk: WebStreamChunk) => {
        if (chunk.requestId !== requestId) return

        switch (chunk.type) {
          case 'delta':
            if (chunk.text) callbacks.onDelta(chunk.text)
            break
          case 'reasoning':
            if (chunk.text) callbacks.onReasoning?.(chunk.text)
            break
          case 'tool-call':
            if (chunk.toolCall) callbacks.onToolCall(chunk.toolCall)
            break
          case 'done':
            if (chunk.stopReason) callbacks.onStopReason?.(chunk.stopReason)
            callbacks.onDone()
            break
          case 'error':
            callbacks.onError(chunk.error || 'Unknown error')
            callbacks.onDone()
            break
          case 'ping':
            // Keepalive, no action needed
            break
        }
      }

      // Start the stream
      startStream()

      return {
        cancel() {
          cancelled = true
          abortController?.abort()
        },
      }
    },
  }
}

interface WebStreamChunk {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  toolCall?: AgentToolCall
  error?: string
  stopReason?: string
}

/**
 * Web IPC 客户端
 * 
 * 用于调用 Web Server 的 IPC 通道
 */
export class WebIpcClient {
  private baseUrl: string
  private apiKey?: string

  constructor(baseUrl: string, apiKey?: string) {
    this.baseUrl = baseUrl.replace(/\/$/, '')
    this.apiKey = apiKey
  }

  private getHeaders(): Record<string, string> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    }
    if (this.apiKey) {
      headers['Authorization'] = `Bearer ${this.apiKey}`
    }
    return headers
  }

  /**
   * 调用 IPC 通道
   */
  async invoke<T = unknown>(channel: string, ...args: unknown[]): Promise<T> {
    const response = await fetch(`${this.baseUrl}/api/ipc/${channel}`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({ args }),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error?.message || `HTTP ${response.status}`)
    }

    const result = await response.json()
    if (!result.ok) {
      throw new Error(result.error?.message || 'Unknown error')
    }

    return result.result as T
  }

  /**
   * 创建 SSE 连接用于事件监听
   */
  createEventSource(sessionId: string): EventSource {
    let url = `${this.baseUrl}/api/ipc/events?session=${encodeURIComponent(sessionId)}`
    if (this.apiKey) {
      url += `&apiKey=${encodeURIComponent(this.apiKey)}`
    }
    return new EventSource(url)
  }

  /**
   * 创建 IPC 传输实例
   */
  createTransport(): AgentTransport {
    return createWebTransport({
      baseUrl: this.baseUrl,
      apiKey: this.apiKey,
    })
  }
}

/**
 * 创建默认的 Web IPC 客户端
 * 
 * 自动检测 Web Server 地址
 */
export function createWebIpcClient(baseUrl?: string): WebIpcClient {
  const url = baseUrl || (typeof window !== 'undefined' ? window.location.origin : 'http://localhost:8080')
  return new WebIpcClient(url)
}

/**
 * Web 模式下的 AI 设置
 */
export interface WebAiSettings {
  provider: string
  model: string
  apiKey?: string
}

/**
 * 创建 Web 模式 Transport 的工厂函数
 */
export function createWebAgentTransport(
  settings: () => WebAiSettings,
  errorTexts: {
    unknown: string
    timeout?: string
    credits?: string
    network?: string
    overloaded?: string
  },
): AgentTransport {
  return createWebTransport({
    baseUrl: getWebServerUrl(),
    apiKey: settings().apiKey,
  })
}

/**
 * 获取 Web Server URL
 */
function getWebServerUrl(): string {
  if (typeof window !== 'undefined') {
    // 在浏览器环境中，使用当前域名
    return window.location.origin
  }
  return 'http://localhost:8080'
}

// 导出类型
export type { AgentStreamRequest, AgentTransport, AgentStreamHandle, AgentStreamCallbacks }
