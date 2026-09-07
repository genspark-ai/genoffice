/**
 * HTTP Transport for Web Server
 * 
 * 复用 Agent Loop 核心逻辑，通过 HTTP 与后端 AI 服务通信
 */

import type {
  AgentStreamRequest,
  AgentStreamCallbacks,
  AgentStreamHandle,
  AgentTransport,
  AgentToolCall,
} from './types'

export interface HttpStreamChunk {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  toolCall?: AgentToolCall
  error?: string
  errorCode?: 'timeout' | 'credits' | 'network' | 'overloaded'
  stopReason?: string
}

export interface HttpTransportOptions {
  /** Base URL for AI API */
  baseUrl: string
  /** Custom fetch function (for testing) */
  fetch?: typeof fetch
  /** Request timeout in ms */
  timeout?: number
  /** API key for authentication */
  apiKey?: string
  /** Default headers */
  headers?: Record<string, string>
}

const DEFAULT_TIMEOUT = 120_000 // 2 minutes
const PING_INTERVAL = 30_000 // 30 seconds

/**
 * 创建 HTTP Transport 用于 Web 环境
 * 
 * 使用 Server-Sent Events (SSE) 进行流式响应
 */
export function createHttpTransport(options: HttpTransportOptions): AgentTransport {
  const {
    baseUrl,
    fetch: customFetch,
    timeout = DEFAULT_TIMEOUT,
    apiKey,
    headers = {},
  } = options

  const fetchFn = customFetch || globalThis.fetch.bind(globalThis)

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
          const response = await fetchFn(`${baseUrl}/api/ai/stream`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              ...(apiKey ? { 'Authorization': `Bearer ${apiKey}` } : {}),
              ...headers,
            },
            body: JSON.stringify({
              requestId,
              ...request,
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

          const reader = response.body?.getReader()
          if (!reader) {
            callbacks.onError('No response body')
            callbacks.onDone()
            return
          }

          const decoder = new TextDecoder()
          let buffer = ''

          // Ping interval for keepalive
          const pingInterval = setInterval(() => {
            // Send ping to keep connection alive
          }, PING_INTERVAL)

          while (true) {
            const { done, value } = await reader.read()
            if (done || cancelled) break

            buffer += decoder.decode(value, { stream: true })
            const lines = buffer.split('\n')
            buffer = lines.pop() || ''

            for (const line of lines) {
              if (!line.startsWith('data: ')) continue

              const data = line.slice(6)
              if (data === '[DONE]') {
                clearInterval(pingInterval)
                callbacks.onDone()
                return
              }

              try {
                const chunk: HttpStreamChunk = JSON.parse(data)
                handleChunk(chunk)
              } catch {
                // Ignore parse errors
              }
            }
          }

          clearInterval(pingInterval)
          callbacks.onDone()
        } catch (err) {
          clearTimeout(timeoutId)
          if (!cancelled) {
            callbacks.onError(err instanceof Error ? err.message : 'Unknown error')
            callbacks.onDone()
          }
        }
      }

      const handleChunk = (chunk: HttpStreamChunk) => {
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
            break
          case 'error':
            callbacks.onError(chunk.error || 'Unknown error')
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

/**
 * 简单的 HTTP 请求（非流式）
 */
export async function httpRequest<T>(
  url: string,
  options: {
    method?: string
    body?: unknown
    headers?: Record<string, string>
    fetch?: typeof fetch
  } = {},
): Promise<T> {
  const { method = 'GET', body, headers = {}, fetch: customFetch } = options
  const fetchFn = customFetch || globalThis.fetch.bind(globalThis)

  const response = await fetchFn(url, {
    method,
    headers: {
      'Content-Type': 'application/json',
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  })

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${await response.text()}`)
  }

  return response.json()
}
