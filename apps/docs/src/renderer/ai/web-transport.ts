/**
 * Web Mode Transport for Docs AI
 * 
 * 在 Web 环境下使用 Web Server 进行通信
 */

import type { AgentTransport } from '@genoffice/agent-core'
import { createWebTransport, isWebEnvironment } from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t } from '../i18n/locale'

/**
 * 创建 Web Transport
 * 
 * 用于 Web 模式下与 Web Server 通信
 */
export function createWebTransportForDocs(
  baseUrl: string,
  apiKey?: string,
): AgentTransport {
  return createWebTransport({
    baseUrl,
    apiKey,
    timeout: 120_000,
  })
}

/**
 * 检测是否为 Web 环境
 */
export function isRunningInWebMode(): boolean {
  return isWebEnvironment()
}

/**
 * 获取 Web Server URL
 */
export function getWebServerUrl(): string {
  if (typeof window !== 'undefined') {
    // 开发环境: http://localhost:8080
    // 生产环境: 当前域名
    const url = new URL(window.location.href)
    
    // 如果在 localhost 上使用 8080
    if (url.hostname === 'localhost' || url.hostname === '127.0.0.1') {
      return 'http://localhost:8080'
    }
    
    // 否则使用当前域名
    return `${url.protocol}//${url.host}`
  }
  
  return 'http://localhost:8080'
}

/**
 * Web 模式下的 AI 设置接口
 */
export interface WebAiSettings {
  provider: string
  model: string
  apiKey?: string
}

/**
 * 获取 Web AI 设置
 */
export function getWebAiSettings(): WebAiSettings {
  // 从 localStorage 读取设置
  const stored = localStorage.getItem('genoffice-ai-settings')
  if (stored) {
    try {
      return JSON.parse(stored)
    } catch {
      // Ignore
    }
  }
  
  return {
    provider: 'minimax',
    model: 'MiniMax-M3',
    apiKey: '',
  }
}

/**
 * 保存 Web AI 设置
 */
export function saveWebAiSettings(settings: WebAiSettings): void {
  localStorage.setItem('genoffice-ai-settings', JSON.stringify(settings))
}

/**
 * 创建 Web Transport 的工厂函数
 */
export function createWebDocsTransport(): AgentTransport {
  const settings = getWebAiSettings()
  
  return createWebTransport({
    baseUrl: getWebServerUrl(),
    apiKey: settings.apiKey,
    timeout: 120_000,
  })
}

/**
 * Web Transport 的回调接口
 */
export interface WebTransportCallbacks {
  onDelta: (text: string) => void
  onReasoning?: (text: string) => void
  onToolCall?: (toolCall: unknown) => void
  onStopReason?: (reason: string) => void
  onError: (error: string) => void
  onDone: () => void
}

/**
 * 直接使用 fetch 调用 IPC 通道 (非流式)
 */
export async function webIpcInvoke<T = unknown>(
  channel: string,
  ...args: unknown[]
): Promise<T> {
  const baseUrl = getWebServerUrl()
  
  const response = await fetch(`${baseUrl}/api/ipc/${channel}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
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
 * Web 模式下 AI 请求的请求体
 */
export interface WebAiRequest {
  requestId: string
  settings: WebAiSettings
  system: string
  messages: Array<{ role: string; content: string }>
  tools?: Array<{ name: string; description: string; input_schema: unknown }>
  maxTokens?: number
}

/**
 * Web 模式下 AI 响应的 chunk
 */
export interface WebAiChunk {
  requestId: string
  type: 'delta' | 'reasoning' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  toolCall?: unknown
  error?: string
  stopReason?: string
}

/**
 * 在 Web 模式下发送 AI 请求
 */
export async function webAiRequest(
  request: WebAiRequest,
  callbacks: WebTransportCallbacks,
): Promise<() => void> {
  const baseUrl = getWebServerUrl()
  const controller = new AbortController()
  
  const timeoutId = setTimeout(() => {
    controller.abort()
    callbacks.onError('Request timeout')
  }, 120_000)

  try {
    const response = await fetch(`${baseUrl}/api/ai/stream`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(request),
      signal: controller.signal,
    })

    clearTimeout(timeoutId)

    if (!response.ok) {
      const errorText = await response.text()
      callbacks.onError(`HTTP ${response.status}: ${errorText}`)
      callbacks.onDone()
      return () => {}
    }

    const reader = response.body?.getReader()
    if (!reader) {
      callbacks.onError('No response body')
      callbacks.onDone()
      return () => {}
    }

    const decoder = new TextDecoder()
    let buffer = ''

    const processStream = async () => {
      while (true) {
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
            const chunk: WebAiChunk = JSON.parse(data)
            if (chunk.requestId !== request.requestId) continue

            switch (chunk.type) {
              case 'delta':
                if (chunk.text) callbacks.onDelta(chunk.text)
                break
              case 'reasoning':
                if (chunk.text) callbacks.onReasoning?.(chunk.text)
                break
              case 'tool-call':
                if (chunk.toolCall) callbacks.onToolCall?.(chunk.toolCall)
                break
              case 'done':
                if (chunk.stopReason) callbacks.onStopReason?.(chunk.stopReason)
                callbacks.onDone()
                return
              case 'error':
                callbacks.onError(chunk.error || 'Unknown error')
                callbacks.onDone()
                return
            }
          } catch {
            // Ignore parse errors
          }
        }
      }

      callbacks.onDone()
    }

    processStream()

    // 返回取消函数
    return () => {
      controller.abort()
    }
  } catch (err) {
    clearTimeout(timeoutId)
    if (!controller.signal.aborted) {
      callbacks.onError(err instanceof Error ? err.message : 'Unknown error')
      callbacks.onDone()
    }
    return () => {}
  }
}
