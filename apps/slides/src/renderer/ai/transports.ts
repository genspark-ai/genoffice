/**
 * Unified AI Transport Factory
 * 
 * 根据运行环境自动选择合适的 Transport:
 * - Electron 环境: 使用 createIpcTransport
 * - Web 环境: 使用 createHttpTransport (Web Server)
 */

import type { AgentTransport } from '@genoffice/agent-core'
import { createIpcTransport, createWebTransport } from '@genoffice/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { t } from '../i18n/locale'

/检测是否为 Electron 环境
function isElectronRuntime(): boolean {
  return typeof navigator !== 'undefined' && /Electron/i.test(navigator.userAgent)
}

/**
 * 获取 Web Server URL
 */
function getWebServerUrl(): string {
  if (typeof window !== 'undefined') {
    // 开发环境: 检查是否有代理配置
    const url = new URL(window.location.href)
    
    // 如果在 localhost 上使用 8080
    if (url.hostname === 'localhost' || url.hostname === '127.0.0.1') {
      // 检查是否有开发服务器代理
      if (url.port === '5173' || url.port === '5174' || url.port === '5175') {
        // Vite 开发服务器，代理到 Web Server
        return `${url.protocol}//${url.hostname}:8080`
      }
      return 'http://localhost:8080'
    }
    
    // 否则使用当前域名
    return `${url.protocol}//${url.host}`
  }
  
  return 'http://localhost:8080'
}

/**
 * 获取 Web AI 设置
 */
function getWebAiSettings(): { provider: string; model: string; apiKey?: string } {
  if (typeof localStorage !== 'undefined') {
    const stored = localStorage.getItem('genoffice-ai-settings')
    if (stored) {
      try {
        return JSON.parse(stored)
      } catch {
        // Ignore
      }
    }
  }
  
  return {
    provider: 'minimax',
    model: 'MiniMax-M3',
    apiKey: '',
  }
}

/**
 * 创建 Electron Transport (用于 Electron 环境)
 */
function createElectronAiTransport(getSettings: () => AiSettings): AgentTransport {
  // 使用 Electron 的 IPC 通道
  return createIpcTransport<AiSettings>({
    onStream: (listener) => {
      // @ts-expect-error - window.desktop 在 Electron 中可用
      return window.desktop?.onAiStream(listener) || (() => {})
    },
    start: (request) => {
      // @ts-expect-error - window.desktop 在 Electron 中可用
      return window.desktop?.aiStream(request)
    },
    cancel: (requestId) => {
      // @ts-expect-error - window.desktop 在 Electron 中可用
      window.desktop?.aiStreamCancel(requestId)
    },
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    networkErrorText: () => t('aiNetworkError'),
    overloadedErrorText: () => t('aiOverloadedError'),
  })
}

/**
 * 创建 Web Transport (用于 Web 环境)
 */
function createWebAiTransport(): AgentTransport {
  const settings = getWebAiSettings()
  
  return createWebTransport({
    baseUrl: getWebServerUrl(),
    apiKey: settings.apiKey,
    timeout: 120_000,
  })
}

/**
 * 创建统一的 AI Transport
 * 
 * 根据运行环境自动选择:
 * - Electron: 使用 IPC 通道
 * - Web: 使用 HTTP 通道
 */
export function createAiTransport(getSettings: () => AiSettings): AgentTransport {
  if (isElectronRuntime()) {
    return createElectronAiTransport(getSettings)
  }
  return createWebAiTransport()
}

/**
 * 检测是否在 Web 模式下运行
 */
export function isWebMode(): boolean {
  return !isElectronRuntime()
}

/**
 * 获取 Web Server URL
 */
export { getWebServerUrl }
