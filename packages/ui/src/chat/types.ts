/**
 * GenOffice Chat UI - 统一 AI Chat 组件库
 * 
 * 提供统一的 Chat 会话管理、消息展示、命令面板等功能
 */

// ========== 会话类型 ==========

export interface ChatSession {
  id: string
  title: string
  createdAt: number
  updatedAt: number
  model?: string
  pinned?: boolean
  tags?: string[]
}

export interface ChatMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  attachments?: ChatAttachment[]
  model?: string
  timestamp: number
  status?: 'streaming' | 'done' | 'error'
  error?: string
  tools?: ChatToolExecution[]
  citations?: ChatCitation[]
}

export interface ChatAttachment {
  id: string
  name: string
  type: string
  size: number
  url?: string
  preview?: string
}

// ========== 工具执行 ==========

export interface ChatToolExecution {
  name: string
  summary: string
  running?: boolean
  isError?: boolean
  input?: string
  output?: string
}

// ========== 引用 ==========

export interface ChatCitation {
  text: string
  source: string
  url?: string
}

// ========== AI 模型 ==========

export interface AIModel {
  id: string
  name: string
  provider: string
  description?: string
  maxTokens?: number
  supportsStreaming?: boolean
  supportsVision?: boolean
}

// ========== 命令 ==========

export interface ChatCommand {
  id: string
  label: string
  shortcut?: string
  icon?: string
  category?: string
  action: () => void | Promise<void>
}

// ========== 主题 ==========

export type ChatTheme = 'light' | 'dark' | 'system'

// ========== 组件 Props ==========

export interface ChatPanelProps {
  // 会话管理
  sessions?: ChatSession[]
  currentSession?: string
  onSessionChange?: (id: string) => void
  onSessionCreate?: () => void
  onSessionDelete?: (id: string) => void
  onSessionRename?: (id: string, name: string) => void
  
  // 消息
  messages: ChatMessage[]
  onMessageSend: (text: string, attachments?: ChatAttachment[]) => void
  onMessageStop?: () => void
  onMessageCopy?: (id: string, text: string) => void
  onMessageEdit?: (id: string, text: string) => void
  onMessageDelete?: (id: string) => void
  onMessageRetry?: (id: string) => void
  
  // 附件
  onAttachmentAdd?: (files: File[]) => void
  onAttachmentRemove?: (id: string) => void
  
  // AI 设置
  models?: AIModel[]
  currentModel?: string
  onModelChange?: (model: string) => void
  temperature?: number
  onTemperatureChange?: (t: number) => void
  
  // 命令面板
  commands?: ChatCommand[]
  onCommandExecute?: (command: ChatCommand) => void
  
  // 样式
  theme?: ChatTheme
  compact?: boolean
  showSidebar?: boolean
  placeholder?: string
  
  // 回调
  isStreaming?: boolean
  isLoading?: boolean
}

export interface ChatSidebarProps {
  sessions: ChatSession[]
  currentSession?: string
  onSessionSelect: (id: string) => void
  onSessionCreate: () => void
  onSessionDelete: (id: string) => void
  onSessionRename: (id: string, name: string) => void
  onSearch?: (query: string) => void
  theme?: ChatTheme
  collapsed?: boolean
  onToggleCollapse?: () => void
}

export interface ChatCommandPaletteProps {
  commands: ChatCommand[]
  onExecute: (command: ChatCommand) => void
  onClose: () => void
  theme?: ChatTheme
}

export interface ChatSettingsProps {
  models: AIModel[]
  currentModel: string
  onModelChange: (model: string) => void
  temperature: number
  onTemperatureChange: (t: number) => void
  onClose: () => void
  theme?: ChatTheme
}
