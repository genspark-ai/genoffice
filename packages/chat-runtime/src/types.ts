/**
 * Public types for the unified Chat model.
 *
 * Anything that's safe for the renderer to depend on lives here so the
 * four apps can mount a `ChatRuntime` against their own skill + transport
 * without each redefining the same shapes.
 */

import type {
  AgentSkill,
  AgentMessage,
  AgentToolCall,
  AgentToolResult,
  ToolExecutedEvent,
} from '@genoffice/agent-core'

export type ChatApp = 'docs' | 'sheets' | 'slides' | 'pdf'

export type ChatRunStatus =
  | 'idle'
  | 'queued'
  | 'running'
  | 'streaming'
  | 'awaiting-confirm'
  | 'done'
  | 'cancelled'
  | 'error'

export interface ChatAttachment {
  id: string
  name: string
  type: string
  size: number
  /** optional URL (data: or http:) */
  url?: string | undefined
  /** optional in-memory preview (base64 thumbnail, etc.) */
  preview?: string | undefined
}

export interface ChatContextRef {
  /** Stable id within the app (selection id, sheet id, slide id, page id, …) */
  id: string
  /** Free-form label shown in the scope quote */
  label: string
  /** Optional structured payload (e.g. selected text, cell range) */
  payload?: Record<string, unknown> | undefined
}

export interface ChatMessage {
  id: string
  role: 'user' | 'assistant' | 'system'
  content: string
  attachments?: ChatAttachment[] | undefined
  timestamp: number
  /** Status of the message within the current run. */
  status?: 'streaming' | 'done' | 'error' | undefined
  error?: string | undefined
  /** True when this message is still streaming and may receive deltas. */
  streaming?: boolean | undefined
  /** Tool records associated with the assistant message. */
  tools?: ChatToolCallRecord[] | undefined
  /** Free-form citation metadata (source, url, …). */
  citations?: Array<{ text: string; source: string; url?: string }> | undefined
}

export interface ChatToolCallRecord {
  id: string
  name: string
  input: Record<string, unknown>
  status: 'running' | 'executed' | 'error'
  output?: string | undefined
  isError?: boolean | undefined
  startedAt: number
  finishedAt?: number | undefined
  /** Snapshot of state captured before this tool ran (for rollback UIs). */
  snapshotBefore?: unknown | undefined
}

export type ChatChangePlanOp =
  | { kind: 'workbook'; ops: unknown[]; description?: string }
  | { kind: 'doc'; ops: unknown[]; description?: string }
  | { kind: 'slide'; ops: unknown[]; description?: string }
  | { kind: 'pdf'; ops: unknown[]; description?: string }
  | { kind: 'freeform'; ops: unknown[]; description?: string }
  | {
      kind: 'translate'
      ops: Array<{
        /** source text, preserved verbatim for the undo step */
        sourceText: string
        /** translated text to write back */
        targetText: string
        /** BCP-47 / free-form target language tag (e.g. "zh-CN") */
        targetLang: string
        /** editor range anchor (app-specific; null = whole document) */
        range?: { from: number; to: number; scope?: string } | null
        /** True when formatting should be inherited automatically (default true) */
        preserveFormat?: boolean
      }>
      description?: string
    }

export interface ChatChangePlan {
  id: string
  /** Origin app — used by the timeline to pick the right preview renderer. */
  app: ChatApp | 'unknown'
  title: string
  summary: string
  /** Ordered list of operations to apply. */
  ops: ChatChangePlanOp[]
  /** Free-form warnings / risk flags the UI should highlight. */
  warnings?: string[] | undefined
  /** When true, the UI should require explicit confirmation before apply. */
  requireConfirm?: boolean | undefined
  createdAt: number
}

export interface ChatCapability {
  /** Provider key (matches AiSettings.provider). */
  provider: string
  /** Model identifier (free-form). */
  model: string
  /** Human-readable label (e.g. "MiniMax M3"). */
  displayName?: string | undefined
}

export interface ChatSession {
  id: string
  app: ChatApp
  title: string
  createdAt: number
  updatedAt: number
  messages: ChatMessage[]
  capability?: ChatCapability | undefined
  /** Optional cross-session scope (e.g. file id, project id). */
  projectId?: string | undefined
  chatId?: string | undefined
}

export interface ChatRun {
  id: string
  sessionId: string
  status: ChatRunStatus
  startedAt: number
  finishedAt?: number | undefined
  /** Live, accumulating assistant text for the current turn. */
  streamingText: string
  /** Tool calls observed during this run (in order). */
  tools: ChatToolCallRecord[]
  /** Most recent change plan emitted by a tool, if any. */
  lastChangePlan?: ChatChangePlan | undefined
  /** Normalised error, if the run failed. */
  error?: import('./errors.js').AIError | undefined
  /** True when the run ended because the user cancelled. */
  cancelled?: boolean | undefined
}

export interface ChatRuntimeOptions {
  app: ChatApp
  sessionId: string
  skill: AgentSkill
  /** Initial list of context refs (selection, sheet range, etc.). */
  initialContext?: ChatContextRef[] | undefined
  /** Initial session messages; runtime appends new turns to this. */
  initialMessages?: ChatMessage[] | undefined
  /** Capability label for the run header (provider + model). */
  capability?: ChatCapability | undefined
  /** Optional snapshot capture callback (forwarded to AgentLoop). */
  captureSnapshot?: (() => unknown) | undefined
  /** Optional system-suffix callback forwarded to AgentLoop. */
  systemSuffix?: (() => string) | undefined
  /** Optional message-format callback forwarded to AgentLoop. */
  formatUserMessage?: ((instruction: string, context: string) => string) | undefined
  /** Maximum turns forwarded to AgentLoop (default 100). */
  maxTurns?: number | undefined
  /** Persistence sink; when provided, sessions are loaded/saved through it. */
  persistence?: ChatPersistence | undefined
}

export interface ChatPersistence {
  load(sessionId: string): Promise<ChatSession | null>
  save(session: ChatSession): Promise<void>
  /** Append a single message to the persisted store; defaults to a full save. */
  append?(sessionId: string, message: ChatMessage): Promise<void>
}

export interface ChatSendOptions {
  instruction: string
  attachments?: ChatAttachment[] | undefined
  context?: ChatContextRef[]
  /** When true, the assistant text + tools are streamed but not persisted. */
  transient?: boolean
}

export interface ChatSendResult {
  run: ChatRun
  messageId: string
  cancelled: boolean
}

/** Snapshot emitted by `ChatRuntime.snapshot()` for one-click rollback UIs. */
export interface ChatSnapshot {
  runId: string
  capturedAt: number
  state: unknown
}

export type ChatEvent =
  | { type: 'session-updated'; session: ChatSession }
  | { type: 'run-start'; run: ChatRun }
  | { type: 'run-text'; runId: string; delta: string; cumulative: string }
  | { type: 'run-tool-start'; runId: string; tool: ChatToolCallRecord }
  | { type: 'run-tool-executed'; runId: string; tool: ChatToolCallRecord }
  | { type: 'run-change-plan'; runId: string; plan: ChatChangePlan }
  | { type: 'run-finish'; run: ChatRun }
  | { type: 'run-error'; run: ChatRun }

export type ChatEventListener = (event: ChatEvent) => void

// Re-export agent-core types so renderer code can `import type { ChatToolCallRecord } from '@genoffice/chat-runtime/types'`
// without needing a second import line for `AgentSkill` / `AgentMessage`.
export type {
  AgentSkill,
  AgentMessage,
  AgentToolCall,
  AgentToolResult,
  ToolExecutedEvent,
}
