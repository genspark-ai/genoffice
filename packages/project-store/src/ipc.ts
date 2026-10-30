/**
 * IPC interface type definitions (shared by the renderer and main processes).
 * No Electron dependency; importable from the renderer.
 */
import type { ProjectStore } from './store.js'
import type {
  ChatAttachment,
  ChatMessage,
  ChatMeta,
  ChatScope,
  ProjectSummary,
  RunEvent,
  RunEventKind,
  RunSummary,
  TimelineEntry,
  ToolActivity,
} from './types.js'

export type {
  ChatAttachment,
  ChatMessage,
  ChatMeta,
  ChatScope,
  ProjectSummary,
  RunEvent,
  RunEventKind,
  RunSummary,
  TimelineEntry,
  ToolActivity,
}

export interface AppendChatArgs {
  projectId: string
  chatId: string
  role: 'user' | 'assistant'
  text: string
  tools?: ToolActivity[]
  attachments?: ChatAttachment[]
  scope?: ChatScope
}

export interface LoadChatArgs {
  projectId: string
  chatId: string
  limit?: number
}

export interface ResolveChatArgs {
  /** Absolute path of the currently open file; null means an unsaved new file */
  filePath: string | null
  /** Temp chatId for unsaved files, e.g. "unsaved-<timestamp>" */
  tempChatId?: string
  /** Sheets mode: look up the file path by sessionId (handled in the main process) */
  sessionId?: string
}

export interface ResolveChatResult {
  projectId: string
  chatId: string
}

export interface RebindChatArgs {
  projectId: string
  tempChatId: string
  /** Specify the new chatId directly; one of newChatId / newFilePath / sessionId */
  newChatId?: string
  /** File path where the unsaved session first hit disk: the main process derives the chatId from it and registers fileMap */
  newFilePath?: string
  /** Sheets mode: the renderer can't get the path, so it passes sessionId for the main process to look up and rebind */
  sessionId?: string
}

// ── Agent run records (first-class runs API over IPC) ──────────────────────

/** the channels the three runs handlers register under */
export const PROJECT_RUN_CHANNELS = {
  appendRunEvent: 'project:appendRunEvent',
  listRuns: 'project:listRuns',
  deleteRun: 'project:deleteRun',
} as const

/** one run-lifecycle event as it crosses IPC; the store stamps the timestamp */
export interface AppendRunEventArgs {
  projectId: string
  runId: string
  kind: RunEventKind
  instruction?: string
  chatId?: string
  outcome?: 'done' | 'cancelled' | 'discarded'
  error?: string
}

/** the slice of Electron's IpcMain the runs handlers register through (tests pass a fake) */
export interface ProjectRunIpcMain {
  handle(channel: string, listener: (event: unknown, args: unknown) => unknown): void
}

/**
 * Register the runs half of the project IPC: appendRunEvent / listRuns /
 * deleteRun against `<project>/runs/*.jsonl`. The store is injected lazily so
 * registration order never matters (the same getProjectStore() laziness the
 * chat handlers rely on); argument shapes are validated here so a malformed
 * renderer payload fails with a named reason instead of reaching the fs layer.
 */
export function registerProjectRunIpc(ipc: ProjectRunIpcMain, store: () => ProjectStore): void {
  ipc.handle(PROJECT_RUN_CHANNELS.appendRunEvent, (_event, args: unknown) => {
    const event = asAppendRunEventArgs(args)
    store().appendRunEvent(event.projectId, event.runId, event)
  })
  ipc.handle(PROJECT_RUN_CHANNELS.listRuns, (_event, args: unknown) => {
    const { projectId } = asProjectRunProjectArgs(args)
    return store().listRuns(projectId)
  })
  ipc.handle(PROJECT_RUN_CHANNELS.deleteRun, (_event, args: unknown) => {
    const { projectId, runId } = asProjectRunTarget(args)
    store().deleteRun(projectId, runId)
  })
}

function asAppendRunEventArgs(args: unknown): AppendRunEventArgs {
  asProjectRunTarget(args)
  const a = args as Partial<AppendRunEventArgs>
  if (a.kind !== 'start' && a.kind !== 'finish' && a.kind !== 'error') {
    throw new Error(`Invalid run event kind: ${String(a.kind)}`)
  }
  for (const [field, kind] of [
    ['instruction', 'string'],
    ['chatId', 'string'],
    ['error', 'string'],
  ] as const) {
    const value = a[field]
    if (value !== undefined && typeof value !== kind) {
      throw new Error(`Invalid run event ${field}: must be a string`)
    }
  }
  if (
    a.outcome !== undefined &&
    a.outcome !== 'done' &&
    a.outcome !== 'cancelled' &&
    a.outcome !== 'discarded'
  ) {
    throw new Error(`Invalid run outcome: ${String(a.outcome)}`)
  }
  return args as AppendRunEventArgs
}

function asProjectRunTarget(args: unknown): { projectId: string; runId: string } {
  const { projectId } = asProjectRunProjectArgs(args)
  const a = args as Partial<AppendRunEventArgs>
  if (typeof a.runId !== 'string' || !a.runId) throw new Error('Invalid runId')
  return { projectId, runId: a.runId }
}

function asProjectRunProjectArgs(args: unknown): { projectId: string } {
  if (!args || typeof args !== 'object') throw new Error('Invalid run request: expected an object')
  const a = args as Partial<AppendRunEventArgs>
  if (typeof a.projectId !== 'string' || !a.projectId) throw new Error('Invalid projectId')
  return { projectId: a.projectId }
}

/** Project storage API the main process exposes to the renderer */
export interface ProjectApi {
  /**
   * Resolves projectId and chatId from a file path.
   * When filePath is null, returns the default project + tempChatId (if provided).
   */
  resolveChat(args: ResolveChatArgs): Promise<ResolveChatResult>
  /** Appends one message to the JSONL */
  appendChat(args: AppendChatArgs): Promise<void>
  /** Reads the most recent `limit` messages */
  loadChat(args: LoadChatArgs): Promise<ChatMessage[]>
  /** Renames the JSONL file (called after the file first hits disk); returns the new projectId/chatId */
  rebindChat(args: RebindChatArgs): Promise<ResolveChatResult>
  /** Every conversation in a project, so a host can list what it kept. */
  listChats(args: { projectId: string }): Promise<ChatMeta[]>
  /** Removes one conversation's transcript. */
  deleteChat(args: { projectId: string; chatId: string }): Promise<void>
}
