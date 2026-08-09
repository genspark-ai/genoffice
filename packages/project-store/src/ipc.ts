/**
 * IPC interface type definitions (shared by the renderer and main processes).
 * No Electron dependency; importable from the renderer.
 */
import type {
  ChatAttachment,
  ChatMessage,
  ChatMeta,
  ProjectSummary,
  TimelineEntry,
  ToolActivity,
} from './types.js'

export type { ChatAttachment, ChatMessage, ChatMeta, ProjectSummary, TimelineEntry, ToolActivity }

export interface AppendChatArgs {
  projectId: string
  chatId: string
  role: 'user' | 'assistant'
  text: string
  tools?: ToolActivity[]
  attachments?: ChatAttachment[]
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

// ── P1 extensions ──────────────────────────────────────────

export interface CreateProjectArgs {
  name: string
}

export interface RenameProjectArgs {
  id: string
  name: string
}

export interface DeleteProjectArgs {
  id: string
}

export interface MoveFileArgs {
  filePath: string
  projectId: string
}

export interface GetTimelineArgs {
  projectId: string
  limit?: number
}

/** One file's conversations: the session picker's list, and the two ways to move between them */
export interface FileChatsArgs {
  /** Absolute path of the currently open file; null means an unsaved new file */
  filePath: string | null
  /** Sheets mode: look up the file path by sessionId (handled in the main process) */
  sessionId?: string
  /** Temp chatId for unsaved files, e.g. "unsaved-<timestamp>" */
  tempChatId?: string
}

export interface SwitchChatArgs extends FileChatsArgs {
  chatId: string
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
  // ── P1 extensions ──
  /** Lists all projects (with file count + last active time) */
  listProjects(): Promise<ProjectSummary[]>
  /** Creates a project */
  createProject(args: CreateProjectArgs): Promise<ProjectSummary>
  /** Renames a project */
  renameProject(args: RenameProjectArgs): Promise<void>
  /** Soft-deletes a project (directory moved into .trash) */
  deleteProject(args: DeleteProjectArgs): Promise<void>
  /** Moves a file into the given project */
  moveFile(args: MoveFileArgs): Promise<void>
  /** Gets the project timeline */
  getTimeline(args: GetTimelineArgs): Promise<TimelineEntry[]>
  // ── sessions ──
  /** Every conversation belonging to one file, oldest first */
  listChatsForFile(args: FileChatsArgs): Promise<ChatMeta[]>
  /** Starts a fresh conversation for a file and makes it the active one */
  newChat(args: FileChatsArgs): Promise<ResolveChatResult>
  /** Loads an earlier conversation by making it active again */
  switchChat(args: SwitchChatArgs): Promise<ResolveChatResult>
}
