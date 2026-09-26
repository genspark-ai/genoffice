/**
 * project-store core implementation
 *
 * Storage layout (baseDir = userData/projects/):
 *   index.json
 *   <project-id>/
 *     project.json
 *     chats/
 *       <chat-id>.jsonl
 *
 * Design principles:
 * - No Electron dependency; the userData path is injected by the caller
 * - All write failures warn silently, never throw (append path)
 * - JSONL parsing is line-by-line tolerant: bad lines are skipped, no crash
 * - seq is maintained by the store layer: auto-incremented on each appendChatMessage
 */

import { createHash, randomBytes } from 'node:crypto'
import {
  appendFileSync,
  closeSync,
  existsSync,
  mkdirSync,
  openSync,
  readFileSync,
  readSync,
  realpathSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
  readdirSync,
} from 'node:fs'
import { basename, dirname, join } from 'node:path'
import type {
  ChatMeta,
  ChatMessage,
  ProjectData,
  ProjectIndex,
  ProjectInfo,
  ProjectSummary,
  TimelineEntry,
} from './types.js'

// ────────────────────────────────────────────────────────────
// Internal helpers
// ────────────────────────────────────────────────────────────

/** Max stored characters for a single tool input/output field */
const TOOL_FIELD_MAX_CHARS = 16_000

/**
 * Max stored characters for message text. A model that falls into a repetition
 * loop can emit megabytes in one turn; stored whole it would both bloat the
 * JSONL line and be replayed into the model context when the file reopens.
 */
const TEXT_MAX_CHARS = 32_000
/** Max stored characters of a scope excerpt */
const SCOPE_TEXT_MAX_CHARS = 400
const TEXT_TRUNCATED_MARK = '\n\n[truncated]'
/**
 * Max opening messages buffered in memory per chat before the first
 * assistant message materializes the file. Without a cap, thousands of
 * pre-first-reply user messages accumulate unboundedly in the map.
 */
export const MAX_PENDING_OPENING_MESSAGES = 200

function nowIso(): string {
  return new Date().toISOString()
}

function ensureDir(dir: string): void {
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true })
}

// Default number of chat messages returned by loadChat when limit is missing or not finite
const DEFAULT_CHAT_LIMIT = 200
// Upper bound for loadChat limit to avoid unbounded reads
const MAX_CHAT_LIMIT = 10_000
const MAX_CHAT_FILE_READ_BYTES = 8 * 1024 * 1024
/** Max project name chars: prevents MB names bloating index.json/project.json. */
export const MAX_PROJECT_NAME_CHARS = 128
/** Default timeline entries; upper bound avoids loading every chat fully. */
const DEFAULT_TIMELINE_LIMIT = 20
const MAX_TIMELINE_LIMIT = 1_000

function normalizeTimelineLimit(limit: number): number {
  if (!Number.isFinite(limit)) return DEFAULT_TIMELINE_LIMIT
  const floored = Math.floor(limit)
  if (floored < 1) return 1
  if (floored > MAX_TIMELINE_LIMIT) return MAX_TIMELINE_LIMIT
  return floored
}

// Allowlist for project and chat ids (fail-closed: rejects traversal and separators)
const SAFE_ID_PATTERN = /^[A-Za-z0-9_-]+$/

// Throws a descriptive Error when an id could escape the store directory
function assertSafeId(value: string, kind: 'projectId' | 'chatId'): void {
  if (typeof value !== 'string' || !SAFE_ID_PATTERN.test(value)) {
    throw new Error(
      `Invalid ${kind} "${value}": must be non-empty and match ${String(SAFE_ID_PATTERN)} (rejects "..", "/" and backslash)`,
    )
  }
}

// Clamps limit to a finite integer in 1..MAX_CHAT_LIMIT (non-finite falls back to default)
function normalizeChatLimit(limit: number): number {
  if (!Number.isFinite(limit)) return DEFAULT_CHAT_LIMIT
  const floored = Math.floor(limit)
  if (floored < 1) return 1
  if (floored > MAX_CHAT_LIMIT) return MAX_CHAT_LIMIT
  return floored
}

/**
 * The one key a file path is stored and hashed under. The shell canonicalizes
 * open paths with realpathSync.native, so the store has to agree: a symlink, a
 * Windows case or 8.3 variant and a macOS Unicode-normalization variant of one
 * physical document must all land on the same project membership and chat id.
 *
 * A leaf that cannot be resolved (not written yet, or already renamed away) is
 * keyed under its resolved directory, and a path whose directory does not resolve
 * either is used as given.
 */
export function canonicalPathKey(filePath: string): string {
  try {
    return realpathSync.native(filePath).normalize('NFC')
  } catch {
    // The leaf may be gone (renamed or deleted) while its directory still
    // resolves, which is the case fileRenamed has to look up.
  }
  try {
    return join(realpathSync.native(dirname(filePath)), basename(filePath)).normalize('NFC')
  } catch {
    return unresolvedPathKey(filePath)
  }
}

/** Key for a path that has no resolvable target yet, used to find earlier entries. */
function unresolvedPathKey(filePath: string): string {
  return filePath.normalize('NFC')
}

function hashPathKey(key: string): string {
  return createHash('sha256').update(key).digest('hex').slice(0, 16)
}

function readChatTail(filePath: string): string {
  const size = statSync(filePath).size
  if (size === 0) return ''
  const partialTail = size > MAX_CHAT_FILE_READ_BYTES
  const length = partialTail ? MAX_CHAT_FILE_READ_BYTES - 1 : size
  const buffer = Buffer.allocUnsafe(length)
  const fd = openSync(filePath, 'r')
  try {
    const start = size - length
    const bytesRead = readSync(fd, buffer, 0, length, start)
    let contentStart = 0
    if (partialTail) {
      const boundary = Buffer.allocUnsafe(1)
      readSync(fd, boundary, 0, 1, start - 1)
      if (boundary[0] !== 0x0a) {
        const newline = buffer.subarray(0, bytesRead).indexOf(0x0a)
        contentStart = newline >= 0 ? newline + 1 : bytesRead
      }
    }
    return buffer.subarray(contentStart, bytesRead).toString('utf8')
  } finally {
    closeSync(fd)
  }
}

function parseChatRecords(raw: string): ChatMessage[] {
  const messages: ChatMessage[] = []
  for (const line of raw.split('\n')) {
    if (!line.trim()) continue
    try {
      const msg = JSON.parse(line) as ChatMessage
      if (
        typeof msg.seq === 'number' &&
        typeof msg.role === 'string' &&
        typeof msg.text === 'string'
      ) {
        messages.push(msg)
      }
    } catch {
      continue
    }
  }
  return messages
}

function readAllChatRecords(filePath: string): ChatMessage[] {
  return parseChatRecords(readFileSync(filePath, 'utf8'))
}

function appendJsonLines(filePath: string, lines: string): void {
  if (!lines) return
  let prefix = ''
  try {
    const size = statSync(filePath).size
    if (size > 0) {
      const fd = openSync(filePath, 'r')
      try {
        const lastByte = Buffer.allocUnsafe(1)
        readSync(fd, lastByte, 0, 1, size - 1)
        if (lastByte[0] !== 0x0a) prefix = '\n'
      } finally {
        closeSync(fd)
      }
    }
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
  appendFileSync(filePath, prefix + lines, 'utf8')
}

function mergeChatFiles(oldPath: string, newPath: string): number {
  const existing = readAllChatRecords(newPath)
  const moved = readAllChatRecords(oldPath)
  let seq = existing.reduce((m, msg) => Math.max(m, msg.seq), -1) + 1
  const movedLines = moved
    .map((message) => JSON.stringify({ ...message, seq: seq++ }) + '\n')
    .join('')
  if (movedLines) {
    const target = readFileSync(newPath)
    const boundary =
      target.length > 0 && target[target.length - 1] !== 0x0a ? Buffer.from('\n') : Buffer.alloc(0)
    const tmpPath = `${newPath}.${randomBytes(6).toString('hex')}.tmp`
    try {
      writeFileSync(tmpPath, Buffer.concat([target, boundary, Buffer.from(movedLines, 'utf8')]))
      renameSync(tmpPath, newPath)
    } catch (error) {
      try {
        unlinkSync(tmpPath)
      } catch {}
      throw error
    }
  }
  unlinkSync(oldPath)
  return seq - 1
}

function readJson<T>(filePath: string): T | null {
  try {
    if (!existsSync(filePath)) return null
    return JSON.parse(readFileSync(filePath, 'utf8')) as T
  } catch {
    return null
  }
}

/** Atomic write: write to .tmp then rename, so a process interruption can't leave half-written JSON */
function writeJson(filePath: string, data: unknown): void {
  ensureDir(dirname(filePath))
  const tmpPath = `${filePath}.tmp`
  writeFileSync(tmpPath, JSON.stringify(data, null, 2), 'utf8')
  renameSync(tmpPath, filePath)
}

// ────────────────────────────────────────────────────────────
// ProjectStore class
// ────────────────────────────────────────────────────────────

export class ProjectStore {
  private baseDir: string

  constructor(userDataPath: string) {
    this.baseDir = join(userDataPath, 'projects')
  }

  // ── Path helpers ──────────────────────────────────────────

  private indexPath(): string {
    return join(this.baseDir, 'index.json')
  }

  private projectDir(projectId: string): string {
    assertSafeId(projectId, 'projectId')
    return join(this.baseDir, projectId)
  }

  private projectJsonPath(projectId: string): string {
    return join(this.projectDir(projectId), 'project.json')
  }

  private chatsDir(projectId: string): string {
    return join(this.projectDir(projectId), 'chats')
  }

  private chatPath(projectId: string, chatId: string): string {
    assertSafeId(chatId, 'chatId')
    return join(this.chatsDir(projectId), `${chatId}.jsonl`)
  }

  /**
   * The key a path is stored under in one of the index maps, looking through the
   * alternate spellings as well: an older version keyed entries under the raw
   * path exactly as the shell passed it (on macOS that is the NFD spelling the
   * file system reports, which canonicalization folds to NFC), and a path
   * registered while the file did not exist yet is keyed unresolved and resolves
   * to a real path once it is written.
   */
  private findMapKey(
    map: Record<string, string> | undefined,
    filePath: string,
  ): string | undefined {
    if (!map) return undefined
    const key = canonicalPathKey(filePath)
    if (map[key] !== undefined) return key
    for (const alt of [filePath, unresolvedPathKey(filePath)]) {
      if (alt !== key && map[alt] !== undefined) return alt
    }
    return undefined
  }

  /**
   * Moves an entry found under a legacy or unresolved key to the canonical key,
   * so later lookups through any spelling hit the same single entry.
   */
  private static rekeyEntry(map: Record<string, string>, fromKey: string, toKey: string): void {
    if (fromKey === toKey) return
    map[toKey] = map[fromKey]!
    delete map[fromKey]
  }

  /** Same identity test for the raw path lists kept in project.json */
  private static ownsFile(files: readonly string[], filePath: string): boolean {
    const key = canonicalPathKey(filePath)
    return files.some((f) => canonicalPathKey(f) === key)
  }

  // ── seq counters (in-memory cache, initialized from JSONL line count on first read) ──

  /** projectId:chatId → current max seq */
  private readonly seqCounters = new Map<string, number>()

  private seqKey(projectId: string, chatId: string): string {
    return `${projectId}:${chatId}`
  }

  private nextSeq(projectId: string, chatId: string): number {
    const key = this.seqKey(projectId, chatId)
    const cur = this.seqCounters.get(key)
    if (cur !== undefined) {
      const next = cur + 1
      this.seqCounters.set(key, next)
      return next
    }
    // Initialization: scan the existing file for the max seq
    const existing = this.loadChat(projectId, chatId, 10_000)
    const maxSeq = existing.reduce((m, msg) => Math.max(m, msg.seq), -1)
    const next = maxSeq + 1
    this.seqCounters.set(key, next)
    return next
  }

  // ── Index read/write ──────────────────────────────────────

  private readIndex(): ProjectIndex {
    return readJson<ProjectIndex>(this.indexPath()) ?? { projects: [], fileMap: {} }
  }

  private writeIndex(index: ProjectIndex): void {
    ensureDir(this.baseDir)
    writeJson(this.indexPath(), index)
  }

  // ── Project read/write ────────────────────────────────────

  private readProject(projectId: string): ProjectData | null {
    return readJson<ProjectData>(this.projectJsonPath(projectId))
  }

  private writeProject(data: ProjectData): void {
    writeJson(this.projectJsonPath(data.id), data)
  }

  // ────────────────────────────────────────────────────────────
  // Public API
  // ────────────────────────────────────────────────────────────

  /**
   * Ensures the default project exists (id: "default", name: the Chinese "default project" label).
   * Idempotent: returns directly if it already exists.
   */
  ensureDefaultProject(): ProjectData {
    const existing = this.readProject('default')
    if (existing) return existing

    const now = nowIso()
    const data: ProjectData = {
      id: 'default',
      name: 'Default Project',
      createdAt: now,
      updatedAt: now,
      files: [],
    }
    ensureDir(this.projectDir('default'))
    this.writeProject(data)

    const index = this.readIndex()
    if (!index.projects.find((p) => p.id === 'default')) {
      index.projects.unshift({ id: data.id, name: data.name, createdAt: now, updatedAt: now })
      this.writeIndex(index)
    }
    return data
  }

  /**
   * Looks up the projectId by absolute file path.
   * If not found, assign to default and register in fileMap.
   */
  resolveProjectForFile(filePath: string): string {
    this.ensureDefaultProject()
    const index = this.readIndex()
    const key = canonicalPathKey(filePath)
    const existingKey = this.findMapKey(index.fileMap, filePath)
    if (existingKey !== undefined) {
      const projectId = index.fileMap[existingKey]!
      if (existingKey !== key) {
        ProjectStore.rekeyEntry(index.fileMap, existingKey, key)
        this.writeIndex(index)
      }
      return projectId
    }

    // Assign to default
    index.fileMap[key] = 'default'
    this.writeIndex(index)

    // Update the files list in project.json
    const proj = this.readProject('default')
    if (proj && !ProjectStore.ownsFile(proj.files, filePath)) {
      proj.files.push(filePath)
      proj.updatedAt = nowIso()
      this.writeProject(proj)
    }
    return 'default'
  }

  /**
   * Derives a chatId from a file path (first 16 hex chars of sha256 of the
   * canonical path key). Unsaved files use an externally provided temp id
   * (e.g. "unsaved-<timestamp>").
   * Only a fallback derivation for old data without a mapping; new code uses
   * resolveChatForFile (stable mapping).
   */
  static chatIdForFile(filePath: string): string {
    return hashPathKey(canonicalPathKey(filePath))
  }

  /**
   * The ids a path had before canonicalization, so chats written by an older
   * version under the raw path hash are still found: the raw string exactly as
   * given (older macOS entries are NFD) and its NFC form.
   */
  private static legacyChatIdsForFile(filePath: string): string[] {
    return [...new Set([hashPathKey(filePath), hashPathKey(unresolvedPathKey(filePath))])]
  }

  /**
   * Chat id for a path with no registered mapping: the canonical hash, unless an
   * older version already wrote this chat under a raw path hash.
   */
  private fallbackChatId(projectId: string | undefined, filePath: string): string {
    const chatId = ProjectStore.chatIdForFile(filePath)
    if (!projectId) return chatId
    if (existsSync(this.chatPath(projectId, chatId))) return chatId
    for (const legacy of ProjectStore.legacyChatIdsForFile(filePath)) {
      if (legacy !== chatId && existsSync(this.chatPath(projectId, legacy))) return legacy
    }
    return chatId
  }

  /** Gets the chatId from the mapping; falls back to the path hash without registering. */
  chatIdForPath(filePath: string, projectId?: string): string {
    const index = this.readIndex()
    const key = this.findMapKey(index.chatIdByPath, filePath)
    if (key !== undefined) return index.chatIdByPath![key]!
    return this.fallbackChatId(projectId, filePath)
  }

  /**
   * Resolves { projectId, chatId } by file path (the core of the resolveChat IPC).
   * The chatId is registered into chatIdByPath on first resolve; from then on,
   * renaming/moving the file only changes the mapping key — the chatId stays
   * stable and history always follows the file.
   */
  resolveChatForFile(filePath: string): { projectId: string; chatId: string } {
    const projectId = this.resolveProjectForFile(filePath)
    const index = this.readIndex()
    const key = canonicalPathKey(filePath)
    const mappedKey = this.findMapKey(index.chatIdByPath, filePath)
    if (mappedKey !== undefined) {
      const chatId = index.chatIdByPath![mappedKey]!
      if (mappedKey !== key) {
        ProjectStore.rekeyEntry(index.chatIdByPath!, mappedKey, key)
        this.writeIndex(index)
      }
      return { projectId, chatId }
    }
    const chatId = this.fallbackChatId(projectId, filePath)
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [key]: chatId }
    this.writeIndex(index)
    return { projectId, chatId }
  }

  /**
   * Called after a file is renamed/moved on disk: the keys in fileMap,
   * project.files and chatIdByPath are updated accordingly, while the chatId
   * stays the same (history needs no relocation).
   */
  /** every file path the index keys on (project membership and chat ids) */
  knownFilePaths(): string[] {
    const index = this.readIndex()
    return [...new Set([...Object.keys(index.fileMap), ...Object.keys(index.chatIdByPath ?? {})])]
  }

  fileRenamed(oldPath: string, newPath: string): void {
    if (oldPath === newPath) return
    const index = this.readIndex()
    const oldKey = canonicalPathKey(oldPath)
    const newKey = canonicalPathKey(newPath)
    const pidKey = this.findMapKey(index.fileMap, oldPath)
    if (pidKey !== undefined) {
      const pid = index.fileMap[pidKey]!
      delete index.fileMap[pidKey]
      index.fileMap[newKey] = pid
      const proj = this.readProject(pid)
      if (proj) {
        proj.files = proj.files.map((f) => (canonicalPathKey(f) === oldKey ? newPath : f))
        proj.updatedAt = nowIso()
        this.writeProject(proj)
      }
    }
    // Old data without a mapping: the chatId was derived from the old path hash; register the mapping under that hash on rename so history keeps up
    const chatKey = this.findMapKey(index.chatIdByPath, oldPath)
    const chatId =
      chatKey !== undefined
        ? index.chatIdByPath![chatKey]!
        : this.fallbackChatId(pidKey !== undefined ? index.fileMap[pidKey] : undefined, oldPath)
    if (chatKey !== undefined) delete index.chatIdByPath![chatKey]
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [newKey]: chatId }
    this.writeIndex(index)
  }

  /**
   * Buffer for the opening messages of a chat that has no file yet: the file
   * is not created until the first assistant reply arrives, so aborted or
   * failed requests never leave behind an empty record with a lone user
   * message.
   */
  private readonly pendingFirstWrite = new Map<string, ChatMessage[]>()

  /** Flushes buffered opening messages to disk (materialized before rebind: once the file is saved, the opening messages should be kept). */
  private flushPending(projectId: string, chatId: string): void {
    // Validate before any IO so traversal ids throw instead of being swallowed below
    assertSafeId(projectId, 'projectId')
    assertSafeId(chatId, 'chatId')
    const key = this.seqKey(projectId, chatId)
    const buf = this.pendingFirstWrite.get(key)
    this.pendingFirstWrite.delete(key)
    if (!buf || buf.length === 0) return
    try {
      ensureDir(this.chatsDir(projectId))
      const lines = buf.map((r) => JSON.stringify(r) + '\n').join('')
      appendJsonLines(this.chatPath(projectId, chatId), lines)
    } catch (err) {
      console.warn('[project-store] flushPending failed:', err)
    }
  }

  /**
   * Appends one message to the JSONL. Write failures warn silently, never throw.
   * seq is auto-assigned by the store layer (monotonically increasing).
   * When the record file doesn't exist yet, non-assistant messages are buffered;
   * the file is created and flushed only when the first assistant message arrives.
   */
  appendChatMessage(
    projectId: string,
    chatId: string,
    msg: Omit<ChatMessage, 'seq' | 'ts'> & { ts?: string },
  ): void {
    // Validate ids before the IO try block so traversal attempts throw fail-closed
    assertSafeId(projectId, 'projectId')
    assertSafeId(chatId, 'chatId')
    try {
      const seq = this.nextSeq(projectId, chatId)
      const ts = msg.ts ?? nowIso()
      const text =
        msg.text.length > TEXT_MAX_CHARS
          ? msg.text.slice(0, TEXT_MAX_CHARS) + TEXT_TRUNCATED_MARK
          : msg.text
      const record: ChatMessage = { seq, ts, role: msg.role, text }
      if (msg.fileRef !== undefined) record.fileRef = msg.fileRef
      if (msg.tools && msg.tools.length > 0) {
        // Truncate tool inputs/outputs so one JSONL line can't blow up on a huge payload
        record.tools = msg.tools.map((t) => ({
          ...t,
          ...(t.input !== undefined ? { input: t.input.slice(0, TOOL_FIELD_MAX_CHARS) } : {}),
          ...(t.output !== undefined ? { output: t.output.slice(0, TOOL_FIELD_MAX_CHARS) } : {}),
        }))
      }
      if (msg.attachments !== undefined) record.attachments = msg.attachments
      if (msg.scope !== undefined) {
        record.scope = {
          label: msg.scope.label,
          ...(msg.scope.text !== undefined
            ? { text: msg.scope.text.slice(0, SCOPE_TEXT_MAX_CHARS) }
            : {}),
        }
      }

      const key = this.seqKey(projectId, chatId)
      if (msg.role !== 'assistant' && !existsSync(this.chatPath(projectId, chatId))) {
        const buf = this.pendingFirstWrite.get(key) ?? []
        buf.push(record)
        // Bound the in-memory buffer: overflow materializes the file early
        // instead of dropping user messages.
        if (buf.length >= MAX_PENDING_OPENING_MESSAGES) {
          ensureDir(this.chatsDir(projectId))
          this.pendingFirstWrite.delete(key)
          const lines = buf.map((r) => JSON.stringify(r) + '\n').join('')
          appendJsonLines(this.chatPath(projectId, chatId), lines)
          return
        }
        this.pendingFirstWrite.set(key, buf)
        return
      }
      ensureDir(this.chatsDir(projectId))
      const buf = this.pendingFirstWrite.get(key) ?? []
      this.pendingFirstWrite.delete(key)
      const lines = [...buf, record].map((r) => JSON.stringify(r) + '\n').join('')
      appendJsonLines(this.chatPath(projectId, chatId), lines)
    } catch (err) {
      console.warn('[project-store] appendChatMessage failed:', err)
    }
  }

  /**
   * Reads the most recent `limit` messages (in ascending seq order).
   * A bad JSONL line is skipped without crashing.
   */
  loadChat(projectId: string, chatId: string, limit = DEFAULT_CHAT_LIMIT): ChatMessage[] {
    // Validate ids fail-closed before touching the filesystem
    assertSafeId(projectId, 'projectId')
    assertSafeId(chatId, 'chatId')
    // Clamp limit so 0 no longer returns all messages via slice(-0)
    const safeLimit = normalizeChatLimit(limit)
    const pending = this.pendingFirstWrite.get(this.seqKey(projectId, chatId)) ?? []
    const filePath = this.chatPath(projectId, chatId)
    const messages: ChatMessage[] = [...pending]
    try {
      if (existsSync(filePath)) messages.push(...parseChatRecords(readChatTail(filePath)))
      // Sort by seq and take the most recent entries (safeLimit is always >= 1)
      messages.sort((a, b) => a.seq - b.seq)
      return messages.slice(-safeLimit)
    } catch {
      return messages
    }
  }

  /**
   * Lists metadata of all chats in a project.
   */
  listChats(projectId: string): ChatMeta[] {
    const dir = this.chatsDir(projectId)
    if (!existsSync(dir)) return []
    try {
      const files = readdirSync(dir).filter((f) => f.endsWith('.jsonl'))
      return files.map((f) => {
        const chatId = f.replace(/\.jsonl$/, '')
        const fullPath = join(dir, f)
        let updatedAt = nowIso()
        let approxCount = 0
        try {
          const st = statSync(fullPath)
          updatedAt = st.mtime.toISOString()
          // Estimate: assume an average of 120 bytes per line
          approxCount = Math.round(st.size / 120)
        } catch {
          // use defaults if stat fails
        }
        return { chatId, updatedAt, approxCount }
      })
    } catch {
      return []
    }
  }

  /**
   * Moves chats/<fromId>.jsonl to chats/<toId>.jsonl (possibly across projects).
   * If the target exists, don't overwrite: renumber the source messages' seq and
   * append them at the target's end (old conversations of a same-named file are kept).
   */
  private renameOrMergeChat(
    fromProjectId: string,
    fromId: string,
    toProjectId: string,
    toId: string,
  ): void {
    if (fromProjectId === toProjectId && fromId === toId) return
    // The source may still have buffered opening messages: materialize them first (once the file is saved, they should be kept)
    this.flushPending(fromProjectId, fromId)
    const oldPath = this.chatPath(fromProjectId, fromId)
    const newPath = this.chatPath(toProjectId, toId)
    let mergedMaxSeq: number | undefined
    try {
      if (existsSync(oldPath)) {
        ensureDir(dirname(newPath))
        if (!existsSync(newPath)) {
          renameSync(oldPath, newPath)
        } else {
          mergedMaxSeq = mergeChatFiles(oldPath, newPath)
        }
      }
    } catch (err) {
      console.warn('[project-store] rebindChat rename failed:', err)
    }

    // Migrate the seq counter (when merged, the renumbered max seq wins)
    const oldKey = this.seqKey(fromProjectId, fromId)
    const curSeq = this.seqCounters.get(oldKey)
    this.seqCounters.delete(oldKey)
    this.seqCounters.delete(this.seqKey(toProjectId, toId))
    const next = mergedMaxSeq ?? curSeq
    if (next !== undefined) {
      this.seqCounters.set(this.seqKey(toProjectId, toId), next)
    }
  }

  /**
   * Renames the JSONL file (called after an unsaved file is first written to disk):
   * chats/<tempId>.jsonl → chats/<newChatId>.jsonl (if the target exists, merge by continuing seq).
   */
  rebindChat(projectId: string, tempId: string, newChatId: string): void {
    this.renameOrMergeChat(projectId, tempId, projectId, newChatId)
  }

  /**
   * After an unsaved session first gets a real file path: register in fileMap,
   * compute the chatId from the path, and move the temp JSONL into the target project.
   * Returns the new { projectId, chatId } for the renderer to update its references.
   */
  rebindChatToFile(
    tempProjectId: string,
    tempChatId: string,
    filePath: string,
  ): { projectId: string; chatId: string } {
    const { projectId, chatId } = this.resolveChatForFile(filePath)
    this.renameOrMergeChat(tempProjectId, tempChatId, projectId, chatId)
    return { projectId, chatId }
  }

  /**
   * Gets project info.
   */
  getProject(projectId: string): ProjectData | null {
    return this.readProject(projectId)
  }

  /**
   * Lists all projects.
   */
  listProjects(): ProjectInfo[] {
    return this.readIndex().projects
  }

  // ── P1 extended API ────────────────────────────────────────

  /**
   * Lists all projects (with file count + last active time).
   */
  listProjectsSummary(): ProjectSummary[] {
    this.ensureDefaultProject()
    const index = this.readIndex()
    return index.projects.map((info) => {
      const fileCount = this.listProjectFiles(info.id).length
      // Take the max of project.json updatedAt and all chat mtimes
      let lastActiveAt = info.updatedAt
      const chats = this.listChats(info.id)
      for (const c of chats) {
        if (c.updatedAt > lastActiveAt) lastActiveAt = c.updatedAt
      }
      return {
        ...info,
        fileCount,
        lastActiveAt,
        isDefault: info.id === 'default',
      }
    })
  }

  /**
   * Lists files that currently exist for a project. Stored paths are historical
   * records and may outlive files deleted or moved outside GenOffice.
   */
  listProjectFiles(projectId: string): string[] {
    const proj = this.readProject(projectId)
    if (!proj) return []
    return [...new Set(proj.files)].filter((filePath) => existsSync(filePath))
  }

  /**
   * Creates a project (name must be non-empty; id is the first 12 hex chars of
   * sha256(name) plus a timestamp suffix to avoid collisions).
   * Returns the newly created ProjectData.
   */
  createProject(name: string): ProjectData {
    const trimmed = name.trim()
    if (!trimmed) throw new Error('Project name cannot be empty')
    if (trimmed.length > MAX_PROJECT_NAME_CHARS) {
      throw new Error(
        `Project name too long: ${trimmed.length} chars (max ${MAX_PROJECT_NAME_CHARS})`,
      )
    }
    const now = nowIso()
    // Generate a stable yet unique id
    const hash = createHash('sha256')
      .update(trimmed + now)
      .digest('hex')
      .slice(0, 12)
    const id = `proj-${hash}`
    const data: ProjectData = {
      id,
      name: trimmed,
      createdAt: now,
      updatedAt: now,
      files: [],
    }
    ensureDir(this.projectDir(id))
    this.writeProject(data)
    const index = this.readIndex()
    // Append at the end (default always stays first)
    index.projects.push({ id, name: trimmed, createdAt: now, updatedAt: now })
    this.writeIndex(index)
    return data
  }

  /**
   * Renames a project (the default project cannot be renamed).
   */
  renameProject(id: string, name: string): void {
    if (id === 'default') throw new Error('The default project cannot be renamed')
    const trimmed = name.trim()
    if (!trimmed) throw new Error('Project name cannot be empty')
    if (trimmed.length > MAX_PROJECT_NAME_CHARS) {
      throw new Error(
        `Project name too long: ${trimmed.length} chars (max ${MAX_PROJECT_NAME_CHARS})`,
      )
    }
    const now = nowIso()
    const proj = this.readProject(id)
    if (!proj) throw new Error(`Project does not exist: ${id}`)
    proj.name = trimmed
    proj.updatedAt = now
    this.writeProject(proj)
    const index = this.readIndex()
    const entry = index.projects.find((p) => p.id === id)
    if (entry) {
      entry.name = trimmed
      entry.updatedAt = now
    }
    this.writeIndex(index)
  }

  /**
   * Soft-deletes a project:
   * 1. Move each of its files' chats into the default project, so the transcript
   *    follows the file instead of staying behind in the trashed directory
   * 2. Move the directory into projects/.trash/<id>-<ts>/
   * 3. Reassign all of its files in fileMap back to default
   * 4. Remove the project from index.projects
   * The default project cannot be deleted.
   */
  deleteProject(id: string): void {
    if (id === 'default') throw new Error('The default project cannot be deleted')
    const proj = this.readProject(id)
    if (!proj) throw new Error(`Project does not exist: ${id}`)

    this.ensureDefaultProject()
    const index = this.readIndex()
    const ownedFiles = Object.entries(index.fileMap)
      .filter(([, pid]) => pid === id)
      .map(([filePath]) => filePath)

    // 1. Migrate the chats first: the transcript has to be readable from the
    // default project before the directory it lives in is moved to the trash.
    for (const filePath of ownedFiles) {
      const chatId = this.chatIdForPath(filePath, id)
      this.renameOrMergeChat(id, chatId, 'default', chatId)
    }

    // 2. Soft-delete the directory
    const src = this.projectDir(id)
    const ts = Date.now()
    const trashDir = join(this.baseDir, '.trash')
    ensureDir(trashDir)
    const dst = join(trashDir, `${id}-${ts}`)
    try {
      if (existsSync(src)) renameSync(src, dst)
    } catch (err) {
      console.warn('[project-store] deleteProject rename to trash failed:', err)
    }

    // 3. Reassign this project's files in fileMap back to default
    const movedFiles: string[] = []
    for (const [filePath, pid] of Object.entries(index.fileMap)) {
      if (pid === id) {
        index.fileMap[filePath] = 'default'
        movedFiles.push(filePath)
      }
    }
    // Update the default project.json, listing each file under the spelling the
    // deleted project showed it as rather than under its canonical map key
    if (movedFiles.length > 0) {
      const defaultProj = this.readProject('default')
      if (defaultProj) {
        for (const key of movedFiles) {
          const f = proj.files.find((listed) => canonicalPathKey(listed) === key) ?? key
          if (!ProjectStore.ownsFile(defaultProj.files, f)) defaultProj.files.push(f)
        }
        defaultProj.updatedAt = nowIso()
        this.writeProject(defaultProj)
      }
    }

    // 4. Remove the index.projects entry
    index.projects = index.projects.filter((p) => p.id !== id)
    this.writeIndex(index)
  }

  /**
   * Moves a file from its current project into a target project:
   * 1. Update fileMap
   * 2. Update the files lists in both project.json files
   * 3. Move the corresponding chat's jsonl file to the new project directory
   */
  moveFileToProject(filePath: string, targetProjectId: string): void {
    this.ensureDefaultProject()
    const index = this.readIndex()
    const existingKey = this.findMapKey(index.fileMap, filePath)
    const fromProjectId = existingKey !== undefined ? index.fileMap[existingKey]! : 'default'
    const newKey = canonicalPathKey(filePath)

    if (fromProjectId === targetProjectId) return // nothing to move

    // The target project must exist
    const targetProj = this.readProject(targetProjectId)
    if (!targetProj) throw new Error(`Target project does not exist: ${targetProjectId}`)

    // 1. Update fileMap
    if (existingKey !== undefined && existingKey !== newKey) delete index.fileMap[existingKey]
    index.fileMap[newKey] = targetProjectId
    this.writeIndex(index)

    // 2. Update fromProject.files
    const fromProj = this.readProject(fromProjectId)
    if (fromProj) {
      fromProj.files = fromProj.files.filter(
        (f) => canonicalPathKey(f) !== canonicalPathKey(filePath),
      )
      fromProj.updatedAt = nowIso()
      this.writeProject(fromProj)
    }

    // 3. Update targetProject.files
    if (!ProjectStore.ownsFile(targetProj.files, filePath)) targetProj.files.push(filePath)
    targetProj.updatedAt = nowIso()
    this.writeProject(targetProj)

    // 4. Move the corresponding chat's JSONL (materialize buffered opening messages first)
    const chatId = this.chatIdForPath(filePath, fromProjectId)
    this.flushPending(fromProjectId, chatId)
    const srcChatPath = this.chatPath(fromProjectId, chatId)
    const dstChatPath = this.chatPath(targetProjectId, chatId)
    try {
      if (existsSync(srcChatPath)) {
        ensureDir(this.chatsDir(targetProjectId))
        renameSync(srcChatPath, dstChatPath)
      }
    } catch (err) {
      console.warn('[project-store] moveFileToProject chat rename failed:', err)
    }

    // 5. Migrate the seq counter cache
    const oldKey = this.seqKey(fromProjectId, chatId)
    const movedSeqKey = this.seqKey(targetProjectId, chatId)
    const cur = this.seqCounters.get(oldKey)
    this.seqCounters.delete(oldKey)
    if (cur !== undefined) this.seqCounters.set(movedSeqKey, cur)
  }

  /**
   * Aggregates messages from all chats in a project, sorted by ts descending,
   * returning the most recent `limit` entries. Each entry includes the file path
   * (reverse-looked-up from fileMap by chatId), role, and preview text.
   */
  getProjectTimeline(projectId: string, limit = 20): TimelineEntry[] {
    const boundedLimit = normalizeTimelineLimit(limit)
    const index = this.readIndex()
    // Build the reverse chatId → filePath map (files in this project only); mapping wins, old data falls back to the path hash
    const chatToFile = new Map<string, string>()
    for (const [filePath, pid] of Object.entries(index.fileMap)) {
      if (pid === projectId) {
        const chatId = this.fallbackChatId(projectId, filePath)
        const chatKey = this.findMapKey(index.chatIdByPath, filePath)
        chatToFile.set(chatKey !== undefined ? index.chatIdByPath![chatKey]! : chatId, filePath)
      }
    }

    const entries: TimelineEntry[] = []
    const chats = this.listChats(projectId)
    for (const { chatId } of chats) {
      const filePath = chatToFile.get(chatId) ?? ''
      const msgs = this.loadChat(projectId, chatId, 200)
      for (const msg of msgs) {
        entries.push({
          filePath,
          fileName: filePath ? basename(filePath) : chatId,
          chatId,
          ts: msg.ts,
          role: msg.role,
          preview: msg.text.slice(0, 120),
          seq: msg.seq,
        })
      }
    }

    // Sort by ts descending
    entries.sort((a, b) => {
      if (b.ts > a.ts) return 1
      if (b.ts < a.ts) return -1
      return b.seq - a.seq
    })
    return entries.slice(0, boundedLimit)
  }
}
