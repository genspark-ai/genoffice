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

import { createHash } from 'node:crypto'
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
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
const TEXT_TRUNCATED_MARK = '\n\n[truncated]'

function nowIso(): string {
  return new Date().toISOString()
}

function ensureDir(dir: string): void {
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true })
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
    return join(this.baseDir, projectId)
  }

  private projectJsonPath(projectId: string): string {
    return join(this.projectDir(projectId), 'project.json')
  }

  private chatsDir(projectId: string): string {
    return join(this.projectDir(projectId), 'chats')
  }

  private chatPath(projectId: string, chatId: string): string {
    return join(this.chatsDir(projectId), `${chatId}.jsonl`)
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
    const existing = index.fileMap[filePath]
    if (existing) return existing

    // Assign to default
    index.fileMap[filePath] = 'default'
    this.writeIndex(index)

    // Update the files list in project.json
    const proj = this.readProject('default')
    if (proj && !proj.files.includes(filePath)) {
      proj.files.push(filePath)
      proj.updatedAt = nowIso()
      this.writeProject(proj)
    }
    return 'default'
  }

  /**
   * Derives a chatId from a file path (first 16 hex chars of sha256).
   * Unsaved files use an externally provided temp id (e.g. "unsaved-<timestamp>").
   * Only a fallback derivation for old data without a mapping; new code uses
   * resolveChatForFile (stable mapping).
   */
  static chatIdForFile(filePath: string): string {
    return createHash('sha256').update(filePath).digest('hex').slice(0, 16)
  }

  /** Gets the chatId from the mapping; falls back to the path hash without registering. */
  chatIdForPath(filePath: string): string {
    const index = this.readIndex()
    return index.chatIdByPath?.[filePath] ?? ProjectStore.chatIdForFile(filePath)
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
    const mapped = index.chatIdByPath?.[filePath]
    if (mapped) return { projectId, chatId: mapped }
    const chatId = ProjectStore.chatIdForFile(filePath)
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [filePath]: chatId }
    this.writeIndex(index)
    return { projectId, chatId }
  }

  /**
   * Called after a file is renamed/moved on disk: the keys in fileMap,
   * project.files and chatIdByPath are updated accordingly, while the chatId
   * stays the same (history needs no relocation).
   */
  fileRenamed(oldPath: string, newPath: string): void {
    if (oldPath === newPath) return
    const index = this.readIndex()
    const pid = index.fileMap[oldPath]
    if (pid !== undefined) {
      delete index.fileMap[oldPath]
      index.fileMap[newPath] = pid
      const proj = this.readProject(pid)
      if (proj) {
        proj.files = proj.files.map((f) => (f === oldPath ? newPath : f))
        proj.updatedAt = nowIso()
        this.writeProject(proj)
      }
    }
    // Old data without a mapping: the chatId was derived from the old path hash; register the mapping under that hash on rename so history keeps up
    const chatId = index.chatIdByPath?.[oldPath] ?? ProjectStore.chatIdForFile(oldPath)
    if (index.chatIdByPath?.[oldPath] !== undefined) delete index.chatIdByPath[oldPath]
    index.chatIdByPath = { ...(index.chatIdByPath ?? {}), [newPath]: chatId }
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
    const key = this.seqKey(projectId, chatId)
    const buf = this.pendingFirstWrite.get(key)
    this.pendingFirstWrite.delete(key)
    if (!buf || buf.length === 0) return
    try {
      ensureDir(this.chatsDir(projectId))
      const lines = buf.map((r) => JSON.stringify(r) + '\n').join('')
      appendFileSync(this.chatPath(projectId, chatId), lines, 'utf8')
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

      const key = this.seqKey(projectId, chatId)
      if (msg.role !== 'assistant' && !existsSync(this.chatPath(projectId, chatId))) {
        const buf = this.pendingFirstWrite.get(key) ?? []
        buf.push(record)
        this.pendingFirstWrite.set(key, buf)
        return
      }
      ensureDir(this.chatsDir(projectId))
      const buf = this.pendingFirstWrite.get(key) ?? []
      this.pendingFirstWrite.delete(key)
      const lines = [...buf, record].map((r) => JSON.stringify(r) + '\n').join('')
      appendFileSync(this.chatPath(projectId, chatId), lines, 'utf8')
    } catch (err) {
      console.warn('[project-store] appendChatMessage failed:', err)
    }
  }

  /**
   * Reads the most recent `limit` messages (in ascending seq order).
   * A bad JSONL line is skipped without crashing.
   */
  loadChat(projectId: string, chatId: string, limit = 200): ChatMessage[] {
    const pending = this.pendingFirstWrite.get(this.seqKey(projectId, chatId)) ?? []
    const filePath = this.chatPath(projectId, chatId)
    const messages: ChatMessage[] = [...pending]
    try {
      if (existsSync(filePath)) {
        const raw = readFileSync(filePath, 'utf8')
        const lines = raw.split('\n').filter((l) => l.trim())
        for (const line of lines) {
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
            // skip bad lines
          }
        }
      }
      // Sort by seq and take the most recent `limit` entries
      messages.sort((a, b) => a.seq - b.seq)
      return messages.slice(-limit)
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
          const existing = this.loadChat(toProjectId, toId, 10_000)
          let seq = existing.reduce((m, msg) => Math.max(m, msg.seq), -1) + 1
          const moved = this.loadChat(fromProjectId, fromId, 10_000)
          const lines = moved.map((m) => JSON.stringify({ ...m, seq: seq++ }) + '\n').join('')
          if (lines) appendFileSync(newPath, lines, 'utf8')
          unlinkSync(oldPath)
          mergedMaxSeq = seq - 1
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
   * 1. Move the directory into projects/.trash/<id>-<ts>/
   * 2. Reassign all of its files in fileMap back to default
   * 3. Remove the project from index.projects
   * The default project cannot be deleted.
   */
  deleteProject(id: string): void {
    if (id === 'default') throw new Error('The default project cannot be deleted')
    const proj = this.readProject(id)
    if (!proj) throw new Error(`Project does not exist: ${id}`)

    // 1. Soft-delete the directory
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

    // 2. Reassign this project's files in fileMap back to default
    this.ensureDefaultProject()
    const index = this.readIndex()
    const movedFiles: string[] = []
    for (const [filePath, pid] of Object.entries(index.fileMap)) {
      if (pid === id) {
        index.fileMap[filePath] = 'default'
        movedFiles.push(filePath)
      }
    }
    // Update the default project.json
    if (movedFiles.length > 0) {
      const defaultProj = this.readProject('default')
      if (defaultProj) {
        for (const f of movedFiles) {
          if (!defaultProj.files.includes(f)) defaultProj.files.push(f)
        }
        defaultProj.updatedAt = nowIso()
        this.writeProject(defaultProj)
      }
    }

    // 3. Remove the index.projects entry
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
    const fromProjectId = index.fileMap[filePath] ?? 'default'

    if (fromProjectId === targetProjectId) return // nothing to move

    // The target project must exist
    const targetProj = this.readProject(targetProjectId)
    if (!targetProj) throw new Error(`Target project does not exist: ${targetProjectId}`)

    // 1. Update fileMap
    index.fileMap[filePath] = targetProjectId
    this.writeIndex(index)

    // 2. Update fromProject.files
    const fromProj = this.readProject(fromProjectId)
    if (fromProj) {
      fromProj.files = fromProj.files.filter((f) => f !== filePath)
      fromProj.updatedAt = nowIso()
      this.writeProject(fromProj)
    }

    // 3. Update targetProject.files
    if (!targetProj.files.includes(filePath)) targetProj.files.push(filePath)
    targetProj.updatedAt = nowIso()
    this.writeProject(targetProj)

    // 4. Move the corresponding chat's JSONL (materialize buffered opening messages first)
    const chatId = this.chatIdForPath(filePath)
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
    const newKey = this.seqKey(targetProjectId, chatId)
    const cur = this.seqCounters.get(oldKey)
    this.seqCounters.delete(oldKey)
    if (cur !== undefined) this.seqCounters.set(newKey, cur)
  }

  /**
   * Aggregates messages from all chats in a project, sorted by ts descending,
   * returning the most recent `limit` entries. Each entry includes the file path
   * (reverse-looked-up from fileMap by chatId), role, and preview text.
   */
  getProjectTimeline(projectId: string, limit = 20): TimelineEntry[] {
    const index = this.readIndex()
    // Build the reverse chatId → filePath map (files in this project only); mapping wins, old data falls back to the path hash
    const chatToFile = new Map<string, string>()
    for (const [filePath, pid] of Object.entries(index.fileMap)) {
      if (pid === projectId) {
        const chatId = index.chatIdByPath?.[filePath] ?? ProjectStore.chatIdForFile(filePath)
        chatToFile.set(chatId, filePath)
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
    return entries.slice(0, limit)
  }
}
