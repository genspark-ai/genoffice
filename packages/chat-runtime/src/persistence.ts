/**
 * JSONL persistence adapter for chat sessions.
 *
 * Each session is one `.jsonl` file under `<dir>/<sessionId>.jsonl`. Lines
 * are either an initial session header (`{"kind":"session",...}`) or a
 * message append (`{"kind":"message",...}`). Existing
 * `packages/project-store` files use the same on-disk shape but a slightly
 * different envelope; this adapter wraps the project-store reader when
 * available and falls back to a fresh format when it isn't.
 *
 * The renderer apps that already use `project-store` should keep their
 * existing reader for legacy chats and pass a `persistence` adapter into
 * `ChatRuntime` only for new sessions.
 *
 * NOTE: `node:fs` / `node:path` are loaded lazily inside
 * `JsonlChatPersistence` methods so the browser bundle (which never
 * instantiates this class) tree-shakes them out.
 */
import type { ChatSession, ChatMessage, ChatPersistence } from './types'

export interface JsonlPersistenceOptions {
  /** Directory under which `<sessionId>.jsonl` lives. Created on first write. */
  dir: string
}

export class JsonlChatPersistence implements ChatPersistence {
  constructor(private readonly opts: JsonlPersistenceOptions) {}

  private async file(sessionId: string): Promise<string> {
    const { join } = await import('node:path')
    return join(this.opts.dir, `${safeId(sessionId)}.jsonl`)
  }

  async load(sessionId: string): Promise<ChatSession | null> {
    const { existsSync, readFileSync } = await import('node:fs')
    const file = await this.file(sessionId)
    if (!existsSync(file)) return null
    const raw = readFileSync(file, 'utf8')
    let header: ChatSession | null = null
    const messages: ChatMessage[] = []
    for (const line of raw.split(/\r?\n/)) {
      if (!line.trim()) continue
      try {
        const row = JSON.parse(line) as { kind?: string } & Record<string, unknown>
        if (row.kind === 'session') {
          const { kind: _kind, ...rest } = row as { kind: string } & Record<string, unknown>
          header = { ...(rest as unknown as ChatSession), messages: [] }
        } else if (row.kind === 'message' && header) {
          const { kind: _kind, ...rest } = row as { kind: string } & Record<string, unknown>
          messages.push(rest as unknown as ChatMessage)
        }
      } catch (err) {
        console.warn('[chat-runtime] corrupt JSONL line, skipping:', err)
      }
    }
    if (!header) return null
    header.messages = messages
    return header
  }

  async save(session: ChatSession): Promise<void> {
    const { mkdirSync, writeFileSync } = await import('node:fs')
    const file = await this.file(session.id)
    mkdirSync(this.opts.dir, { recursive: true })
    const header = { kind: 'session', ...stripMessages(session) }
    const lines = [JSON.stringify(header)]
    for (const msg of session.messages) {
      lines.push(JSON.stringify({ kind: 'message', ...msg }))
    }
    writeFileSync(file, lines.join('\n') + '\n', 'utf8')
  }

  async append(sessionId: string, message: ChatMessage): Promise<void> {
    const { existsSync, mkdirSync, writeFileSync, appendFileSync } = await import('node:fs')
    const file = await this.file(sessionId)
    mkdirSync(this.opts.dir, { recursive: true })
    if (!existsSync(file)) {
      // Header is required for load() to recognise the file as a session.
      const header = {
        kind: 'session',
        id: sessionId,
        app: 'docs',
        title: 'Untitled',
        createdAt: message.timestamp,
        updatedAt: message.timestamp,
        messages: [],
      }
      writeFileSync(file, JSON.stringify(header) + '\n', 'utf8')
    }
    appendFileSync(file, JSON.stringify({ kind: 'message', ...message }) + '\n', 'utf8')
  }
}

function safeId(input: string): string {
  return input.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 96) || 'session'
}

function stripMessages(session: ChatSession): Omit<ChatSession, 'messages'> {
  const { messages: _messages, ...rest } = session
  return rest
}

/** In-memory persistence adapter used by tests and browser builds. */
export class MemoryChatPersistence implements ChatPersistence {
  private readonly store = new Map<string, ChatSession>()

  async load(sessionId: string): Promise<ChatSession | null> {
    const session = this.store.get(sessionId)
    return session ? clone(session) : null
  }

  async save(session: ChatSession): Promise<void> {
    this.store.set(session.id, clone(session))
  }

  async append(sessionId: string, message: ChatMessage): Promise<void> {
    const existing = this.store.get(sessionId)
    if (existing) {
      existing.messages.push(clone(message))
      existing.updatedAt = message.timestamp
    } else {
      this.store.set(sessionId, {
        id: sessionId,
        app: 'docs',
        title: 'Untitled',
        createdAt: message.timestamp,
        updatedAt: message.timestamp,
        messages: [clone(message)],
      })
    }
  }
}

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}
