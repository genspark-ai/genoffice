/**
 * `ChatRuntime` — the single mount point for any app's AI panel.
 *
 * Owns: session state, the current/last run, an event stream the renderer
 * can subscribe to, and helpers for cancel / snapshot / restore.
 *
 * The renderer dispatches tool lifecycle events into the runtime via the
 * `bridge()` method; this lets each app keep using its own tool executor
 * (docs engine, xlsx gateway, slide master, pdf ops) without exposing the
 * details here.
 */

import { classifyError } from './errors'
import { loadOrCreateSession, appendMessage } from './session'
import { startRun } from './run'
import type { RunBridge } from './run'
import type {
  ChatCapability,
  ChatChangePlan,
  ChatEvent,
  ChatEventListener,
  ChatMessage,
  ChatPersistence,
  ChatRun,
  ChatRuntimeOptions,
  ChatSendOptions,
  ChatSession,
  ChatSnapshot,
  ChatToolCallRecord,
} from './types'

const userMessageCounter = { n: 0 }
const toolCounter = { n: 0 }
function newMessageId(): string {
  userMessageCounter.n += 1
  return `msg-${Date.now().toString(36)}-${userMessageCounter.n.toString(36)}`
}
function newToolId(): string {
  toolCounter.n += 1
  return `tool-${Date.now().toString(36)}-${toolCounter.n.toString(36)}`
}

export class ChatRuntime {
  private session: ChatSession
  private currentRun: ChatRun | null = null
  private listeners = new Set<ChatEventListener>()
  private handle: ReturnType<typeof startRun> | null = null
  private options: ChatRuntimeOptions
  private persistence: ChatPersistence | undefined
  private changePlanResolver: ((plan: ChatChangePlan | null) => void) | null = null

  constructor(options: ChatRuntimeOptions) {
    this.options = options
    this.persistence = options.persistence
    this.session = createInitialSession(options)
  }

  /** Synchronously initialise; call `init()` to load persisted messages. */
  static create(options: ChatRuntimeOptions): ChatRuntime {
    return new ChatRuntime(options)
  }

  async init(): Promise<void> {
    if (!this.persistence) return
    const loaded = await this.persistence.load(this.session.id)
    if (loaded) {
      this.session = loaded
      this.emit({ type: 'session-updated', session: this.session })
    }
  }

  getSession(): ChatSession {
    return this.session
  }

  getCurrentRun(): ChatRun | null {
    return this.currentRun
  }

  subscribe(listener: ChatEventListener): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  private emit(event: ChatEvent): void {
    for (const l of this.listeners) {
      try {
        l(event)
      } catch (err) {
        console.warn('[chat-runtime] listener threw:', err)
      }
    }
  }

  /** Begin a new run. Returns a promise that resolves with the final run. */
  async send(opts: ChatSendOptions): Promise<ChatRun> {

// DEBUG MARKER: send() entered

    if (this.currentRun && !this.currentRun.cancelled && this.currentRun.status !== 'done' && this.currentRun.status !== 'error') {
      throw new Error('A run is already in progress; cancel it before starting another.')
    }

    const userMessage: ChatMessage = {
      id: newMessageId(),
      role: 'user',
      content: opts.instruction,
      attachments: opts.attachments,
      timestamp: Date.now(),
      status: 'done',
    }
    this.session = appendMessage(this.session, userMessage)
    await this.persistMessage(userMessage)
    this.emit({ type: 'session-updated', session: this.session })

    const assistantMessage: ChatMessage = {
      id: newMessageId(),
      role: 'assistant',
      content: '',
      timestamp: Date.now(),
      status: 'streaming',
      streaming: true,
      tools: [],
    }
    this.session = appendMessage(this.session, assistantMessage)
    await this.persistMessage(assistantMessage)
    this.emit({ type: 'session-updated', session: this.session })

    return new Promise<ChatRun>(resolve => {
      this.handle = startRun({
        sessionId: this.session.id,
        skill: this.options.skill,
        captureSnapshot: this.options.captureSnapshot,
        systemSuffix: this.options.systemSuffix,
        formatUserMessage: this.options.formatUserMessage,
        maxTurns: this.options.maxTurns,
        onChangePlan: plan => {
          this.emit({ type: 'run-change-plan', runId: this.handle?.run.id ?? '', plan })
          if (this.changePlanResolver) {
            const resolver = this.changePlanResolver
            this.changePlanResolver = null
            resolver(plan)
          }
        },
        onUpdate: run => {
          this.currentRun = run
          assistantMessage.content = run.streamingText
          assistantMessage.tools = run.tools
          assistantMessage.error = run.error?.message
          assistantMessage.status = run.status === 'done' ? 'done' : run.status === 'error' ? 'error' : 'streaming'
          assistantMessage.streaming = run.status === 'running' || run.status === 'streaming'
          if (run.status === 'done' || run.status === 'error' || run.status === 'cancelled') {
            assistantMessage.streaming = false
          }
          if (run.status === 'streaming' || run.status === 'running') {
            // Map the current run as the active assistant message update.
          }
          if (run.status === 'done' || run.status === 'cancelled') {
            this.emit({ type: 'run-finish', run })
            resolve(run)
          } else if (run.status === 'error') {
            this.emit({ type: 'run-error', run })
            resolve(run)
          }
        },
      })
      this.emit({ type: 'run-start', run: this.handle.run })
    })
  }

  cancel(): void {
    if (!this.handle) return
    this.handle.cancel()
  }

  /** Snapshot for rollback UIs (delegates to the in-flight run). */
  snapshot(): ChatSnapshot | null {
    if (!this.handle) return null
    const state = this.handle.snapshot()
    if (state === undefined) return null
    return { runId: this.handle.run.id, capturedAt: Date.now(), state }
  }

  /**
   * Bridge the renderer's tool executor into the runtime. Calling this
   * while a run is active feeds tool lifecycle events into the chat
   * timeline. The renderer is responsible for actually executing the
   * tool and producing the output.
   */
  bridge(): RunBridge {
    const run = this.handle?.run
    if (!run || !this.handle) {
      throw new Error('No active run to bridge into')
    }
    return {
      emitToolStart: call => {
        const rec: ChatToolCallRecord = {
          id: call.id || newToolId(),
          name: call.name,
          input: call.input,
          status: 'running',
          startedAt: Date.now(),
        }
        if (!run.tools.find(t => t.id === rec.id)) run.tools.push(rec)
        this.emit({ type: 'run-tool-start', runId: run.id, tool: rec })
      },
      emitToolExecuted: (call, execution) => {
        const rec = run.tools.find(t => t.id === call.id)
        if (rec) {
          rec.status = execution.isError ? 'error' : 'executed'
          rec.output = execution.output
          rec.isError = execution.isError
          rec.finishedAt = Date.now()
        }
        this.emit({ type: 'run-tool-executed', runId: run.id, tool: rec ?? {
          id: call.id,
          name: call.name,
          input: call.input,
          status: execution.isError ? 'error' : 'executed',
          output: execution.output,
          isError: execution.isError,
          startedAt: Date.now(),
          finishedAt: Date.now(),
        } })
      },
      emitText: (delta, cumulative) => {
        run.streamingText = cumulative
        this.emit({ type: 'run-text', runId: run.id, delta, cumulative })
      },
      emitDone: result => {
        run.cancelled = result.cancelled
        run.status = result.cancelled ? 'cancelled' : 'done'
        run.finishedAt = Date.now()
        this.emit({ type: 'run-finish', run })
      },
      emitError: message => {
        run.error = classifyError(message)
        run.status = 'error'
        run.finishedAt = Date.now()
        this.emit({ type: 'run-error', run })
      },
    }
  }

  /**
   * Apply a `ChatChangePlan` (caller resolves any confirm dialog first).
   * The default implementation just emits `run-finish`; apps override by
   * subscribing to `run-change-plan` and using their own applier.
   */
  applyChangePlan(plan: ChatChangePlan): void {
    this.emit({ type: 'run-finish', run: this.currentRun ?? this.handle!.run })
    void plan
  }

  rejectChangePlan(plan: ChatChangePlan): void {
    this.emit({ type: 'run-finish', run: this.currentRun ?? this.handle!.run })
    void plan
  }

  /** Wait for the next `change-plan` event (or null when the run ends without one). */
  waitForChangePlan(timeoutMs = 60_000): Promise<ChatChangePlan | null> {
    return new Promise(resolve => {
      const timer = setTimeout(() => {
        this.changePlanResolver = null
        resolve(null)
      }, timeoutMs)
      this.changePlanResolver = plan => {
        clearTimeout(timer)
        resolve(plan)
      }
    })
  }

  private async persistMessage(message: ChatMessage): Promise<void> {
    if (!this.persistence) return
    try {
      if (this.persistence.append) {
        await this.persistence.append(this.session.id, message)
      } else {
        await this.persistence.save(this.session)
      }
    } catch (err) {
      console.warn('[chat-runtime] failed to persist message:', err)
    }
  }
}

function createInitialSession(options: ChatRuntimeOptions): ChatSession {
  return {
    id: options.sessionId,
    app: options.app,
    title: 'New Chat',
    createdAt: Date.now(),
    updatedAt: Date.now(),
    messages: options.initialMessages ? [...options.initialMessages] : [],
    capability: options.capability,
  }
}

// Re-export so apps don't need a second import line.
export { loadOrCreateSession }
