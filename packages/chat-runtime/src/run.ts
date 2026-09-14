/**
 * Run lifecycle — wraps `AgentLoop` so the rest of the suite sees a single
 * `ChatRun` interface with `status`, `streamingText`, `tools`, `error`,
 * and a typed `snapshot()` / `restore()` for rollback UIs.
 *
 * All existing `AgentLoopEvents` callbacks (`onText`, `onToolStart`,
 * `onToolExecuted`, `onDone`, `onError`) keep their original semantics so
 * the four AiPanel renderers can subscribe without re-learning.
 */

import {
  AgentLoop,
  type AgentLoopEvents,
  type AgentRunResult,
  type AgentSkill,
  type AgentToolCall,
  type ToolExecution,
} from '@genoffice/agent-core'

import { classifyError } from './errors.js'
import type {
  ChatRun,
  ChatRunStatus,
  ChatToolCallRecord,
  ChatChangePlan,
} from './types.js'

export interface RunHandle {
  run: ChatRun
  cancel(): void
  snapshot(): unknown | undefined
}

export interface StartRunOptions {
  sessionId: string
  skill: AgentSkill
  captureSnapshot?: (() => unknown) | undefined
  systemSuffix?: (() => string) | undefined
  formatUserMessage?: ((instruction: string, context: string) => string) | undefined
  maxTurns?: number | undefined
  /** Hook invoked when a tool emits a `ChangePlan` payload (via the `kind:'change-plan'` tool display). */
  onChangePlan?: ((plan: ChatChangePlan) => void) | undefined
  onUpdate?: ((run: ChatRun) => void) | undefined
}

let counter = 0
function nextRunId(): string {
  counter += 1
  return `run-${Date.now().toString(36)}-${counter.toString(36)}`
}

export function startRun(opts: StartRunOptions): RunHandle {
  let status: ChatRunStatus = 'queued'
  const tools: ChatToolCallRecord[] = []
  let streamingText = ''
  let lastChangePlan: ChatChangePlan | undefined
  let snapshotBeforeFirstTool: unknown | undefined
  let cancelled = false

  const emit = (): void => {
    opts.onUpdate?.({
      id: runId,
      sessionId: opts.sessionId,
      status,
      startedAt,
      streamingText,
      tools: [...tools],
      lastChangePlan,
      cancelled: status === 'cancelled' || cancelled,
    })
  }

  const runId = nextRunId()
  const startedAt = Date.now()

  const events: AgentLoopEvents<unknown> = {
    onText(text) {
      streamingText = text
      status = 'streaming'
      emit()
    },
    onToolStart(call: AgentToolCall) {
      status = 'running'
      tools.push({
        id: call.id,
        name: call.name,
        input: call.input,
        status: 'running',
        startedAt: Date.now(),
      })
      if (snapshotBeforeFirstTool === undefined && opts.captureSnapshot) {
        snapshotBeforeFirstTool = opts.captureSnapshot()
      }
      emit()
    },
    onToolExecuted(event) {
      const rec = tools.find(t => t.id === event.call.id)
      const exec: ToolExecution = event.execution
      if (rec) {
        rec.status = exec.isError ? 'error' : 'executed'
        rec.output = exec.output
        rec.isError = exec.isError
        rec.finishedAt = Date.now()
        if (event.snapshotBefore !== undefined) {
          rec.snapshotBefore = event.snapshotBefore
        } else if (snapshotBeforeFirstTool !== undefined && rec.snapshotBefore === undefined) {
          rec.snapshotBefore = snapshotBeforeFirstTool
        }
      }
      const display = (exec as { display?: unknown }).display
      if (display) {
        const plan = extractChangePlan(display)
        if (plan) {
          lastChangePlan = plan
          opts.onChangePlan?.(plan)
        }
      }
      emit()
    },
    onTurnEnd() {
      // a turn requested tools and they ran; back to the model.
    },
    onDone(result: AgentRunResult) {
      cancelled = result.cancelled
      status = result.cancelled ? 'cancelled' : 'done'
      emit()
    },
    onError(message: string) {
      const err = classifyError(message)
      ;(runState as { error?: unknown }).error = err
      status = 'error'
      emit()
    },
  }

  const runState: { id: string; sessionId: string; startedAt: number; error?: unknown } = {
    id: runId,
    sessionId: opts.sessionId,
    startedAt,
  }

  const buildRun = (): ChatRun => ({
    id: runId,
    sessionId: opts.sessionId,
    status,
    startedAt,
    finishedAt: status === 'done' || status === 'cancelled' || status === 'error' ? Date.now() : undefined,
    streamingText,
    tools: [...tools],
    lastChangePlan,
    cancelled: status === 'cancelled' || cancelled,
    error: (runState as { error?: unknown }).error as ChatRun['error'],
  })

  let loopCancel: { fn?: () => void } = {}

  emit()

  return {
    get run() {
      return buildRun()
    },
    cancel() {
      cancelled = true
      status = 'cancelled'
      loopCancel.fn?.()
      emit()
    },
    snapshot() {
      return snapshotBeforeFirstTool
    },
  }
}

/**
 * Decode a `ChatChangePlan` out of a tool display payload.
 */
function extractChangePlan(display: unknown): ChatChangePlan | undefined {
  if (!display || typeof display !== 'object') return undefined
  const d = display as { kind?: unknown; plan?: unknown; text?: unknown }
  if (d.kind === 'change-plan' && d.plan && typeof d.plan === 'object') {
    return d.plan as ChatChangePlan
  }
  if (d.kind === 'text' && typeof d.text === 'string') {
    try {
      const parsed = JSON.parse(d.text) as { plan?: unknown }
      if (parsed.plan && typeof parsed.plan === 'object') {
        return parsed.plan as ChatChangePlan
      }
    } catch {
      /* not JSON, ignore */
    }
  }
  return undefined
}

export interface RunBridge {
  emitToolStart(call: AgentToolCall): void
  emitToolExecuted(call: AgentToolCall, execution: { output: string; isError?: boolean; display?: unknown }): void
  emitText(delta: string, cumulative: string): void
  emitDone(result: AgentRunResult): void
  emitError(message: string): void
}
