/**
 * Opt-in model-turn trace for building fine-tuning corpora. With
 * GENOFFICE_AI_TRACE_DIR set, every ai:stream request appends one JSON line
 * (full request as sent to the provider, full response as streamed back) to
 * <dir>/ai-trace-<yyyymmdd>-<pid>.jsonl. Off by default; never written from the
 * renderer; the provider config (API keys) is not part of the record.
 */
import { appendFileSync, mkdirSync } from 'node:fs'
import { join } from 'node:path'
import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import type { AiTokenUsage } from './types'

export interface AiTraceRequest {
  requestId: string
  sessionId?: string | undefined
  /** which window/app issued it (e.g. the sender's URL) */
  source?: string | undefined
  provider: string
  model: string
  maxTokens: number
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
}

export interface AiTraceRecorder {
  onDelta(text: string): void
  onReasoningDelta(text: string): void
  onToolCall(call: AgentToolCall): void
  onStopReason(reason: string): void
  onUsage(usage: AiTokenUsage): void
  /** writes the line; `error` for a failed turn, 'aborted' when the user stopped it */
  finish(error?: string): void
}

export function aiTraceDir(env: NodeJS.ProcessEnv = process.env): string | null {
  const dir = env.GENOFFICE_AI_TRACE_DIR?.trim()
  return dir ? dir : null
}

function traceFile(dir: string, now: Date): string {
  const d = now.toISOString().slice(0, 10).replace(/-/g, '')
  return join(dir, `ai-trace-${d}-${process.pid}.jsonl`)
}

export function startAiTrace(
  request: AiTraceRequest,
  env: NodeJS.ProcessEnv = process.env,
): AiTraceRecorder | null {
  const dir = aiTraceDir(env)
  if (!dir) return null
  const startedAt = new Date()
  let text = ''
  let reasoning = ''
  const toolCalls: AgentToolCall[] = []
  let stopReason: string | undefined
  let usage: AiTokenUsage | undefined
  let done = false
  return {
    onDelta: (t) => {
      text += t
    },
    onReasoningDelta: (t) => {
      reasoning += t
    },
    onToolCall: (call) => {
      toolCalls.push(call)
    },
    onStopReason: (reason) => {
      stopReason = reason
    },
    onUsage: (u) => {
      usage = u
    },
    finish: (error) => {
      if (done) return
      done = true
      const line = JSON.stringify({
        ts: startedAt.toISOString(),
        durationMs: Date.now() - startedAt.getTime(),
        ...request,
        response: { text, reasoning, toolCalls, stopReason, usage, error },
      })
      try {
        mkdirSync(dir, { recursive: true })
        appendFileSync(traceFile(dir, startedAt), line + '\n')
      } catch (err) {
        console.warn('[ai-trace] write failed:', err instanceof Error ? err.message : err)
      }
    },
  }
}
