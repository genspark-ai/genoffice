import type { ChatMessage, RunSummary } from '@genoffice/project-store'

/**
 * Run-record persistence for the home ask panel, over the project-store chat
 * channel. The panel's loop lives in this renderer, so a run dies with the
 * window: what survives is the record pair written around it — a start line
 * when the run begins, a terminal line when it settles. A start with no
 * terminal reads back after a restart as "interrupted", which is exactly what
 * happened.
 *
 * The records live in their own file-less chat under a reserved id, so they
 * reuse the store's JSONL append path (its caps, its tolerant reader) without
 * touching the panel's own transcript, and the thread list filters them out
 * (panelThreads matches kb-ask ids only). The same RunEvent shapes are what
 * ProjectStore's first-class runs API (listRuns/appendRunEvent) writes, so
 * switching this bridge to dedicated IPC later is a transport change, not a
 * format change.
 */

/** reserved chat id for run records; deliberately not a kb-ask-* id */
export const RUNS_CHAT_ID = 'kb-runs'

/** how many stored lines the panel reads back on mount (covers the recent past) */
export const RUNS_LOAD_LIMIT = 500

/** one run-record line, version-tagged so ordinary text is never mistaken for one */
export interface RunRecord {
  v: 1
  runId: string
  kind: 'start' | 'finish' | 'error'
  instruction?: string
  chatId?: string
  outcome?: 'done' | 'cancelled' | 'discarded'
  error?: string
}

/** instruction length cap, kept in step with the store's run-record cap */
const INSTRUCTION_MAX = 200
const ERROR_MAX = 500

function record(rec: RunRecord): string {
  return JSON.stringify(rec)
}

export function runStartRecord(runId: string, instruction: string, chatId: string): string {
  return record({
    v: 1,
    runId,
    kind: 'start',
    instruction: instruction.slice(0, INSTRUCTION_MAX),
    chatId,
  })
}

export function runFinishRecord(runId: string, outcome: 'done' | 'cancelled'): string {
  return record({ v: 1, runId, kind: 'finish', outcome })
}

/** the clear action on an interrupted row: the record is settled, not deleted */
export function runDiscardRecord(runId: string): string {
  return record({ v: 1, runId, kind: 'finish', outcome: 'discarded' })
}

export function runErrorRecord(runId: string, error: string): string {
  return record({ v: 1, runId, kind: 'error', error: error.slice(0, ERROR_MAX) })
}

function isRunRecord(value: unknown): value is RunRecord {
  if (!value || typeof value !== 'object') return false
  const rec = value as Partial<RunRecord>
  return (
    rec.v === 1 &&
    typeof rec.runId === 'string' &&
    rec.runId.length > 0 &&
    (rec.kind === 'start' || rec.kind === 'finish' || rec.kind === 'error')
  )
}

/**
 * Fold stored chat lines back into run summaries, mirroring the store's own
 * fold: the start line fixes instruction/chat/startedAt; the last terminal
 * line fixes the status; a run with no terminal is interrupted. Unparseable
 * lines and non-record turns are skipped, never fatal.
 */
export function parseRunRecords(messages: readonly ChatMessage[]): RunSummary[] {
  const runs = new Map<string, RunSummary>()
  for (const message of messages) {
    let parsed: unknown
    try {
      parsed = JSON.parse(message.text)
    } catch {
      continue
    }
    if (!isRunRecord(parsed)) continue
    const existing = runs.get(parsed.runId)
    if (parsed.kind === 'start') {
      runs.set(parsed.runId, {
        runId: parsed.runId,
        ...(parsed.chatId ? { chatId: parsed.chatId } : {}),
        instruction: parsed.instruction ?? '',
        status: 'interrupted',
        startedAt: message.ts,
      })
    } else if (existing) {
      const terminal: RunSummary = {
        ...existing,
        endedAt: message.ts,
        status:
          parsed.kind === 'error'
            ? 'error'
            : parsed.outcome === 'cancelled'
              ? 'cancelled'
              : parsed.outcome === 'discarded'
                ? 'discarded'
                : 'done',
        ...(parsed.kind === 'error' && parsed.error ? { error: parsed.error } : {}),
      }
      runs.set(parsed.runId, terminal)
    }
  }
  return [...runs.values()].sort((a, b) =>
    a.startedAt < b.startedAt ? 1 : a.startedAt > b.startedAt ? -1 : a.runId.localeCompare(b.runId),
  )
}

/** the rows the background-task list shows from storage: what never finished */
export function interruptedOf(summaries: readonly RunSummary[]): RunSummary[] {
  return summaries.filter((run) => run.status === 'interrupted')
}
