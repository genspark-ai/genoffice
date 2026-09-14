/**
 * `useChatRuntime` — the React-side hook that mounts a `ChatRuntime`
 * into the existing AiPanel render tree.
 *
 * It is intentionally thin: every `busy`/`phase`/`streaming`/`snapshot`
 * value the four panels used to maintain locally is now *derived* from
 * the runtime state, so the renderers never hold their own copy.
 */

import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'

import { ChatRuntime } from './runtime.js'
import type {
  ChatChangePlan,
  ChatEvent,
  ChatRuntimeOptions,
  ChatRun,
  ChatSendOptions,
  ChatSession,
} from './types.js'

export interface UseChatRuntimeResult {
  runtime: ChatRuntime
  session: ChatSession
  currentRun: ChatRun | null
  /** True while a run is in progress (running / streaming / awaiting-confirm). */
  busy: boolean
  /** Live phase for UI badges ('idle' | 'running' | 'streaming' | 'done' | 'error' | 'cancelled'). */
  phase: ChatRun['status']
  /** Last error encountered (normalised AIError or null). */
  lastError: import('./errors.js').AIError | null
  /** Most recent change plan emitted (still pending apply). */
  lastChangePlan: ChatChangePlan | null
  send(opts: ChatSendOptions): Promise<ChatRun>
  cancel(): void
  applyChangePlan(plan: ChatChangePlan): void
  rejectChangePlan(plan: ChatChangePlan): void
  snapshot(): import('./types.js').ChatSnapshot | null
}

export function useChatRuntime(options: ChatRuntimeOptions): UseChatRuntimeResult {
  // Keep options in a ref so callers can change them without re-mounting.
  const optionsRef = useRef(options)
  optionsRef.current = options

  const runtime = useMemo(() => ChatRuntime.create(options), [options.app, options.sessionId, options.skill, options.persistence])
  const [, force] = useState(0)
  const [currentRun, setCurrentRun] = useState<ChatRun | null>(null)
  const [lastChangePlan, setLastChangePlan] = useState<ChatChangePlan | null>(null)
  const [lastError, setLastError] = useState<import('./errors.js').AIError | null>(null)
  const [session, setSession] = useState<ChatSession>(() => runtime.getSession())

  useEffect(() => {
    let cancelled = false
    void runtime.init()
    const unsubscribe = runtime.subscribe((event: ChatEvent) => {
      if (cancelled) return
      if (event.type === 'session-updated') {
        setSession(event.session)
      }
      if (event.type === 'run-start' || event.type === 'run-finish' || event.type === 'run-error') {
        setCurrentRun(event.run)
        if (event.run.error) setLastError(event.run.error)
      }
      if (event.type === 'run-change-plan') {
        setLastChangePlan(event.plan)
      }
      force(n => n + 1)
    })
    return () => {
      cancelled = true
      unsubscribe()
    }
  }, [runtime])

  const send = useCallback(
    (opts: ChatSendOptions) => {
      setLastChangePlan(null)
      setLastError(null)
      return runtime.send(opts)
    },
    [runtime],
  )

  const cancel = useCallback(() => runtime.cancel(), [runtime])
  const applyChangePlan = useCallback((plan: ChatChangePlan) => {
    runtime.applyChangePlan(plan)
    setLastChangePlan(null)
  }, [runtime])
  const rejectChangePlan = useCallback((plan: ChatChangePlan) => {
    runtime.rejectChangePlan(plan)
    setLastChangePlan(null)
  }, [runtime])
  const snapshot = useCallback(() => runtime.snapshot(), [runtime])

  const phase = currentRun?.status ?? 'idle'
  const busy =
    phase === 'queued' || phase === 'running' || phase === 'streaming' || phase === 'awaiting-confirm'

  return {
    runtime,
    session,
    currentRun,
    busy,
    phase,
    lastError,
    lastChangePlan,
    send,
    cancel,
    applyChangePlan,
    rejectChangePlan,
    snapshot,
  }
}

// `useSyncExternalStore` is exported here so apps that want a stricter
// subscribe path can swap the internal useState-based mirror above.
export { useSyncExternalStore }
