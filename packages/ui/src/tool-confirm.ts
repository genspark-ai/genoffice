import { useCallback, useEffect, useRef, useState } from 'react'

/** approval requests wait this long before silence counts as denial */
export const TOOL_CONFIRM_TIMEOUT_MS = 5 * 60 * 1000

/** what one confirmation request looks like to the UI */
export interface ToolConfirmRequest {
  readonly name: string
  readonly input: Record<string, unknown>
}

export interface ToolConfirmCard {
  readonly toolName: string
  /** short JSON preview of the call's arguments */
  readonly args: string
  /** awaiting: buttons live; allowed/denied: the outcome note stays in the flow */
  readonly outcome: 'awaiting' | 'allowed' | 'denied'
  /** true when the denial came from the timeout rather than a click */
  readonly timedOut: boolean
}

export interface UseToolConfirm {
  /** the card the panel renders inline in the conversation (null when idle) */
  readonly card: ToolConfirmCard | null
  /** the confirmTool hook: parks the request, resolves when decided (or timed out) */
  readonly confirm: (call: ToolConfirmRequest) => Promise<boolean>
  readonly allow: () => void
  readonly deny: () => void
  /**
   * Settle an outstanding request without a click — Stop, unmount, thread
   * switch. `allowed` false is the safe default: an unanswered dangerous
   * action must not run.
   */
  readonly settleAll: (allowed: boolean) => void
  /** drop the settled card from the flow (a new run owns the stream now) */
  readonly clear: () => void
}

/** longest arguments preview kept on the card */
const ARGS_PREVIEW_MAX = 400

function argsPreview(input: Record<string, unknown>): string {
  let text: string
  try {
    text = JSON.stringify(input) ?? '{}'
  } catch {
    return '{}'
  }
  return text.length > ARGS_PREVIEW_MAX ? `${text.slice(0, ARGS_PREVIEW_MAX)}…` : text
}

/**
 * The host side of the agent loop's confirmTool gate: one approval request at
 * a time (the loop serializes high-risk calls, so there is never a queue), a
 * card state the panel renders inline, and a timeout after which silence
 * counts as denial. A second request while one is pending settles the first
 * as denied — defensively, since the loop never issues one.
 */
export function useToolConfirm(
  options: {
    /** default 5 minutes; silence past it denies the call */
    timeoutMs?: number
  } = {},
): UseToolConfirm {
  const timeoutMs = options.timeoutMs ?? TOOL_CONFIRM_TIMEOUT_MS
  const [card, setCard] = useState<ToolConfirmCard | null>(null)
  const resolverRef = useRef<((allowed: boolean) => void) | null>(null)
  const timerRef = useRef<number | null>(null)

  const settle = useCallback((allowed: boolean, timedOut = false) => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current)
      timerRef.current = null
    }
    const resolve = resolverRef.current
    resolverRef.current = null
    // the card stays in the flow with its outcome, so the decision is part of
    // the visible transcript until the next run clears it
    setCard((prev) =>
      prev
        ? {
            ...prev,
            outcome: allowed ? 'allowed' : 'denied',
            ...(timedOut ? { timedOut: true } : {}),
          }
        : prev,
    )
    resolve?.(allowed)
  }, [])

  const confirm = useCallback(
    (call: ToolConfirmRequest): Promise<boolean> => {
      // a request left over from a run that never settled: deny it before
      // showing the new one, so exactly one dialog is ever answerable
      if (resolverRef.current) settle(false)
      return new Promise<boolean>((resolve) => {
        resolverRef.current = resolve
        setCard({
          toolName: call.name,
          args: argsPreview(call.input),
          outcome: 'awaiting',
          timedOut: false,
        })
        timerRef.current = window.setTimeout(() => {
          timerRef.current = null
          settle(false, true)
        }, timeoutMs)
      })
    },
    [settle, timeoutMs],
  )

  // unmount must not strand the loop's await: deny whatever is pending
  useEffect(
    () => () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current)
      resolverRef.current?.(false)
      resolverRef.current = null
    },
    [],
  )

  // stable identities: hosts hand these to memoized message rows, and a
  // fresh closure per render would invalidate every row on every parent
  // render (the streaming path re-renders the panel per frame)
  const allow = useCallback(() => settle(true), [settle])
  const deny = useCallback(() => settle(false), [settle])
  const settleAll = useCallback(
    (allowed: boolean) => {
      if (resolverRef.current) settle(allowed)
    },
    [settle],
  )
  const clear = useCallback(() => {
    if (resolverRef.current) settle(false)
    setCard(null)
  }, [settle])

  return {
    card,
    confirm,
    allow,
    deny,
    settleAll,
    clear,
  }
}
