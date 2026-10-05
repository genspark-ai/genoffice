import { useEffect, useRef, useState } from 'react'

/**
 * The chat message queue behind "Enter queues while a reply is running"
 * (MiniMax-style): messages typed while the panel is busy wait in a strip
 * above the composer, and a pump starts the next one the moment the current
 * run settles — finished, failed, or stopped by the user. Nothing here knows
 * how a panel runs its loop: it observes `busy` transitions and calls
 * `submit`, which every panel maps onto its own runWith/send.
 *
 * Queue semantics:
 * - settling a run always pumps the next message, including after a stop or
 *   an error (a failed bubble keeps its retry action) — stopping everything
 *   for good is the strip's clear button;
 * - a `resetKey` change (new chat) wipes the queue; a panel that remounts per
 *   document loses it with the unmount instead;
 * - the queue lives in memory only, same footing as the composer draft.
 */
export interface QueuedChatMessage<M> {
  readonly id: string
  readonly text: string
  /** panel-defined snapshot taken when the message was queued (attachments, scope, …) */
  readonly meta: M
}

export interface ChatRunQueue<M> {
  readonly queued: readonly QueuedChatMessage<M>[]
  /** queue a typed draft; called while the panel is busy (the composer routes Enter here) */
  enqueue: (text: string, meta: M) => void
  /** edit a queued message's text in place (attachments/scope stay as queued) */
  update: (id: string, text: string) => void
  remove: (id: string) => void
  clear: () => void
}

export function useChatRunQueue<M>(args: {
  busy: boolean
  /**
   * start a run for the dequeued message; return false to keep it queued when
   * the panel is momentarily busy some other way (e.g. a post-run QC pass) —
   * the pump retries while the queue is non-empty and the panel is idle
   */
  submit: (text: string, meta: M) => boolean
  resetKey?: string | number
}): ChatRunQueue<M> {
  const { busy, submit, resetKey } = args
  const [queued, setQueued] = useState<QueuedChatMessage<M>[]>([])
  // the pump reads the queue synchronously (clear and pump can land in one commit)
  const queuedRef = useRef(queued)
  const busyRef = useRef(busy)
  const submitRef = useRef(submit)
  const nextIdRef = useRef(0)

  const setAll = (next: QueuedChatMessage<M>[]) => {
    queuedRef.current = next
    setQueued(next)
  }

  const tryPump = () => {
    if (busyRef.current) return
    const next = queuedRef.current[0]
    if (!next) return
    if (submitRef.current(next.text, next.meta)) setAll(queuedRef.current.slice(1))
  }
  const tryPumpRef = useRef(tryPump)

  useEffect(() => {
    submitRef.current = submit
  }, [submit])

  useEffect(() => {
    const was = busyRef.current
    busyRef.current = busy
    // a settle: whatever ended the run (finish, error, stop), the queue continues
    if (was && !busy) tryPumpRef.current()
  }, [busy])

  useEffect(() => {
    // only when a queued message is waiting and the panel is idle: covers
    // panels whose run start was refused (a QC pass, a stale guard) and will
    // not see another busy transition on its own
    if (busy || queued.length === 0) return
    const timer = window.setInterval(() => tryPumpRef.current(), 400)
    return () => window.clearInterval(timer)
  }, [busy, queued.length])

  useEffect(() => {
    if (resetKey === undefined) return
    queuedRef.current = []
    setQueued([])
  }, [resetKey])

  return {
    queued,
    enqueue: (text, meta) => {
      const trimmed = text.trim()
      if (!trimmed) return
      const id = `q${++nextIdRef.current}`
      setAll([...queuedRef.current, { id, text: trimmed, meta }])
      // a race can leave the panel idle at enqueue time (busy flipped between
      // render and keypress): pump straight away instead of stranding the message
      tryPump()
    },
    update: (id, text) => {
      const trimmed = text.trim()
      if (!trimmed) return
      setAll(queuedRef.current.map((m) => (m.id === id ? { ...m, text: trimmed } : m)))
    },
    remove: (id) => {
      setAll(queuedRef.current.filter((m) => m.id !== id))
    },
    clear: () => {
      setAll([])
    },
  }
}
