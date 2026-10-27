import { useCallback, useEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'
import { anchorAt, discardAnchored, writeAnchored, type VoiceAnchor } from './voice-text-splice'

/** What a voice session is doing right now. */
export type VoicePhase = 'idle' | 'requesting' | 'recording' | 'transcribing'

/** Why a session ended without a transcript. */
export type VoiceError = 'denied' | 'unavailable' | 'failed'

/**
 * Turns captured audio into a transcript. `onPartial` may be called any number of
 * times with a better interim transcript before the promise settles on the final one;
 * a transcriber that never calls it degrades to final-only, which is correct.
 *
 * Audio is handed over in memory and never persisted: see
 * `docs/adr/0002-audio-never-leaves-the-device.md`.
 */
export type VoiceTranscriber = (
  audio: Blob,
  options: { readonly signal: AbortSignal; readonly onPartial?: (text: string) => void },
) => Promise<string>

export interface UseVoiceInputOptions {
  /** The composer draft the session writes into. */
  readonly draft: string
  readonly onChange: (next: string) => void
  readonly textareaRef: RefObject<HTMLTextAreaElement | null>
  readonly transcribe: VoiceTranscriber
  /** Ask the OS for the microphone; macOS gates it. Resolves false when refused. */
  readonly requestMicrophoneAccess: () => Promise<boolean>
}

export interface VoiceInput {
  readonly phase: VoicePhase
  readonly error: VoiceError | null
  /** False when the runtime cannot capture audio at all; the button hides itself. */
  readonly supported: boolean
  /** The microphone button's whole behaviour: start listening, or stop and settle. */
  readonly toggle: () => void
  /** End the session and take the interim transcript back out of the draft. */
  readonly discard: () => void
  readonly clearError: () => void
}

/** A forgotten open microphone is worse than a truncated sentence. */
const MAX_SESSION_MS = 120_000
/** How many of our own writes to accept as "not a user edit" while renders settle. */
const WRITTEN_MEMORY = 4

interface Session {
  readonly stream: MediaStream
  readonly recorder: MediaRecorder
  readonly chunks: Blob[]
  anchor: VoiceAnchor | null
  /** Values this session wrote, so a lagging render is not mistaken for typing. */
  readonly written: string[]
  /** The session is over: any transcript still in flight must not touch the draft. */
  ended: boolean
  timer: ReturnType<typeof setTimeout> | null
  abort: AbortController | null
}

function stopStream(stream: MediaStream): void {
  for (const track of stream.getTracks()) track.stop()
}

/**
 * The draft this session is responsible for: what it last wrote, not the prop, so a
 * render that has not committed yet cannot make the session write to a stale string.
 */
function sessionDraft(session: Session, fallback: string): string {
  return session.written[session.written.length - 1] ?? fallback
}

/**
 * The voice session state machine behind the composer microphone. It owns the
 * microphone stream, the interim transcript's span of the draft, and nothing else:
 * transcription itself is injected, so the engine can change without touching this.
 *
 * If the user types during a session the session ends and keeps the interim
 * transcript, rather than fighting the caret for ownership of the draft.
 */
export function useVoiceInput({
  draft,
  onChange,
  textareaRef,
  transcribe,
  requestMicrophoneAccess,
}: UseVoiceInputOptions): VoiceInput {
  const [phase, setPhase] = useState<VoicePhase>('idle')
  const [error, setError] = useState<VoiceError | null>(null)
  const sessionRef = useRef<Session | null>(null)
  const phaseRef = useRef<VoicePhase>('idle')
  /** Bumped on every start, so a slow getUserMedia cannot open a stale session. */
  const startToken = useRef(0)
  const latest = useRef({ draft, onChange, transcribe, requestMicrophoneAccess })
  const mounted = useRef(true)

  latest.current = { draft, onChange, transcribe, requestMicrophoneAccess }

  const supported =
    typeof navigator !== 'undefined' && typeof navigator.mediaDevices?.getUserMedia === 'function'

  const move = useCallback((next: VoicePhase) => {
    phaseRef.current = next
    if (mounted.current) setPhase(next)
  }, [])

  /** Write the final transcript and leave the caret after it. */
  const commit = useCallback(
    (session: Session, text: string) => {
      const value = sessionDraft(session, latest.current.draft)
      const next = session.anchor ? writeAnchored(value, session.anchor, text) : null
      if (next) {
        latest.current.onChange(next.value)
        session.anchor = next.anchor
        session.written.push(next.value)
        const caret = next.anchor.end
        requestAnimationFrame(() => {
          const textarea = textareaRef.current
          if (!textarea) return
          textarea.focus()
          textarea.setSelectionRange(caret, caret)
        })
      }
      stopStream(session.stream)
      sessionRef.current = null
      move('idle')
    },
    [move, textareaRef],
  )

  const fail = useCallback(
    (reason: VoiceError) => {
      const session = sessionRef.current
      if (session) {
        stopStream(session.stream)
        sessionRef.current = null
      }
      if (mounted.current) setError(reason)
      move('idle')
    },
    [move],
  )

  const settle = useCallback(
    async (session: Session) => {
      stopStream(session.stream)
      if (session.timer) clearTimeout(session.timer)
      session.timer = null

      if (session.ended || sessionRef.current !== session) return

      const audio = new Blob(session.chunks, { type: session.recorder.mimeType || 'audio/webm' })
      if (audio.size === 0) {
        fail('failed')
        return
      }

      const abort = new AbortController()
      session.abort = abort
      move('transcribing')

      try {
        const text = await latest.current.transcribe(audio, {
          signal: abort.signal,
          onPartial: (partial) => {
            // A session that lost its span, or one the user typed into, keeps what it has.
            if (session.anchor === null || sessionRef.current !== session) return
            const value = latest.current.draft
            if (!session.written.includes(value)) {
              // The user took over the draft: end here and keep the interim transcript.
              session.anchor = null
              session.ended = true
              stopStream(session.stream)
              sessionRef.current = null
              move('idle')
              return
            }
            const written = writeAnchored(value, session.anchor, partial)
            session.anchor = written.anchor
            session.written.push(written.value)
            if (session.written.length > WRITTEN_MEMORY) session.written.shift()
            latest.current.onChange(written.value)
          },
        })
        if (session.ended || sessionRef.current !== session) return
        // The user typed while the final transcript was in flight: their draft wins.
        // Without this the final is spliced onto the last value the session wrote,
        // which silently drops whatever was typed in between.
        if (!session.written.includes(latest.current.draft)) {
          stopStream(session.stream)
          sessionRef.current = null
          move('idle')
          return
        }
        const final = text.trim()
        if (!final) {
          if (session.anchor) {
            const value = sessionDraft(session, latest.current.draft)
            latest.current.onChange(discardAnchored(value, session.anchor))
          }
          stopStream(session.stream)
          sessionRef.current = null
          move('idle')
          return
        }
        commit(session, final)
      } catch {
        if (session.ended || sessionRef.current !== session) return
        fail('failed')
      }
    },
    [commit, fail, move],
  )

  const start = useCallback(async () => {
    if (phaseRef.current !== 'idle') return
    if (!supported) {
      fail('unavailable')
      return
    }
    move('requesting')
    setError(null)
    const token = (startToken.current += 1)
    /** Held outside the try so a recorder that cannot be built does not leak the mic. */
    let opened: MediaStream | null = null

    try {
      if (!(await latest.current.requestMicrophoneAccess())) {
        fail('denied')
        return
      }
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: { echoCancellation: true, noiseSuppression: true },
      })
      opened = stream
      if (!mounted.current || startToken.current !== token) {
        stopStream(stream)
        return
      }

      const textarea = textareaRef.current
      const value = latest.current.draft
      const anchor = textarea
        ? anchorAt(value, textarea.selectionStart, textarea.selectionEnd)
        : anchorAt(value, value.length)

      const recorder = new MediaRecorder(stream)
      const session: Session = {
        stream,
        recorder,
        chunks: [],
        anchor,
        written: [value],
        ended: false,
        timer: null,
        abort: null,
      }
      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) session.chunks.push(event.data)
      }
      recorder.onstop = () => void settle(session)
      // A timeslice keeps what was said if the session ends abruptly.
      recorder.start(1000)

      session.timer = setTimeout(() => {
        if (recorder.state !== 'inactive') recorder.stop()
      }, MAX_SESSION_MS)

      sessionRef.current = session
      move('recording')
    } catch {
      // `sessionRef` is only set once the session is whole, so `fail` cannot close a
      // stream that was opened before the recorder existed.
      if (opened) stopStream(opened)
      fail('unavailable')
    }
  }, [fail, move, settle, supported, textareaRef])

  const stop = useCallback(() => {
    const session = sessionRef.current
    if (!session) return
    if (session.recorder.state === 'inactive') return
    session.recorder.stop()
  }, [])

  const discard = useCallback(() => {
    const session = sessionRef.current
    if (!session) return
    session.ended = true
    session.abort?.abort()
    if (session.anchor) {
      const value = sessionDraft(session, latest.current.draft)
      latest.current.onChange(discardAnchored(value, session.anchor))
    }
    if (session.recorder.state !== 'inactive') {
      // `onstop` runs `settle`, which bails out on `ended`.
      session.recorder.stop()
    } else {
      stopStream(session.stream)
      sessionRef.current = null
    }
    move('idle')
  }, [move])

  const toggle = useCallback(() => {
    if (phaseRef.current === 'recording') stop()
    else if (phaseRef.current === 'idle') void start()
  }, [start, stop])

  const clearError = useCallback(() => setError(null), [])

  useEffect(() => {
    mounted.current = true
    return () => {
      mounted.current = false
      const session = sessionRef.current
      if (!session) return
      session.ended = true
      session.abort?.abort()
      if (session.timer) clearTimeout(session.timer)
      if (session.recorder.state !== 'inactive') session.recorder.stop()
      stopStream(session.stream)
      sessionRef.current = null
    }
  }, [])

  return { phase, error, supported, toggle, discard, clearError }
}
