import { useCallback, useEffect, useRef, useState } from 'react'
import { NarrationRecorder } from './recorder'
import type { NarrationMark } from './segments'

/** A whole recorded show: the audio and where each slide starts in it. */
export interface NarrationTake {
  audio: Blob
  marks: NarrationMark[]
  endMs: number
}

export type RecordingPhase = 'starting' | 'recording' | 'paused' | 'failed'

/**
 * Records the microphone while `active`, marking each slide `slideIndex` enters (the end screen,
 * `ended`, belongs to no slide). `finish` ends the recording and hands it over once.
 */
export function useNarrationRecording(
  active: boolean,
  slideIndex: number | undefined,
  ended: boolean,
) {
  const recorderRef = useRef<NarrationRecorder | null>(null)
  const marksRef = useRef<NarrationMark[]>([])
  const [phase, setPhase] = useState<RecordingPhase>('starting')
  const [, setTick] = useState(0)
  const markIndex = ended ? -1 : slideIndex
  const markIndexRef = useRef(markIndex)
  markIndexRef.current = markIndex

  useEffect(() => {
    if (!active) return
    let cancelled = false
    void (async () => {
      try {
        if (!(await window.slidesApi.microphoneAccess())) throw new Error('microphone refused')
        const recorder = await NarrationRecorder.start()
        if (cancelled) return recorder.dispose()
        recorderRef.current = recorder
        marksRef.current = [{ slideIndex: markIndexRef.current ?? 0, atMs: 0 }]
        setPhase('recording')
      } catch {
        if (!cancelled) setPhase('failed')
      }
    })()
    const tick = window.setInterval(() => setTick((n) => n + 1), 500)
    return () => {
      cancelled = true
      window.clearInterval(tick)
      recorderRef.current?.dispose()
      recorderRef.current = null
    }
  }, [active])

  useEffect(() => {
    const recorder = recorderRef.current
    if (!recorder || markIndex == null) return
    if (marksRef.current.at(-1)?.slideIndex === markIndex) return
    marksRef.current.push({ slideIndex: markIndex, atMs: recorder.elapsedMs() })
  }, [markIndex])

  const togglePause = useCallback(() => {
    const recorder = recorderRef.current
    if (!recorder) return
    if (recorder.paused) recorder.resume()
    else recorder.pause()
    setPhase(recorder.paused ? 'paused' : 'recording')
  }, [])

  const finish = useCallback(async (): Promise<NarrationTake | null> => {
    const recorder = recorderRef.current
    recorderRef.current = null
    if (!recorder) return null
    const { audio, endMs } = await recorder.stop()
    return { audio, marks: marksRef.current, endMs }
  }, [])

  const elapsedMs = recorderRef.current?.elapsedMs() ?? 0
  const slideStartMs = marksRef.current.at(-1)?.atMs ?? 0
  return { phase, elapsedMs, slideElapsedMs: elapsedMs - slideStartMs, togglePause, finish }
}
